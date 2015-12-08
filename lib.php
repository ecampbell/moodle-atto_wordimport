<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Atto text editor import Microsoft Word files.
 *
 * @package    atto_wordimport
 * @copyright  2015 Eoin Campbell
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

defined('MOODLE_INTERNAL') || die();

require_once($CFG->libdir . '/filestorage/file_storage.php');
require_once($CFG->dirroot . '/repository/lib.php');
require_once("$CFG->libdir/xmlize.php");


/**
 * Initialise the strings required for js
 *
 * @return void
 */
function atto_wordimport_strings_for_js() {
    global $PAGE;

    $strings = array(
        'uploading',
        'transformationfailed',
        'fileuploadfailed',
        'fileconversionfailed',
        'pluginname'
    );

    $PAGE->requires->strings_for_js($strings, 'atto_wordimport');
}


/**
 * Extract the WordProcessingML XML files from the .docx file, and use a sequence of XSLT
 * steps to convert it into XHTML
 *
 * @param string $filename name of file uploaded to file repository as a draft
 * @param int $usercontextid ID of draft file area where images should be stored
 * @param int $draftitemid ID of particular group in draft file area where images should be stored
 * @return string XHTML content extracted from Word file
 */
function atto_wordimport_convert_to_xhtml($filename, $usercontextid, $draftitemid) {
    global $CFG, $USER;

    $word2xmlstylesheet1 = __DIR__ . "/wordml2xhtmlpass1.xsl"; // Convert WordML into basic XHTML.
    $word2xmlstylesheet2 = __DIR__ . "/wordml2xhtmlpass2.xsl"; // Refine basic XHTML into Word-compatible XHTML.

    // Check that we can unzip the Word .docx file into its component files.
    $zipres = zip_open($filename);
    if (!is_resource($zipres)) {
        // Cannot unzip file.
        atto_wordimport_debug_unlink($filename);
        throw new moodle_exception('cannotunzipfile', 'error');
    }

    // Check that XSLT is installed.
    if (!class_exists('XSLTProcessor') || !function_exists('xslt_create')) {
        // PHP extension 'xsl' is required for this action.
        throw new moodle_exception(get_string('extensionrequired', 'tool_xmldb', 'xsl'));
    }

    // Give XSLT as much memory as possible, to enable larger Word files to be imported.
    raise_memory_limit(MEMORY_HUGE);

    if (!file_exists($word2xmlstylesheet1)) {
        // XSLT stylesheet to transform WordML into XHTML is missing.
        throw new moodle_exception('filemissing', 'moodle', $word2xmlstylesheet1);
    }

    // Set common parameters for all XSLT transformations.
    $parameters = array (
        'moodle_language' => current_language(),
        'moodle_textdirection' => (right_to_left()) ? 'rtl' : 'ltr',
        'moodle_release' => $CFG->release,
        'moodle_url' => $CFG->wwwroot . "/",
        'heading1stylelevel' => get_config('atto_wordimport', 'heading1stylelevel'),
        'pluginname' => 'atto_wordimport', // Include plugin name to control image data handling.
        'debug_flag' => DEBUG_WORDIMPORT
    );

    // Pre-XSLT preparation: merge the WordML and image content from the .docx Word file into one large XML file.
    // Initialise an XML string to use as a wrapper around all the XML files.
    $xmldeclaration = '<?xml version="1.0" encoding="UTF-8"?>';
    $wordmldata = $xmldeclaration . "\n<pass1Container>\n";
    $imagestring = "";

    $xmldeclregexp = '/<\?xml version="1.0" ([^>]*)>/';

    $fs = get_file_storage();
    // Prepare filerecord array for creating each new image file.
    $fileinfo = array(
        'contextid' => $usercontextid,
        'component' => 'user',
        'filearea' => 'draft',
        'userid' => $USER->id,
        'itemid' => $draftitemid,
        'filepath' => '/',
        'filename' => ''
        );

    $zipentry = zip_read($zipres);
    while ($zipentry) {
        if (!zip_entry_open($zipres, $zipentry, "r")) {
            // Can't read the XML file from the Word .docx file.
            zip_close($zipres);
            throw new moodle_exception('errorunzippingfiles', 'error');
        }

        $zefilename = zip_entry_name($zipentry);
        $zefilesize = zip_entry_filesize($zipentry);

        // Insert internal images into the files table.
        if (strpos($zefilename, "media")) {
            // @codingStandardsIgnoreLine $imageformat = substr($zefilename, strrpos($zefilename, ".") + 1);
            $imagedata = zip_entry_read($zipentry, $zefilesize);
            $imagename = basename($zefilename);
            $imagesuffix = strtolower(substr(strrchr($zefilename, "."), 1));
            // GIF, PNG, JPG and JPEG handled OK, but bmp and other non-Internet formats are not.
            if ($imagesuffix == 'gif' or $imagesuffix == 'png' or $imagesuffix == 'jpg' or $imagesuffix == 'jpeg') {
                // Prepare the file details for storage, ensuring the image name is unique.
                $imagenameunique = $imagename;
                $file = $fs->get_file($usercontextid, 'user', 'draft', $draftitemid, '/', $imagenameunique);
                while ($file) {
                    $imagenameunique = basename($imagename, '.' . $imagesuffix) . '_' . substr(uniqid(), 8, 4) .
                        '.' . $imagesuffix;
                    $file = $fs->get_file($usercontextid, 'user', 'draft', $draftitemid, '/', $imagenameunique);
                }

                $fileinfo['filename'] = $imagenameunique;
                $fs->create_file_from_string($fileinfo, $imagedata);
                debugging(__FUNCTION__ . ":" . __LINE__ . ": stored \"{$imagename}\"" .
                    " as \"{$imagenameunique}\" with itemid {$draftitemid}", DEBUG_WORDIMPORT);

                $imageurl = "$CFG->wwwroot/draftfile.php/$usercontextid/user/draft/$draftitemid/$imagenameunique";
                // Return all the details of where the file is stored, even though we don't need them at the moment.
                $imagestring .= "<file filename=\"media/{$imagename}\"";
                $imagestring .= " contextid=\"{$usercontextid}\" itemid=\"{$draftitemid}\"";
                $imagestring .= " name=\"{$imagenameunique}\" url=\"{$imageurl}\">{$imageurl}</file>\n";
            } else {
                debugging(__FUNCTION__ . ":" . __LINE__ . ": ignore unsupported media file $zefilename" .
                    " = $imagename, imagesuffix = $imagesuffix", DEBUG_WORDIMPORT);
            }
        } else {
            // Look for required XML files, read and wrap it, remove the XML declaration, and add it to the XML string.
            switch ($zefilename) {
                case "word/document.xml":
                    $wordmldata .= "<wordmlContainer>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</wordmlContainer>\n";
                    break;
                case "docProps/core.xml":
                    $wordmldata .= "<dublinCore>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</dublinCore>\n";
                    break;
                case "docProps/custom.xml":
                    $wordmldata .= "<customProps>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</customProps>\n";
                    break;
                case "word/styles.xml":
                    $wordmldata .= "<styleMap>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</styleMap>\n";
                    break;
                case "word/_rels/document.xml.rels":
                    $wordmldata .= "<documentLinks>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</documentLinks>\n";
                    break;
                case "word/footnotes.xml":
                    $wordmldata .= "<footnotesContainer>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</footnotesContainer>\n";
                    break;
                case "word/_rels/footnotes.xml.rels":
                    $wordmldata .= "<footnoteLinks>" . preg_replace($xmldeclregexp,
                        zip_entry_read($zipentry, $zefilesize), "") . "</footnoteLinks>\n";
                    break;
                /* @codingStandardsIgnoreStart
                case "word/_rels/settings.xml.rels":
                    $wordmldata .= "<settingsLinks>" . preg_replace($xmldeclregexp, "",
                        zip_entry_read($zipentry, $zefilesize)) . "</settingsLinks>\n";
                    break;
                    @codingStandardsIgnoreEnd
                */
                default:
                    // @codingStandardsIgnoreLine debugging(__FUNCTION__ . ":" . __LINE__ . ": Ignore $zefilename", DEBUG_WORDIMPORT);
            }
        }
        // Get the next file in the Zip package.
        $zipentry = zip_read($zipres);
    }  // End while loop.
    zip_close($zipres);

    // Add images section and close the merged XML file.
    $wordmldata .= "<imagesContainer>\n" . $imagestring . "</imagesContainer>\n"  . "</pass1Container>";

    // Pass 1 - convert WordML into linear XHTML.
    // Create a temporary file to store the merged WordML XML content to transform.
    $tempwordmlfilename = $CFG->dataroot . '/temp/' . basename($filename, ".tmp") . ".wml";
    if (($nbytes = file_put_contents($tempwordmlfilename, $wordmldata)) == 0) {
        // Cannot save the file.
        throw new moodle_exception('cannotsavefile', 'error', $tempwordmlfilename);
    }

    $xsltproc = xslt_create();
    if (!($xsltoutput = xslt_process($xsltproc, $tempwordmlfilename, $word2xmlstylesheet1, null, null, $parameters))) {
        // Transformation failed.
        atto_wordimport_debug_unlink($tempwordmlfilename);
        throw new moodle_exception('transformationfailed', 'atto_wordimport', $tempwordmlfilename);
    }
    atto_wordimport_debug_unlink($tempwordmlfilename);
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 1 succeeded, XHTML output fragment = " .
        str_replace("\n", "", substr($xsltoutput, 0, 200)), DEBUG_WORDIMPORT);

    // Write output of Pass 1 to a temporary file, for use in Pass 2.
    $tempxhtmlfilename = $CFG->dataroot . '/temp/' . basename($filename, ".tmp") . ".if1";
    if (($nbytes = file_put_contents($tempxhtmlfilename, $xsltoutput )) == 0) {
        // Cannot save the file.
        throw new moodle_exception('cannotsavefile', 'error', $tempxhtmlfilename);
    }

    // Pass 2 - tidy up linear XHTML a bit.
    debugging(__FUNCTION__ . ":" . __LINE__ . ": XSLT Pass 2 using \"" . $word2xmlstylesheet2 . "\"", DEBUG_WORDIMPORT);
    if (!($xsltoutput = xslt_process($xsltproc, $tempxhtmlfilename, $word2xmlstylesheet2, null, null, $parameters))) {
        // Transformation failed.
        atto_wordimport_debug_unlink($tempxhtmlfilename);
        throw new moodle_exception('transformationfailed', 'atto_wordimport', $tempxhtmlfilename);
    }
    atto_wordimport_debug_unlink($tempxhtmlfilename);

    // Strip out superfluous namespace declarations on paragraph elements, which Moodle 2.7+ on Windows seems to throw in.
    $xsltoutput = str_replace('<p xmlns="http://www.w3.org/1999/xhtml"', '<p', $xsltoutput);
    $xsltoutput = str_replace(' xmlns=""', '', $xsltoutput);
    // Remove 'mml:' prefix from child MathML element and attributes for compatibility with MathJax.
    $xsltoutput = str_replace('<mml:', '<', $xsltoutput);
    $xsltoutput = str_replace('</mml:', '</', $xsltoutput);
    $xsltoutput = str_replace(' mathvariant="normal"', '', $xsltoutput);
    $xsltoutput = str_replace(' xmlns:mml="http://www.w3.org/1998/Math/MathML"', '', $xsltoutput);
    $xsltoutput = str_replace('<math>', '<math xmlns="http://www.w3.org/1998/Math/MathML">', $xsltoutput);

    // Keep the converted XHTML file for debugging if developer debugging enabled.
    if (debugging(null, DEBUG_WORDIMPORT)) {
        $tempxhtmlfilename = $CFG->dataroot . '/temp/' . basename($filename, ".tmp") . ".xhtml";
        file_put_contents($tempxhtmlfilename, $xsltoutput);
    }

    return $xsltoutput;
}   // End function convert_to_xhtml.


/**
 * Get the HTML body from the converted Word file
 *
 * @param string $xhtmlstring complete XHTML text including head element metadata
 * @return string XHTML text inside <body> element
 */
function atto_wordimport_get_html_body($xhtmlstring) {
    // @codingStandardsIgnoreLine debugging(__FUNCTION__ . "(xhtmlstring = \"" . substr($xhtmlstring, 0, 100) . "\")", DEBUG_WORDIMPORT);

    $bodystart = stripos($xhtmlstring, '<body>') + strlen('<body>');
    $bodylength = strripos($xhtmlstring, '</body>') - $bodystart;
    // @codingStandardsIgnoreLine debugging(__FUNCTION__ . ":" . __LINE__ . ": bodystart = {$bodystart}, bodylength = {$bodylength}", DEBUG_WORDIMPORT);
    if ($bodystart !== false || $bodylength !== false) {
        $xhtmlbody = substr($xhtmlstring, $bodystart, $bodylength);
    } else {
        // @codingStandardsIgnoreLine debugging(__FUNCTION__ . "() -> Invalid XHTML, using original cdata string", DEBUG_WORDIMPORT);
        $xhtmlbody = $xhtmlstring;
    }

    // @codingStandardsIgnoreLine debugging(__FUNCTION__ . "() -> |" . str_replace("\n", "", substr($xhtmlbody, 0, 100)) . " ...|", DEBUG_WORDIMPORT);
    return $xhtmlbody;
}

/**
 * Delete temporary files if debugging disabled
 *
 * @param string $filename name of file to be deleted
 * @return void
 */
function atto_wordimport_debug_unlink($filename) {
    if (DEBUG_WORDIMPORT == 0) {
        unlink($filename);
    }
}

