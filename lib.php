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


/**
 * Initialise this plugin
 * @param string $elementid
 */
function atto_wordimport_strings_for_js() {
    global $PAGE;

    $strings = array(
        'insert',
        'converting',
        'uploading',
        'enterflavor',
        'dialogtitle',
        'xmlnotsupported',
        'pluginname'
    );

    debugging(__FUNCTION__ . "()", DEBUG_DEVELOPER);
    $PAGE->requires->strings_for_js($strings, 'atto_wordimport');
}


/**
 * Extract the WordProcessingML XML files from the .docx file, and use a sequence of XSLT
 * steps to convert it into XHTML
 *
 * @param $filename name of file uploaded to file repository as a draft
 * @return string XHTML content extracted from Word file
 */
function convert_to_xhtml($filename, $contextid) {
    global $CFG, $OUTPUT;

    $word2mqxmlstylesheet1 = 'wordml2xhtml_pass1.xsl';      // Convert WordML into basic XHTML
    $word2mqxmlstylesheet2 = 'wordml2xhtml_pass2.xsl';      // Refine basic XHTML into Word-compatible XHTML

    debugging(__FUNCTION__ . ":" . __LINE__ . ": Word file = $filename", DEBUG_DEVELOPER);
    // Give XSLT as much memory as possible, to enable larger Word files to be imported.
    raise_memory_limit(MEMORY_HUGE);

    // XSLT stylesheet to convert WordML into initial XHTML format.
    $stylesheet = dirname(basename(__FILE__)) . "/" . $word2mqxmlstylesheet1;

    // Check that XSLT is installed, and the XSLT stylesheet is present.
    if (!class_exists('XSLTProcessor') || !function_exists('xslt_create')) {
        debugging(__FUNCTION__ . " (" . __LINE__ . "): XSLT not installed", DEBUG_DEVELOPER);
        //echo $OUTPUT->notification(get_string('xsltunavailable', 'atto_wordimport'));
        return false;
    } else if (!file_exists($stylesheet)) {
        // XSLT stylesheet to transform WordML into XHTML doesn't exist.
        debugging(__FUNCTION__ . " (" . __LINE__ . "): XSLT stylesheet missing: $stylesheet", DEBUG_DEVELOPER);
        //echo $OUTPUT->notification(get_string('stylesheetunavailable', 'atto_wordimport', $stylesheet));
        return false;
    }

    // Set common parameters for all XSLT transformations.
    $parameters = array (
        'moodle_language' => current_language(),
        'moodle_textdirection' => (right_to_left())? 'rtl': 'ltr',
        'moodle_release' => $CFG->release,
        'moodle_url' => $CFG->wwwroot . "/",
        'debug_flag' => debugging('', DEBUG_DEVELOPER)
    );

    // Pre-XSLT preparation: merge the WordML and image content from the .docx Word file into one large XML file.
    // Initialise an XML string to use as a wrapper around all the XML files.
    $xmldeclaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $wordmldata = $xmldeclaration . "\n<pass1Container>\n";
    $imagestring = "";

    $fs = get_file_storage();
    // Prepare filerecord array for creating each new image file.
    $fileinfo = array(
        'contextid' => $contextid,
        'component' => 'user',
        'filearea' => 'draft',
        'itemid' => 0,
        'filepath' => '/',
        'filename' => ''
        );

    // Open the Word 2010 Zip-formatted file and extract the WordProcessingML XML files.
    $zfh = zip_open($filename);
    if ($zfh) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Opened Zip file for reading", DEBUG_DEVELOPER);
        $zipentry = zip_read($zfh);
        while ($zipentry) {
            if (zip_entry_open($zfh, $zipentry, "r")) {
                $zefilename = zip_entry_name($zipentry);
                $zefilesize = zip_entry_filesize($zipentry);
                debugging(__FUNCTION__ . ":" . __LINE__ . ": zip_entry_name = $zefilename, size = $zefilesize", DEBUG_DEVELOPER);

                // Insert internal images into the files table.
                if (strpos($zefilename, "media")) { 
                    $imageformat = substr($zefilename, strrpos($zefilename, ".") +1);
                    $imagedata = zip_entry_read($zipentry, $zefilesize);
                    $imagename = basename($zefilename);
                    $imagesuffix = strtolower(substr(strrchr($zefilename, "."), 1));
                    // gif, png, jpg and jpeg handled OK, but bmp and other non-Internet formats are not.
                    $imagemimetype = "image/";
                    if ($imagesuffix == 'gif' or $imagesuffix == 'png') {
                        $imagemimetype .= $imagesuffix;
                    }

                    // 
                    $fileinfo['filename'] = $imagename;
                    $fileinfo['itemid'] = file_get_unused_draft_itemid();
                    $fs->create_file_from_string($fileinfo, $imagedata);
                    debugging(__FUNCTION__ . ":" . __LINE__ . ": created file " . $fileinfo['filename'] . 
                        ' with itemid = ' . $fileinfo['itemid'], DEBUG_DEVELOPER);

                    $imageurl = $CFG->wwwroot . '/draftfile.php/' . $contextid . '/user/draft/' . $fileinfo['itemid'] . '/' . $fileinfo['filename'];
                    if ($imagesuffix == 'jpg' or $imagesuffix == 'jpeg') {
                        $imagemimetype .= "jpeg";
                    }
                    // Handle recognised Internet formats only.
                    if ($imagemimetype != '') {
                        debugging(__FUNCTION__ . ":" . __LINE__ . ": media file name = $zefilename, imagename = " .
                            $imagename . ", imagesuffix = $imagesuffix, imagemimetype = $imagemimetype", DEBUG_DEVELOPER);
                        $imagestring .= '<file filename="media/' . $imagename . '" mime-type="' . $imagemimetype 
                            . '">' . $imageurl . "</file>\n";
                    }
                    else {
                        debugging(__FUNCTION__ . ":" . __LINE__ . ": ignore unsupported media file name $zefilename, imagename " .
                            " = $imagename, imagesuffix = $imagesuffix, imagemimetype = $imagemimetype", DEBUG_DEVELOPER);
                    }
                } else {
                    // Look for required XML files, read and wrap it, remove the XML declaration, and add it to the XML string.
                    switch ($zefilename) {
                        case "word/document.xml":
                            $wordmldata .= "<wordmlContainer>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</wordmlContainer>\n";
                            break;
                        case "docProps/core.xml":
                            $wordmldata .= "<dublinCore>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</dublinCore>\n";
                            break;
                        case "docProps/custom.xml":
                            $wordmldata .= "<customProps>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</customProps>\n";
                            break;
                        case "word/styles.xml":
                            $wordmldata .= "<styleMap>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</styleMap>\n";
                            break;
                        case "word/_rels/document.xml.rels":
                            $wordmldata .= "<documentLinks>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</documentLinks>\n";
                            break;
                        case "word/footnotes.xml":
                            $wordmldata .= "<footnotesContainer>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</footnotesContainer>\n";
                            break;
                        case "word/_rels/footnotes.xml.rels":
                            $wordmldata .= "<footnoteLinks>" . str_replace($xmldeclaration, zip_entry_read($zipentry, $zefilesize), "") . "</footnoteLinks>\n";
                            break;
                        case "word/_rels/settings.xml.rels":
                            $wordmldata .= "<settingsLinks>" . str_replace($xmldeclaration, "", zip_entry_read($zipentry, $zefilesize)) . "</settingsLinks>\n";
                            break;
                        default:
                            debugging(__FUNCTION__ . ":" . __LINE__ . ": Ignore $zefilename", DEBUG_DEVELOPER);
                    }
                }
            } else { // Can't read the file from the Word .docx file.
                zip_close($zfh);
                return false;
            }
            // Get the next file in the Zip package.
            $zipentry = zip_read($zfh);
        }  // End while loop.
        zip_close($zfh);
    } else { // Can't open the Word .docx file for reading.
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Cannot unzip Word file ('$filename') to read XML", DEBUG_DEVELOPER);
        debug_unlink($filename);
        return false;
    }

    // Add images section and close the merged XML file.
    $wordmldata .= "<imagesContainer>\n" . $imagestring . "</imagesContainer>\n"  . "</pass1Container>";

    // Pass 1 - convert WordML into linear XHTML.
    // Create a temporary file to store the merged WordML XML content to transform.
    $tempwordmlfilename = $CFG->dataroot . '/temp/' . basename($filename, ".tmp") . ".wml";

    // Write the WordML contents to be imported.
    if (($nbytes = file_put_contents($tempwordmlfilename, $wordmldata)) == 0) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Failed to save XML data to temporary file ('" .
            $tempwordmlfilename . "')", DEBUG_DEVELOPER);
        return false;
    }
    debugging(__FUNCTION__ . ":" . __LINE__ . ": XML data saved to $tempwordmlfilename", DEBUG_DEVELOPER);

    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 1 with stylesheet \"" . $stylesheet . "\"", DEBUG_DEVELOPER);
    $xsltproc = xslt_create();
    if (!($xsltoutput = xslt_process($xsltproc, $tempwordmlfilename, $stylesheet, null, null, $parameters))) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Transformation failed", DEBUG_DEVELOPER);
        debug_unlink($tempwordmlfilename);
        return false;
    }
    debug_unlink($tempwordmlfilename);
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 1 succeeded, XHTML output fragment = " . str_replace("\n", "", substr($xsltoutput, 0, 200)), DEBUG_DEVELOPER);
    // Strip out superfluous namespace declarations on paragraph elements, which Moodle 2.7/2.8 on Windows seems to throw in.
    $xsltoutput = str_replace(' xmlns="http://www.w3.org/1999/xhtml"', '', $xsltoutput);
    $xsltoutput = str_replace(' xmlns=""', '', $xsltoutput);

    // Write output of Pass 1 to a temporary file, for use in Pass 2.
    $tempxhtmlfilename = $CFG->dataroot . '/temp/' . basename($filename, ".tmp") . ".if1";
    if (($nbytes = file_put_contents($tempxhtmlfilename, $xsltoutput )) == 0) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Failed to save XHTML data to temporary file ('" .
            $tempxhtmlfilename . "')", DEBUG_DEVELOPER);
        return false;
    }
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import Pass 1 output XHTML data saved to $tempxhtmlfilename", DEBUG_DEVELOPER);

    // Pass 2 - tidy up linear XHTML a bit.
    $stylesheet = dirname(basename(__FILE__)) . "/" . $word2mqxmlstylesheet2;
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 2 with stylesheet \"" . $stylesheet . "\"", DEBUG_DEVELOPER);
    if (!($xsltoutput = xslt_process($xsltproc, $tempxhtmlfilename, $stylesheet, null, null, $parameters))) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Import Pass 2 Transformation failed", DEBUG_DEVELOPER);
        debug_unlink($tempxhtmlfilename);
        return false;
    }
    debug_unlink($tempxhtmlfilename);
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import Pass 2 succeeded, XHTML output fragment = " . str_replace("\n", "", substr($xsltoutput, 600, 500)), DEBUG_DEVELOPER);

    // Keep the converted XHTML file for debugging if developer debugging enabled.
    if (debugging(null, DEBUG_DEVELOPER)) {
        $tempxhtmlfilename = $CFG->dataroot . '/temp/' . basename($filename, ".tmp") . ".xhtml";
        if (($nbytes = file_put_contents($tempxhtmlfilename, $xsltoutput)) == 0) {
            return false;
        }
    }

    return $xsltoutput;
}   // end convert_to_xhtml


/**
 * Get the HTML body from the converted Word file
 *
 * A string containing XHTML is returned
 *
 * @return string
 */
function get_html_body($xhtmlstring) {
    debugging(__FUNCTION__ . "(xhtmlstring = \"" . substr($xhtmlstring, 0, 100) . "\")", DEBUG_DEVELOPER);

    $bodystart = stripos($xhtmlstring, '<body>') + strlen('<body>');
    $bodylength = strripos($xhtmlstring, '</body>') - $bodystart;
    //debugging(__FUNCTION__ . ":" . __LINE__ . ": bodystart = {$bodystart}, bodylength = {$bodylength}", DEBUG_DEVELOPER);
    if ($bodystart  !== false || $bodylength !== false) {
        $xhtmlbody = substr($xhtmlstring, $bodystart, $bodylength);
    } else {
        debugging(__FUNCTION__ . "() -> Invalid XHTML, using original cdata string", DEBUG_DEVELOPER);
        $xhtmlbody = $xhtmlstring;
    }

    debugging(__FUNCTION__ . "() -> |" . str_replace("\n", "", substr($xhtmlbody, 0, 100)) . " ...|", DEBUG_DEVELOPER);
    return $xhtmlbody;
}

/*
 * Delete temporary files if debugging disabled
 */
function debug_unlink($filename) {
    if (!debugging(null, DEBUG_DEVELOPER)) {
        unlink($filename);
    }
}

