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

use \booktool_wordimport\wordconverter;

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
        'pluginname',
        'editorlabel'
    );

    $PAGE->requires->strings_for_js($strings, 'atto_wordimport');
}

/**
 * Sends the parameters to JS module.
 *
 * @param string $elementid - unused
 * @param array $options the options for the editor, including the context
 * @param null $fpoptions - unused
 * @return array
 */
function atto_wordimport_params_for_js($elementid, $options, $fpoptions) {
    global $CFG, $USER;
    require_once($CFG->dirroot . '/repository/lib.php');  // Load constants.

    // Disabled if:
    // - Not logged in or guest.
    // - Files are not allowed.
    // - Only URL are supported.
    $disabled = !isloggedin() || isguestuser() ||
            (!isset($options['maxfiles']) || $options['maxfiles'] == 0) ||
            (isset($options['return_types']) && !($options['return_types'] & ~FILE_EXTERNAL));

    $params = array('disabled' => $disabled, 'area' => array(), 'usercontext' => null);

    if (!$disabled) {
        $params['usercontext'] = context_user::instance($USER->id)->id;
        foreach (array('itemid', 'context', 'areamaxbytes', 'maxbytes', 'subdirs', 'return_types') as $key) {
            if (isset($options[$key])) {
                if ($key === 'context' && is_object($options[$key])) {
                    // Just context id is enough.
                    $params['area'][$key] = $options[$key]->id;
                } else {
                    $params['area'][$key] = $options[$key];
                }
            }
        }
    }

    return $params;
}

/**
 * Extract the WordProcessingML XML files from the .docx file, and use a sequence of XSLT
 * steps to convert it into XHTML
 *
 * @param string $wordfilename name of file uploaded to file repository as a draft
 * @param int $usercontextid ID of draft file area where images should be stored
 * @param int $draftitemid ID of particular group in draft file area where images should be stored
 * @return string XHTML content extracted from Word file
 */
function atto_wordimport_convert_to_xhtml(string $wordfilename, int $usercontextid, int $draftitemid) {
    global $CFG, $USER;

    // Check that we can unzip the Word .docx file into its component files.
    $zipres = zip_open($wordfilename);
    if (!is_resource($zipres)) {
        // Cannot unzip file.
        atto_wordimport_debug_unlink($wordfilename);
        throw new \moodle_exception('cannotunzipfile', 'error');
    }

    // Pre-XSLT preparation: merge the WordML and image content from the .docx Word file into one large XML file.
    // Initialise an XML string to use as a wrapper around all the XML files.
    $xmldeclaration = '<?xml version="1.0" encoding="UTF-8"?>';
    $wordmldata = $xmldeclaration . "\n<pass1Container>\n";
    $imagestring = "";

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
            throw new \moodle_exception('errorunzippingfiles', 'error');
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

                $imageurl = "$CFG->wwwroot/draftfile.php/$usercontextid/user/draft/$draftitemid/$imagenameunique";
                // Return all the details of where the file is stored, even though we don't need them at the moment.
                $imagestring .= "<file filename=\"media/{$imagename}\"";
                $imagestring .= " contextid=\"{$usercontextid}\" itemid=\"{$draftitemid}\"";
                $imagestring .= " name=\"{$imagenameunique}\" url=\"{$imageurl}\">{$imageurl}</file>\n";
            // @codingStandardsIgnoreLine } else {
                // @codingStandardsIgnoreLine debugging(__FUNCTION__ . ":" . __LINE__ . ": ignore unsupported media file $zefilename" .
                // @codingStandardsIgnoreLine     " = $imagename, imagesuffix = $imagesuffix", DEBUG_WORDIMPORT);
            }
        } else {
            // Look for required XML files, read and wrap it, remove the XML declaration, and add it to the XML string.
            // Read and wrap XML files, remove the XML declaration, and add them to the XML string.
            $xmlfiledata = preg_replace('/<\?xml version="1.0" ([^>]*)>/', "", zip_entry_read($zipentry, $zefilesize));
            switch ($zefilename) {
                case "word/document.xml":
                    $wordmldata .= "<wordmlContainer>" . $xmlfiledata . "</wordmlContainer>\n";
                    break;
                case "docProps/core.xml":
                    $wordmldata .= "<dublinCore>" . $xmlfiledata . "</dublinCore>\n";
                    break;
                case "docProps/custom.xml":
                    $wordmldata .= "<customProps>" . $xmlfiledata . "</customProps>\n";
                    break;
                case "word/styles.xml":
                    $wordmldata .= "<styleMap>" . $xmlfiledata . "</styleMap>\n";
                    break;
                case "word/_rels/document.xml.rels":
                    $wordmldata .= "<documentLinks>" . $xmlfiledata . "</documentLinks>\n";
                    break;
                case "word/footnotes.xml":
                    $wordmldata .= "<footnotesContainer>" . $xmlfiledata . "</footnotesContainer>\n";
                    break;
                case "word/_rels/footnotes.xml.rels":
                    $wordmldata .= "<footnoteLinks>" . $xmlfiledata . "</footnoteLinks>\n";
                    break;
                // @codingStandardsIgnoreLine case "word/_rels/settings.xml.rels":
                    // @codingStandardsIgnoreLine $wordmldata .= "<settingsLinks>" . $xmlfiledata . "</settingsLinks>\n";
                    // @codingStandardsIgnoreLine break;
                default:
                    // @codingStandardsIgnoreLine debugging(__FUNCTION__ . ":" . __LINE__ . ": Ignore $zefilename", DEBUG_WORDIMPORT);
            }
        }
        // Get the next file in the Zip package.
        $zipentry = zip_read($zipres);
    }  // End while loop.
    zip_close($zipres);

    // Add images section.
    $wordmldata .= "<imagesContainer>\n" . $imagestring . "</imagesContainer>\n";
    // Close the merged XML file.
    $wordmldata .= "</pass1Container>";

    // Pass 1 - convert WordML into linear XHTML, embed images, and map "Heading 1" style as defined in config.
    $word2xml = new wordconverter();
    $word2xml->set_heading1styleoffset((int) get_config('atto_wordimport', 'heading1stylelevel'));
    $word2xml->set_imagehandling('embedded');
    $imagesforzipping = array(); // Unused, as images are embedded for Atto.
    $xsltoutput = $word2xml->import($wordfilename, $imagesforzipping);
    $xsltoutput = $word2xml->body_only($xsltoutput);

    return $xsltoutput;
}   // End function atto_wordimport_convert_to_xhtml.

/**
 * Delete temporary files if debugging disabled
 *
 * @param string $filename name of file to be deleted
 * @return void
 */
function atto_wordimport_debug_unlink(string $filename) {
    if (DEBUG_WORDIMPORT == 0) {
        unlink($filename);
    }
}

