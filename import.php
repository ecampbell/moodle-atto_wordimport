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
 * Atto text editor import Microsoft Word file and convert to HTML
 *
 * @package   atto_wordimport
 * @copyright 2015 Eoin Campbell
 * @license   http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

require(dirname(__FILE__) . '/../../../../../config.php');
require_once($CFG->libdir . '/filestorage/file_storage.php');
require_once($CFG->dirroot . '/repository/lib.php');
require_once("$CFG->libdir/xmlize.php");
require_once($CFG->dirroot.'/lib/uploadlib.php');
require(dirname(__FILE__) . '/util.php');
// Include XSLT processor functions
require_once(dirname(__FILE__) . "/xsl_emulate_xslt.inc");


$wordfileurl = required_param('wordfileurl', PARAM_TEXT);
$contextid = optional_param('context', SYSCONTEXTID, PARAM_INT);
$elementid = optional_param('elementid', '', PARAM_TEXT);

$context = context::instance_by_id($contextid);
if ($context->contextlevel == CONTEXT_MODULE) {
    // Module context.
    $cm = $DB->get_record('course_modules', array('id' => $context->instanceid));
    require_login($cm->course, true, $cm);
} else if (($coursecontext = $context->get_course_context(false)) && $coursecontext->id != SITEID) {
    // Course context or block inside the course.
    require_login($coursecontext->instanceid);
    $PAGE->set_context($context);
} else {
    // Block that is not inside the course, user or system context.
    require_login();
    $PAGE->set_context($context);
}

// Guests can never manage files.
if (isguestuser()) {
    print_error('noguest');
}


// IMPORT FUNCTIONS START HERE

/**
 * Perform required pre-processing, i.e. convert Word file into XML
 *
 * Extract the WordProcessingML XML files from the .docx file, and use a sequence of XSLT
 * steps to convert it into Moodle Question XML
 *
 * @return boolean success
 */
function convert_to_xhtml($filename) {
    global $CFG, $USER, $COURSE, $OUTPUT;

    $word2mqxml_stylesheet1 = 'wordml2xhtml_pass1.xsl';      // XSLT stylesheet containing code to convert XHTML into Word-compatible XHTML for question import
    $word2mqxml_stylesheet2 = 'wordml2xhtml_pass2.xsl';      // XSLT stylesheet containing code to convert XHTML into Word-compatible XHTML for question import

    debugging(__FUNCTION__ . ":" . __LINE__ . ": Word file = $filename", DEBUG_DEVELOPER);
    // Give XSLT as much memory as possible, to enable larger Word files to be imported.
    raise_memory_limit(MEMORY_HUGE);


    // XSLT stylesheet to convert WordML into initial XHTML format
    $stylesheet =  dirname(__FILE__) . "/" . $word2mqxml_stylesheet1;

    // Check that XSLT is installed, and the XSLT stylesheet is present
    if (!class_exists('XSLTProcessor') || !function_exists('xslt_create')) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": XSLT not installed", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('xsltunavailable', 'atto_wordimport'));
        return false;
    } else if(!file_exists($stylesheet)) {
        // XSLT stylesheet to transform WordML into XHTML doesn't exist
        debugging(__FUNCTION__ . ":" . __LINE__ . ": XSLT stylesheet missing: $stylesheet", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('stylesheetunavailable', 'atto_wordimport', $stylesheet));
        return false;
    }

    // Set common parameters for all XSLT transformations. Note that we cannot use arguments because the XSLT processor doesn't support them
    $parameters = array (
        'course_id' => $COURSE->id,
        'course_name' => $COURSE->fullname,
        'author_name' => $USER->firstname . ' ' . $USER->lastname,
        'moodle_country' => $USER->country,
        'moodle_language' => current_language(),
        'moodle_textdirection' => (right_to_left())? 'rtl': 'ltr',
        'moodle_release' => $CFG->release,
        'moodle_url' => $CFG->wwwroot . "/",
        'moodle_username' => $USER->username,
        'debug_flag' => debugging('', DEBUG_DEVELOPER)
    );


    // Pre-XSLT conversion preparation - re-package the XML and image content from the .docx Word file into one large XML file, to simplify XSLT processing

    // Initialise an XML string to use as a wrapper around all the XML files
    $xml_declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    $wordmlData = $xml_declaration . "\n<pass1Container>\n";
    $imageString = "";

    // Open the Word 2010 Zip-formatted file and extract the WordProcessingML XML files
    $zfh = zip_open($filename);
    if (is_resource($zfh)) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Opened Zip file for reading", DEBUG_DEVELOPER);
        $zip_entry = zip_read($zfh);
        while ($zip_entry !== FALSE) {
            if (zip_entry_open($zfh, $zip_entry, "r")) {
                $ze_filename = zip_entry_name($zip_entry);
                $ze_filesize = zip_entry_filesize($zip_entry);
                debugging(__FUNCTION__ . ":" . __LINE__ . ": zip_entry_name = $ze_filename, size = $ze_filesize", DEBUG_DEVELOPER);

                // Look for internal images
                if (strpos($ze_filename, "media") !== FALSE) { 
                    $imageFormat = substr($ze_filename, strrpos($ze_filename, ".") +1);
                    $imageData = zip_entry_read($zip_entry, $ze_filesize);
                    $imageName = basename($ze_filename);
                    $imageSuffix = strtolower(substr(strrchr($ze_filename, "."), 1));
                    // gif, png, jpg and jpeg handled OK, but bmp and other non-Internet formats are not
                    $imageMimeType = "image/";
                    if ($imageSuffix == 'gif' or $imageSuffix == 'png') {
                        $imageMimeType .= $imageSuffix;
                    }
                    if ($imageSuffix == 'jpg' or $imageSuffix == 'jpeg') {
                        $imageMimeType .= "jpeg";
                    }
                    // Handle recognised Internet formats only
                    if ($imageMimeType != '') {
                        debugging(__FUNCTION__ . ":" . __LINE__ . ": media file name = $ze_filename, imageName = $imageName, imageSuffix = $imageSuffix, imageMimeType = $imageMimeType", DEBUG_DEVELOPER);
                        $imageString .= '<file filename="media/' . $imageName . '" mime-type="' . $imageMimeType . '">' . base64_encode($imageData) . "</file>\n";
                    }
                    else {
                        debugging(__FUNCTION__ . ":" . __LINE__ . ": ignore unsupported media file name $ze_filename, imageName = $imageName, imageSuffix = $imageSuffix, imageMimeType = $imageMimeType", DEBUG_DEVELOPER);
                    }
                // Look for required XML files
                } else {
                    // If a required XML file is encountered, read it, wrap it, remove the XML declaration, and add it to the XML string
                    switch ($ze_filename) {
                      case "word/document.xml":
                          $wordmlData .= "<wordmlContainer>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</wordmlContainer>\n";
                          break;
                      case "docProps/core.xml":
                          $wordmlData .= "<dublinCore>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</dublinCore>\n";
                          break;
                      case "docProps/custom.xml":
                          $wordmlData .= "<customProps>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</customProps>\n";
                          break;
                      case "word/styles.xml":
                          $wordmlData .= "<styleMap>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</styleMap>\n";
                          break;
                      case "word/_rels/document.xml.rels":
                          $wordmlData .= "<documentLinks>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</documentLinks>\n";
                          break;
                      /*
                      case "word/footnotes.xml":
                          $wordmlData .= "<footnotesContainer>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</footnotesContainer>\n";
                          break;
                      case "word/_rels/footnotes.xml.rels":
                          $wordmlData .= "<footnoteLinks>" . str_replace($xml_declaration, zip_entry_read($zip_entry, $ze_filesize), "") . "</footnoteLinks>\n";
                          break;
                      case "word/_rels/settings.xml.rels":
                          $wordmlData .= "<settingsLinks>" . str_replace($xml_declaration, "", zip_entry_read($zip_entry, $ze_filesize)) . "</settingsLinks>\n";
                          break;
                      */
                      default:
                          debugging(__FUNCTION__ . ":" . __LINE__ . ": Ignore $ze_filename", DEBUG_DEVELOPER);
                    }
                }
            } else { // Can't read the file from the Word .docx file
                echo $OUTPUT->notification(get_string('cannotreadzippedfile', 'atto_wordimport', $filename));
                zip_close($zfh);
                return false;
            }
            // Get the next file in the Zip package
            $zip_entry = zip_read($zfh);
        }  // End while
        zip_close($zfh);
    } else { // Can't open the Word .docx file for reading
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Cannot unzip Word file ('$filename') to read XML", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('cannotopentempfile', 'atto_wordimport', $filename));
        debug_unlink($zipfile);
        return false;
    }

    // Add Base64 images section and close the merged XML file
    $wordmlData .= "<imagesContainer>\n" . $imageString . "</imagesContainer>\n"  . "</pass1Container>";


    // Pass 1 - convert WordML into linear XHTML
    // Create a temporary file to store the merged WordML XML content to transform
    if (!($temp_wordml_filename = tempnam($CFG->dataroot . '/temp/', "awi-"))) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Cannot open temporary file ('$temp_wordml_filename') to store XML", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('cannotopentempfile', 'atto_wordimport', $temp_wordml_filename));
        return false;
    }

    // Write the WordML contents to be imported
    if (($nbytes = file_put_contents($temp_wordml_filename, $wordmlData)) == 0) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Failed to save XML data to temporary file ('$temp_wordml_filename')", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('cannotwritetotempfile', 'atto_wordimport', $temp_wordml_filename . "(" . $nbytes . ")"));
        return false;
    }
    debugging(__FUNCTION__ . ":" . __LINE__ . ": XML data saved to $temp_wordml_filename", DEBUG_DEVELOPER);

    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 1 with stylesheet \"" . $stylesheet . "\"", DEBUG_DEVELOPER);
    $xsltproc = xslt_create();
    if(!($xslt_output = xslt_process($xsltproc, $temp_wordml_filename, $stylesheet, null, null, $parameters))) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Transformation failed", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('transformationfailed', 'atto_wordimport', "(XSLT: " . $stylesheet . "; XML: " . $temp_wordml_filename . ")"));
        debug_unlink($temp_wordml_filename);
        return false;
    }
    debug_unlink($temp_wordml_filename);
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 1 succeeded, XHTML output fragment = " . str_replace("\n", "", substr($xslt_output, 0, 200)), DEBUG_DEVELOPER);
    // Strip out some superfluous namespace declarations on paragraph elements, which Moodle 2.7/2.8 on Windows seems to throw in
    $xslt_output = str_replace('<p xmlns="http://www.w3.org/1999/xhtml"', '<p', $xslt_output);
    $xslt_output = str_replace(' xmlns=""', '', $xslt_output);

    // Write output of Pass 1 to a temporary file, for use in Pass 2
    $temp_xhtml_filename = $CFG->dataroot . '/temp/' . basename($temp_wordml_filename, ".tmp") . ".if1";
    if (($nbytes = file_put_contents($temp_xhtml_filename, $xslt_output )) == 0) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Failed to save XHTML data to temporary file ('$temp_xhtml_filename')", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('cannotwritetotempfile', 'atto_wordimport', $temp_xhtml_filename . "(" . $nbytes . ")"));
        return false;
    }
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import Pass 1 output XHTML data saved to $temp_xhtml_filename", DEBUG_DEVELOPER);


    // Pass 2 - tidy up linear XHTML a bit
    // Prepare for Import Pass 2 XSLT transformation
    $stylesheet =  dirname(__FILE__) . "/" . $word2mqxml_stylesheet2;
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import XSLT Pass 2 with stylesheet \"" . $stylesheet . "\"", DEBUG_DEVELOPER);
    if(!($xslt_output = xslt_process($xsltproc, $temp_xhtml_filename, $stylesheet, null, null, $parameters))) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Import Pass 2 Transformation failed", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('transformationfailed', 'atto_wordimport', "(XSLT: " . $stylesheet . "; XHTML: " . $temp_xhtml_filename . ")"));
        debug_unlink($temp_xhtml_filename);
        return false;
    }
    debug_unlink($temp_xhtml_filename);
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Import Pass 2 succeeded, XHTML output fragment = " . str_replace("\n", "", substr($xslt_output, 600, 500)), DEBUG_DEVELOPER);

    // Write the Pass 2 XHTML output to a temporary file
    $temp_xhtml_filename = $CFG->dataroot . '/temp/' . basename($temp_wordml_filename, ".tmp") . ".if2";
    if (($nbytes = file_put_contents($temp_xhtml_filename, $xslt_output)) == 0) {
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Failed to save XHTML data to temporary file ('$temp_xhtml_filename')", DEBUG_DEVELOPER);
        echo $OUTPUT->notification(get_string('cannotwritetotempfile', 'atto_wordimport', $temp_xhtml_filename . "(" . $nbytes . ")"));
        return false;
    }
    debugging(__FUNCTION__ . ":" . __LINE__ . ": Pass 2 output XHTML data saved to $temp_xhtml_filename", DEBUG_DEVELOPER);
    //file_put_contents($CFG->dataroot . '/temp/' . basename($temp_wordml_filename, ".tmp") . ".if2", "<pass3Container>\n" . $xslt_output . get_text_labels() . "\n</pass3Container>");

    // Keep the original Word file for debugging if developer debugging enabled
    if (debugging(null, DEBUG_DEVELOPER)) {
        $copied_input_file = $CFG->dataroot . '/temp/' . basename($temp_wordml_filename, ".tmp") . ".docx";
        copy($filename, $copied_input_file);
        debugging(__FUNCTION__ . ":" . __LINE__ . ": Copied $filename to $copied_input_file", DEBUG_DEVELOPER);
    }


    return $xslt_output;
}   // end convert_to_xhtml

$html_content = convert_to_xhtml($wordfileurl);
$html_content = str_replace("\n", " ", $html_content);
debugging(__FILE__ . ":" . __LINE__ . ": Conversion succeeded, XHTML fragment = " . substr($html_content , 0, 500), DEBUG_DEVELOPER);

$json_html_content = str_replace('"', '\"', $html_content);
debugging(__FILE__ . ":" . __LINE__ . ": HTML escape succeeded, XHTML fragment = " . substr($json_html_content , 0, 500), DEBUG_DEVELOPER);

echo "{\"html\": \"" . $json_html_content . "\"}";

?>
