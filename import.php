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

require(dirname(basename(__FILE__)) . '/../../../../../config.php');
require(dirname(basename(__FILE__)) . '/lib.php');

$itemid = required_param('itemid', PARAM_INT);
$contextid = required_param('ctx_id', PARAM_INT);

$context = context::instance_by_id($contextid);
$PAGE->set_context($context);

// Get the reference of the uploaded file, save it as a temporary file, and then delete it from the files table.
$fs = get_file_storage();
$filearray = $fs->get_area_files($contextid, 'user', 'draft', $itemid);
foreach ($filearray as $file) {
    if ($file->is_directory()) {
        continue;
    }
    $tmpfilename = $file->copy_content_to_temp();
    debugging(basename(__FILE__) . " (" . __LINE__ . "): tmp_filename = " . $tmpfilename, DEBUG_DEVELOPER);
}
$filearray = $fs->delete_area_files($contextid, 'user', 'draft', $itemid);

// Convert the Word file into XHTML with images.
$htmltext = atto_wordimport_convert_to_xhtml($tmpfilename, $contextid);

// Handle the images by encoding them as JSON items.
$imagescontainerstart = stripos($htmltext, '<imagesContainer>') + strlen('<imagesContainer>');
$imagescontainerend = stripos($htmltext, '</imagesContainer>');
$imagesstring = "";
debugging(basename(__FILE__) . " (" . __LINE__ . "): imagescontainerend: {$imagescontainerend}, imagescontainerstart {$imagescontainerstart}", DEBUG_DEVELOPER);
if ($imagescontainerend and $imagescontainerstart and $imagescontainerend - $imagescontainerstart > 2) {
    // "filename="media/image1.jpg" mime-type="image/jpeg" contextid="5" itemid="693725586" name="image1.jpg" url="http://localhost/m29/draftfile.php/5/user/draft/693725586/image1.jpg"
    $n_matches = preg_match_all('|<file filename="media/([^"]*)" mime-type="image/([^"]*)" contextid="([^"]*)" itemid="([^"]*)" name="([^"]*)" url="([^"]*)"/>|', substr($htmltext, $imagescontainerstart), $imagefiles, PREG_SET_ORDER);
    debugging(basename(__FILE__) . " (" . __LINE__ . "): image file n_matches: \"{$n_matches}\"", DEBUG_DEVELOPER);
    if ($n_matches !== false) {
        $imagesstring = ",";
        foreach ($imagefiles as $file) {
            $imagestring = "\"url\": " . json_encode($file[6]) . ",";
            $imagestring .= "\"id\":\"{$file[4]}\",\"name\":\"{$file[5]}\"";
            $imagesstring .= $imagestring;
            debugging(basename(__FILE__) . " (" . __LINE__ . "): image file details: \"{$imagestring}\"", DEBUG_DEVELOPER);
        }
    }
}

// Get the content inside the HTML body tags only, ignore metadata for now.
$htmltext = atto_wordimport_get_html_body($htmltext);

// Return the XHTML in JSON-encoded format, if it was encoded OK.
$htmltextjson = json_encode($htmltext);
if ($htmltextjson === false) {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): JSON encoding failed ", DEBUG_DEVELOPER);
    echo '{"error": "' . get_string('transformationfailed', 'atto_wordimport') . '"}';
} else {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): jsontext = |" .
        str_replace("\n", " ", substr($htmltextjson, 0, 500)) . "...|", DEBUG_DEVELOPER);
    echo "{\"html\": " . $htmltextjson . $imagesstring . "}";

// Delete the temporary file now that we're finished with it.
atto_wordimport_debug_unlink($tmpfilename);

}
