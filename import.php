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
require_once($CFG->libdir . '/filestorage/file_storage.php');
require_once($CFG->dirroot . '/repository/lib.php');
require_once("$CFG->libdir/xmlize.php");
require(dirname(basename(__FILE__)) . '/lib.php');
// Include XSLT processor functions
require_once(dirname(basename(__FILE__)) . "/xsl_emulate_xslt.inc");


$itemid = required_param('itemid', PARAM_INT);
$contextid = required_param('ctx_id', PARAM_INT);

$context = context::instance_by_id($contextid);
$PAGE->set_context($context);


// Get the reference of the uploaded file, save it as a temporary file, and then delete it from the files table
$fs = get_file_storage();
$filearray = $fs->get_area_files($contextid, 'user', 'draft', $itemid);
foreach($filearray as $file) {
    if ($file->is_directory()) {
        continue;
    }
    $tmp_filename = $file->copy_content_to_temp();
    //debugging(basename(__FILE__) . " (" . __LINE__ . "): tmp_filename " . $file->get_filename() . " saved to " . $tmp_filename, DEBUG_DEVELOPER);
}
$filearray = $fs->delete_area_files($contextid, 'user', 'draft', $itemid);


debugging(basename(__FILE__) . " (" . __LINE__ . "): tmp_filename = " . $tmp_filename, DEBUG_DEVELOPER);
//echo "{\"html\": \"" . $tmp_filename . "\"}";
//exit;

$html_text = convert_to_xhtml($tmp_filename);
$html_text = get_html_body($html_text);

// Delete the temporary file now that we're finished with it.
debug_unlink($tmp_filename);

if (($json_text = json_encode($html_text)) === FALSE) {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): JSON encoding failed ", DEBUG_DEVELOPER);
    echo '{"error": "' . get_string('transformationfailed', 'atto_wordimport') . '"}';
} else {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): json_text = " . str_replace(substr($json_text, 0, 500), "\n", " "), DEBUG_DEVELOPER);
    echo "{\"html\": " . $json_text . "}";
}



?>
