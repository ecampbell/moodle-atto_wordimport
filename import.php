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
// Include XSLT processor functions.
require_once(dirname(basename(__FILE__)) . "/xsl_emulate_xslt.inc");


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

// Convert the Word file into XHTML, and then grab just the <body> content.
$htmltext = convert_to_xhtml($tmpfilename, $contextid);
$htmltext = get_html_body($htmltext);
$jsontext = json_encode($htmltext);

// Delete the temporary file now that we're finished with it.
debug_unlink($tmpfilename);

// Return the XHTML in JSON-encoded format, if it was encoded OK.
$jsontext = json_encode($htmltext);
if ($jsontext === false) {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): JSON encoding failed ", DEBUG_DEVELOPER);
    echo '{"error": "' . get_string('transformationfailed', 'atto_wordimport') . '"}';
} else {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): jsontext = |" . 
        str_replace("\n", " ", substr($jsontext, 0, 500)) . "...|", DEBUG_DEVELOPER);
    echo "{\"html\": " . $jsontext . "}";
}
