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

define('AJAX_SCRIPT', true);
// Development: turn on all debug messages and strict warnings.
define('DEBUG_WORDIMPORT', E_ALL | E_STRICT);
//define('DEBUG_WORDIMPORT', 0);

require(__DIR__ . '/../../../../../config.php');
// Include XSLT processor functions.
require_once(__DIR__ . "/xsl_emulate_xslt.inc");
require(__DIR__ . '/lib.php');

$itemid = required_param('itemid', PARAM_INT);
$contextid = required_param('ctx_id', PARAM_INT);
$filename = required_param('filename', PARAM_TEXT);
$sesskey = required_param('sesskey', PARAM_TEXT);


list($context, $course, $cm) = get_context_info_array($contextid);

$context = context::instance_by_id($contextid);

// Check that this user is logged in before proceeding.
require_login($course, false, $cm);
require_sesskey();

$PAGE->set_context($context);

// Get the reference only of this users' uploaded file, to avoid rogue users' accessing other peoples files.
$fs = get_file_storage();
$usercontext = context_user::instance($USER->id);
$file = $fs->get_file($usercontext->id, 'user', 'draft', $itemid, '/', basename($filename));

if ($file) {
    // We have to save the uploaded as a real temporary file so we can process it using zip_open(), etc.
    $tmpfilename = $file->copy_content_to_temp();
    debugging(basename(__FILE__) . " (" . __LINE__ . "): \"{$filename}\" saved to \"{$tmpfilename}\"", DEBUG_WORDIMPORT);
    // But we delete it from the draft file area to avoid a name-clash message if it is re-uploaded in the same edit.
    $file->delete();

    // Convert the Word file into XHTML with images, and delete it once we're finished.
    $htmltext = atto_wordimport_convert_to_xhtml($tmpfilename, $contextid, $itemid);
    atto_wordimport_debug_unlink($tmpfilename);

    if ($htmltext !== false) {
         debugging(basename(__FILE__) . " (" . __LINE__ . "): htmltext = |" .
                str_replace("\n", " ", substr($htmltext, 0, 500)) . "...|", DEBUG_WORDIMPORT);
         // Get the body content only, ignoring any metadata in the head.
         $bodytext = atto_wordimport_get_html_body($htmltext);
        // Convert the string to JSON-encoded format.
        $htmltextjson = json_encode($bodytext);
        if ($htmltextjson) {
            echo '{"html": ' . $htmltextjson . '}';
        } else {
            debugging(basename(__FILE__) . " (" . __LINE__ . "): JSON encoding failed ", DEBUG_WORDIMPORT);
            echo '{"error": "' . get_string('cannotuploadfile') . "}";
        }
    } else {
        debugging(basename(__FILE__) . " (" . __LINE__ . "): File conversion failed ", DEBUG_WORDIMPORT);
        echo '{"error": "' . get_string('cannotuploadfile') . '"}';
    }
} else {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): File access failed", DEBUG_WORDIMPORT);
    echo '{"error": "' . get_string('filenotreadable') . '"}';
}
