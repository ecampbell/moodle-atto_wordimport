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
//define('DEBUG_WORDIMPORT', E_NONE);


require(dirname(basename(__FILE__)) . '/../../../../../config.php');
// Include XSLT processor functions.
require_once(dirname(basename(__FILE__)) . "/xsl_emulate_xslt.inc");
require(dirname(basename(__FILE__)) . '/lib.php');

$itemid = required_param('itemid', PARAM_INT);
$contextid = required_param('ctx_id', PARAM_INT);
$filename = required_param('filename', PARAM_TEXT);

$context = context::instance_by_id($contextid);
$PAGE->set_context($context);

// Check that this user is logged in before proceeding.
require_login();

// Get the reference of the uploaded file, save it as a temporary file, and then delete it from the files table.
debugging(basename(__FILE__) . " (" . __LINE__ . "): filename = " . $filename, DEBUG_WORDIMPORT);
$fs = get_file_storage();
$file = $fs->get_file($contextid, 'user', 'draft', $itemid, '/', basename($filename));
if ($file) {
    $tmpfilename = $file->copy_content_to_temp();
    $file->delete();
    debugging(basename(__FILE__) . " (" . __LINE__ . "): \"{$filename}\" saved to \"{$tmpfilename}\"", DEBUG_WORDIMPORT);
}


// Convert the Word file into XHTML with images, and delete it once we're finished.
$htmltext = atto_wordimport_convert_to_xhtml($tmpfilename, $contextid, $itemid);
atto_wordimport_debug_unlink($tmpfilename);

// Get the content inside the HTML body tags only, ignore metadata for now.
$htmltext = atto_wordimport_get_html_body($htmltext);

// Return the XHTML in JSON-encoded format, if it was encoded OK.
$htmltextjson = json_encode($htmltext);
if ($htmltextjson === false) {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): JSON encoding failed ", DEBUG_WORDIMPORT);
    echo '{"error": "' . get_string('transformationfailed', 'atto_wordimport') . '"}';
} else {
    debugging(basename(__FILE__) . " (" . __LINE__ . "): jsontext = |" .
        str_replace("\n", " ", substr($htmltextjson, 0, 500)) . "...|", DEBUG_WORDIMPORT);
    echo "{\"html\": " . $htmltextjson . "}";

}
