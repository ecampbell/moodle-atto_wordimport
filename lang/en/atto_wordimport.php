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
 * Strings for component 'atto_wordimport', language 'en'.
 *
 * @package    atto_wordimport
 * @copyright  2015 Eoin Campbell
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

$string['browse'] = 'Browse';
$string['browserepositories'] = 'Browse repositories...';
$string['cancel'] = 'Cancel';
$string['converting'] = 'Importing, please wait...';
$string['createword'] = 'Create Word Import';
$string['customstyle'] = 'Custom style';
$string['defaultflavor'] = 'Default Ice Cream Flavor';
$string['dialogtitle'] = 'Enter Preferences';
$string['enteralt'] = 'Describe this image for someone who cannot see it';
$string['enterflavor'] = 'Enter Ice Cream Flavor';
$string['enterurl'] = 'Enter URL';
$string['importfile'] = 'Import Word file';
$string['insert'] = 'Insert';
$string['nothingtoinsert'] = 'Nothing to insert!';
$string['pluginname'] = 'Word import (Atto)';
$string['settings'] = 'Word Import (Atto)';
$string['uploading'] = 'Uploading, please wait...';
$string['visible'] = 'Visible';

// Strings used in import.php
$string['cannotopentempfile'] = 'Cannot open temporary file <b>{$a}</b>';
$string['cannotreadzippedfile'] = 'Cannot read Zipped file <b>{$a}</b>';
$string['cannotwritetotempfile'] = 'Cannot write to temporary file <b>{$a}</b>';
$string['stylesheetunavailable'] = 'XSLT Stylesheet <b>{$a}</b> is not available';
$string['transformationfailed'] = 'XSLT transformation failed (<b>{$a}</b>)';
$string['xsltunavailable'] = 'You need the XSLT library installed in PHP to save this Word file';

// Strings used in JavaScript
$string['xmlnotsupported'] = 'Files in XML format not supported: <b>{$a}</b>';
$string['docnotsupported'] = 'Files in Word 2003 format not supported: <b>{$a}</b>, use Moodle2Word 3.x instead';
$string['htmlnotsupported'] = 'Files in HTML format not supported: <b>{$a}</b>';
$string['htmldocnotsupported'] = 'Incorrect Word format: please use <i>File>Save As...</i> to save <b>{$a}</b> in native Word 2010 (.docx) format and import again';
