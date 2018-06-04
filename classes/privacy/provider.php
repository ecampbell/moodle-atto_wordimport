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
 * Privacy information for administration tool upload enrolment methods- no user data stored.
 *
 * @package     atto_wordimport
 * @copyright   2018 Eoin Campbell
 * @license     http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

namespace atto_wordimport\privacy;

defined('MOODLE_INTERNAL') || die();

/**
 * Privacy Subsystem for atto_wordimport implementing null_provider.
 *
 * @copyright   2018 Eoin Campbell
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */
class provider implements \core_privacy\local\metadata\null_provider {

    /**
     * Get the language string identifier with the component's language
     * file to explain why this plugin stores no data.
     *
     * @return  string
     */
    public static function get_reason() {
        return 'privacy:metadata';
    }
}


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
 * Atto text editor import Microsoft Word files - version.
 *
 * @package    atto_wordimport
 * @copyright  2015, 2016 Eoin Campbell
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

defined('MOODLE_INTERNAL') || die();

$plugin->version   = 2017111201;        // The current plugin version (Date: YYYYMMDDXX).
$plugin->requires  = 2014051200;        // Requires Moodle 2.7 or higher, when Atto was added to core.
$plugin->component = 'atto_wordimport';  // Full name of the plugin (used for diagnostics).
$plugin->maturity  = MATURITY_STABLE;
$plugin->release   = '1.3.7 (Build 2018060401)'; // Human readable version information.


Release notes
-------------

Date          Version   Comment
2018/06/04    1.3.7     Support Privacy API for GDPR compliance
2017/11/12    1.3.6     Treat formatted lists as Bullet Lists, remove default line-height property
2017/09/11    1.3.5     Disable Word Import button if the editor field does not support filearea
2017/08/25    1.3.4     Don't raise the memory limit for XSLT any more
2017/08/14    1.3.3     Handle list item duplication bug inside adjacent table cells
2017/08/07    1.3.2     Handle Bootstrap Alert components
2017/04/23    1.3.1     Use Bootstrap panel class for textboxes, use canonical tempdir reference
2017/02/12    1.3.0     Handle textboxes (convert from Word tables) and quotations
2017/01/05    1.2.1     Fix error when handling empty table cells
2016/12/13    1.2.0     Improve formatting of tables and figures, support RTL languages better
2016/01/20    1.1.2     Clean up spans to omit default text colour, remove underline from links.
2015/12/08    1.1.1     Fix code to pass codechecker, use POST instead of GET to upload Word file
2015/12/08    1.1.0     Fix error importing some files, clean code to pass codechecker, upgrade maturity to stable
2015/10/15    1.0.0     Make uninstall process safer, mark first officially approved release on Moodle plugins directory
2015/10/14    0.9.5     Use Moodle exception class to flag errors, parameterize heading style to element map
2015/10/12    0.9.4     Correct context ID handling to work on all platforms, improve error handling
2015/09/28    0.9.3     Implement drag and drop support
2015/09/25    0.9.2     Insert content whereever the cursor is, not at the end of the text area
2015/09/24    0.9.1     Automate installation process so that icon is added to files group in editor
2015/09/24    0.9.0     Add support for importing equations
2015/09/21    0.0.9     Fix security issues when deleting imported Word file, other code cleanup
2015/09/17    0.0.8     Ensure image names are always unique when importing multiple times
2015/09/16    0.0.7     Clean up code to comply with Moodle guidelines and implement reviewer comments, import images
