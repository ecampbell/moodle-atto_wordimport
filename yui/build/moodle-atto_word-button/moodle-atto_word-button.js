YUI.add('moodle-atto_worddragdrop-button', function (Y, NAME) {

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

/*
 * @package    atto_worddragdrop
 * @copyright  2015 Eoin Campbell
 * @license    http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

/**
 * @module moodle-atto_worddragdrop-button
 */

/**
 * Atto text editor worddragdrop plugin.
 *
 * This plugin adds the ability to drop a Word file in and have it automatically
 * convert the contents into XHTML and into the text box.
 *
 * @namespace M.atto_worddragdrop
 * @class Button
 * @extends M.editor_atto.EditorPlugin
 */
var COMPONENTNAME = 'atto_worddragdrop';
var WORDTEMPLATE = '' +
            '<img src="{{url}}" alt="{{alt}}" ' +
                ' style="vertical-align:text-bottom;margin: 0 .5em;" class="img-responsive" ' +
                '{{#if id}}id="{{id}}" {{/if}}' +
                '/>';

Y.namespace('M.atto_worddragdrop').Button = Y.Base.create('button', Y.M.editor_atto.EditorPlugin, [], {
    /**
     * A reference to the current selection at the time that the dialogue
     * was opened.
     *
     * @property _currentSelection
     * @type Range
     * @private
     */
    _currentSelection: null,

    /**
     * Add event listeners.
     *
     * @method initializer
     */
    initializer : function() {
        var self = this;
        var host = this.get('host');
        var template = Y.Handlebars.compile(WORDTEMPLATE);
        this.editor.on('drop', function(e) {
            host.saveSelection();
            e = e._event;
            // Only handle the event if a Word file was dropped in.
            if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length && /^application\/vnd.openxmlformats-officedocument.wordprocessingml.document/.test(e.dataTransfer.files[0].type)) {
                e.preventDefault();
                e.stopPropagation();
                var options = host.get('filepickeroptions').image;
                var savepath = (options.savepath === undefined) ? '/' : options.savepath;
                var formData = new FormData();

                formData.append('repo_upload_file', e.dataTransfer.files[0]);
                formData.append('itemid', options.itemid);
                // List of repositories is an object rather than an array.  This makes iteration more awkward.
                var keys = Object.keys(options.repositories);
                for (var i=0; i<keys.length; i++) {
                    if (options.repositories[keys[i]].type === 'upload') {
                        formData.append('repo_id', options.repositories[keys[i]].id);
                        break;
                    }
                }
                formData.append('env', options.env);
                formData.append('sesskey', M.cfg.sesskey);
                formData.append('client_id', options.client_id);
                formData.append('savepath', savepath);
                formData.append('ctx_id', options.context.id);

                // Insert spinner as a placeholder.
                var timestamp = new Date().getTime();
                var uploadid = 'moodleimage_' + Math.round(Math.random()*100000)+'-'+timestamp;

                host.focus();
                host.restoreSelection();
                var imagehtml = template({
                    url: M.util.image_url("i/loading_small", 'moodle'),
                    alt: M.util.get_string('uploading', COMPONENTNAME),
                    id: uploadid
                });
                host.insertContentAtFocusPoint(imagehtml);
                self.markUpdated();

                var xhr = new XMLHttpRequest();

                xhr.onreadystatechange = function() {
                    if (xhr.readyState === 4) {
                        var placeholder = self.editor.one('#' + uploadid);
                        if (xhr.status === 200) {
                            var result = JSON.parse(xhr.responseText);
                            if (result) {
                                if (result.error) {
                                    if (placeholder) {
                                        placeholder.remove(true);
                                    }
                                    return new M.core.ajaxException(result);
                                }

                                var file = result;

                                // Replace placeholder with actual image.
                                var filename = file.filename || file.file;
                                var newhtml = template({
                                    url: file.url,
                                    alt: filename
                                });
                                var newimage = Y.Node.create(newhtml);
                                if (placeholder) {
                                    placeholder.replace(newimage);
                                } else {
                                    self.editor.appendChild(newimage);
                                }
                                self.markUpdated();
                            }
                        } else {
                            alert(M.util.get_string('servererror', 'moodle'));
                            if (placeholder) {
                                placeholder.remove(true);
                            }
                        }
                    }
                };

                // Send the AJAX call.
                xhr.open("POST", M.cfg.wwwroot + '/repository/repository_ajax.php?action=upload', true);
                xhr.send(formData);
                return false;
            }
        }, this);
    }

});


}, '@VERSION@', {"requires": ["moodle-editor_atto-plugin"]});
