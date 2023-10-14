// MIT License

// Copyright (c) 2023 Aitor Jiménez Yañez

// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:

// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

//@target indesign

(function () {
    // Parameters -----------------------------------------------------------------------------------------------------
    
    CHARACTER_STYLE_TO_USE = "HyperlinkGlossary";
    USE_BULLET_NUMBER_INSTEAD_OF_PAGE = true;

    // Utilities ------------------------------------------------------------------------------------------------------

    // Utility function to clean a string
    // https://stackoverflow.com/questions/26156292/trim-specific-character-from-a-string
    function trim_char(str, ch) {
        var start = 0;
        var end = str.length;

        while (start < end && str[start] === ch)
            ++start;

        while (end > start && str[end - 1] === ch)
            --end;

        return (start > 0 || end < str.length) ? str.substring(start, end) : str;
    }

    // Since I'm a lazy slob, instead of adapting the above to accept multiple chars
    // I just concat the calls
    function trim_chars_ordered(str, ch_arr) {
        for (var i = 0; i < ch_arr.length; i++) {
            str = trim_char(str, ch_arr[i])
        }
        return str;
    }

    // Since ExtendedScript uses ECMAScript3, we need this because we cannot do Object.keys(obj)
    function get_keys(obj) {
        var keys = [];
        for (var key in obj) {
            keys.push(key);
        }
        return keys;
    }

    // Functionality --------------------------------------------------------------------------------------------------

    // Returns a map of anchors per name
    // {
    //     "ANCHORNAME" = [
    //         {
    //             "name" = "ANCHORNAME",
    //             "number" = "1.2.3",
    //             "order" = 10203
    //         }
    //         There may be multiple entries per the same Anchor!
    //     ]
    // }
    function get_anchors() {
        var anchors = {};
        var hyperlink_text_destinations = app.documents[0].hyperlinkTextDestinations.everyItem().getElements();
        for (var i = 0; i < hyperlink_text_destinations.length; i++) {
            if (!(hyperlink_text_destinations[i] instanceof HyperlinkTextDestination)) {
                continue; // There may be other types
            }

            var paragraphs = hyperlink_text_destinations[i].destinationText.paragraphs;
            if (paragraphs.length === 0) {
                continue; // We only want those that point to an actual paragraph
            }


            var paragraph = paragraphs.firstItem();
            var current_number = paragraph.bulletsAndNumberingResultText;
            var current_order = 0;
            
            if (current_number != "" && USE_BULLET_NUMBER_INSTEAD_OF_PAGE) {
                current_number = trim_char(page_name, ",");
                current_number = trim_chars_ordered(current_number, [' ', '.'])
                var parts = current_number.split('.');

                // We only care for the first 4 values, it may have more but they are not counted... Each can have up to 99
                if (parts.length >= 1) {
                    current_order += parseInt(parts[0]) * 10 * 10 * 10 * 10 * 10 * 10 * 10;
                }
                if (parts.length >= 2) {
                    current_order += parseInt(parts[1]) * 10 * 10 * 10 * 10;
                }
                if (parts.length >= 3) {
                    current_order += parseInt(parts[2]) * 10 * 10;
                }
                if (parts.length >= 4) {
                    current_order += parseInt(parts[3]);
                }
            } else {
                var page_name = "";
                var textframes = paragraph.parentTextFrames;
                for (var j = 0; j < textframes.length; j++) {
                    if (textframes[j].isValid) {
                        page_name += "p" + textframes[j].parentPage.name + ",";
                        current_order = parseInt(textframes[j].parentPage.name)
                    }
                }
            }

            var original_name = trim_char(hyperlink_text_destinations[i].name, ' ');
            var name = trim_char(original_name, '_');

            if (!anchors[name]) {
                anchors[name] = [];
            }
            anchors[name].push({
                name: original_name,
                order: current_order,
                number: current_number,
                destination: hyperlink_text_destinations[i]
            });
        }
        return anchors;
    }

    // Given all the anchors corresponding with a single name in an array, returns a line
    // corresponding with the entry
    function add_references_to_glossary_entry(story, anchors, hyperlink_style) {
        var numbers = "";
        anchors.sort(function (a, b) {
            return a.order - b.order;
        });
        for (var i = 0; i < anchors.length; i++) {
            var num_as_str = "" + anchors[i].number; // Trick to convert number to str: ""+13="13"
            story.contents += num_as_str;
            
            var last_char_position = story.insertionPoints.length - 1;
            var start_of_last_num = last_char_position - num_as_str.length;
            var chars_story = story.insertionPoints.itemByRange(start_of_last_num, last_char_position);
            
            var previous_style = story.appliedCharacterStyle;
            story.contents += " ";

            var hyperlink_source = app.activeDocument.hyperlinkTextSources.add(chars_story);
            hyperlink_source.appliedCharacterStyle = hyperlink_style;
            app.activeDocument.hyperlinks.add(hyperlink_source, anchors[i].destination);

            story.characters[-1].appliedCharacterStyle = previous_style;

            if (i < (anchors.length - 1)) {
                story.contents += "& ";
            }
        }
    }

    function add_glossary_entries_to_story(story, anchors, hyperlink_style) {
        var anchor_names = get_keys(anchors);
        anchor_names.sort();

        for (var i = 0; i < anchor_names.length; i++) {
            var anchor_name = anchor_names[i];
            story.contents += anchor_name + ": ";
            add_references_to_glossary_entry(story, anchors[anchor_name], hyperlink_style);            
            story.contents += "\r";
        }
    }

    function main() {
        var current_selection = app.selection;
        if (current_selection.length != 1) {
            alert("Nothing selected or too much, only 1 text frame needs to be selected");
            return;
        }
        var selected_obj = current_selection[0];
        if (!(selected_obj instanceof TextFrame)) {
            alert("The selected element is not a textframe");
            return;
        }

        var styles = app.activeDocument.characterStyles.everyItem().getElements();
        var hyperlink_style_to_use = null;
        for(var i=0; i<styles.length; i++) {
            if (styles[i].name == CHARACTER_STYLE_TO_USE) {
                hyperlink_style_to_use = styles[i];
                break;
            }
        }

        if (!hyperlink_style_to_use) {
            alert("The Character style " + CHARACTER_STYLE_TO_USE + " was not found in this doc.");
            return;
        }

        var anchors = get_anchors();
        add_glossary_entries_to_story(selected_obj.parentStory, anchors, hyperlink_style_to_use);
    }

    if (app.documents.length > 0) {
        main();
    }

}());