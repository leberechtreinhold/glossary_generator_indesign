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

#targetengine write_text_anchors;

(function () {
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
            if (current_number === "") {
                var page_name = "?";
                var textframes = paragraph.parentTextFrames;
                for (var j = 0; j < textframes.length; j++) {
                    if (textframes[j].isValid) {
                        page_name += "p" + textframes[j].parentPage.name + ",";
                        current_order = parseInt(textframes[j].parentPage.name)
                    }
                }
                current_number = trim_char(page_name, ",") + "?";
            } else {
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
            }


            var original_name = trim_char(hyperlink_text_destinations[i].name, ' ');
            var name = trim_char(original_name, '_');

            if (!anchors[name]) {
                anchors[name] = [];
            }
            anchors[name].push({
                name: original_name,
                order: current_order,
                number: current_number
            });
        }
        return anchors;
    }

    // Given all the anchors corresponding with a single name in an array, returns a line
    // corresponding with the entry
    function get_anchor_entry(anchor_name, anchors) {
        var numbers = "";
        anchors.sort(function (a, b) {
            return a.order - b.order;
        });
        for (var i = 0; i < anchors.length; i++) {
            numbers += anchors[i].number + " & ";
        }
        numbers = trim_chars_ordered(numbers, [' ', '&', ' '])
        return anchor_name + ": " + numbers;
    }

    function anchors_to_text(anchors) {
        var final_text = "";

        var anchor_names = get_keys(anchors);
        anchor_names.sort();

        for (var i = 0; i < anchor_names.length; i++) {
            var anchor_name = anchor_names[i];
            final_text += get_anchor_entry(anchor_name, anchors[anchor_name]) + "\r";
        }

        return final_text;
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

        var anchors = get_anchors();

        selected_obj.contents = anchors_to_text(anchors);
    }

    if (app.documents.length > 0) {
        main();
    }

}());