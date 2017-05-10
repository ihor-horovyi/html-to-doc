(function ($) {
    $.fn.extend({
        exportToDoc: function (options) {
            var defaultOptions = {
                ignoresTag: [],
                tableName: 'yourTableName',
                type: 'doc',
                escape: 'true',
                htmlContent: 'false'
            };

            var options = $.extend(defaultOptions, options);

            var docFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:" + defaultOptions.type + "' xmlns='http://www.w3.org/TR/REC-html40'>";
            docFile += "<head></head>";
            docFile += "<body>" + parseHtmlElement(this, options.ignoresTag) + "</body>";
            docFile += "</html>";

            var base64data = "base64," + $.base64.encode(docFile);
            window.open('data:application/vnd.ms-' + defaultOptions.type + ';filename=exportData.doc;' + base64data);

            function parseHtmlElement(htmlElement, ignoresTag) {
                var tags = htmlElement.children();
                var parsedContent = "";

                for (var i = 0; i < tags.length; i++) {
                    if (tags[i].tagName.toUpperCase() == "TABLE" && !isIgnoreTag(ignoresTag, "TABLE")) {
                        parsedContent += parseTable(tags[i], ignoresTag);
                    } else if (!isIgnoreTag(ignoresTag, tags[i].tagName)) {
                        parsedContent += parseTag(tags[i]);
                    }
                }

                return parsedContent;
            }

            function parseTable(el, ignoresTag) {
                var table = "<table>";

                table += parseTableTag(el, "thead", ignoresTag);
                table += parseTableTag(el, "tbody", ignoresTag);
                table += parseTableTag(el, "tfoot", ignoresTag);
                table += '</table>'
                return table;
            }

            function parseTableTag(el, tagName, ignoresTag) {
            	var parsedTag = "";
            	if (!isIgnoreTag(ignoresTag, tagName)) {
	            	$(el).find(tagName).find('tr').each(function () {
	                    parsedTag += "<tr>";
	                    $(this).filter(':visible').find('td').each(function (index, data) {
	                        if ($(this).css('display') != 'none') {
	                        	parsedTag += "<td>" + parseString($(this)) + "</td>";
	                        }
	                    });
	                    parsedTag += '</tr>';
	                });
            	}

                return parsedTag;
            }

            function parseTag(tag) {
                var parseTag = "<" + tag.tagName + ">";
                parseTag += tag.textContent;
                parseTag += "</" + tag.tagName + ">";
                return parseTag;
            }

            function isIgnoreTag(ignoresTag, tagName) {
            	for (var index in ignoresTag) {
            		if (ignoresTag[index].toUpperCase() === tagName.toUpperCase()) {
            			return true;
            		}
            	}
            	return false;
            }

            function parseString(data) {
                if (defaultOptions.htmlContent == 'true') {
                    content_data = data.html().trim();
                } else {
                    content_data = data.text().trim();
                }

                if (defaultOptions.escape == 'true') {
                    content_data = escape(content_data);
                }

                return content_data;
            }
        }
    });
})(jQuery);