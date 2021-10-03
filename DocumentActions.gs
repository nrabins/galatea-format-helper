function formatForGalatea() {
  setDocumentFormatting();
  singleQuotesToDoubleQuotes();
  wrapItalicsWithTildes();
}

function setDocumentFormatting() {
  var font = {};
  font[DocumentApp.Attribute.FONT_FAMILY] = 'Times New Roman';
  font[DocumentApp.Attribute.FONT_SIZE] = 12;
  font[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;

  var body = DocumentApp.getActiveDocument().getBody();
  body.setAttributes(font);

  body.getParagraphs().forEach(p => p.setAttributes(font));
}

/* 
  Finds quotes adjacent to non-word boundaries and transforms them to double quotes.
  This finds quotation instances, but does run into false positives when dealing with
  dialectical contractions (e.g. "top o' the mornin' to ya")
*/
function singleQuotesToDoubleQuotes() {
  var boundaryQuoteRegex = "\\B['‘’]|['‘’]\\B";
  var body = DocumentApp.getActiveDocument().getBody();
  body.replaceText(boundaryQuoteRegex, "\"");
}

function wrapItalicsWithTildes() {
  var body = DocumentApp.getActiveDocument().getBody();

  var paragraphs = body.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraph = paragraphs[i];
    for (var j = 0; j < paragraph.getNumChildren(); j++) {
      var element = paragraph.getChild(j);
      if (element.getType() == DocumentApp.ElementType.TEXT) {
        var text = element.asText();
        var textAttributeIndices = text.getTextAttributeIndices();

        // Find contiguous italic portions
        if (textAttributeIndices.length == 0) {
          continue;
        }

        var italicChunks = [];

        var italicStart = null;
        textAttributeIndices.forEach(index => {
          var isItalic = text.getAttributes(index)['ITALIC'];
          if (isItalic) {
            if (italicStart == null) {
              // This is the start of a new italic section
              italicStart = index;
            }
          } else {
            // The character is not italic
            if (italicStart != null) {
              // We have a start, add an end and close the section
              italicChunks.push({ start: italicStart, end: index });
              italicStart = null;
            }
          }
        });

        if (italicStart != null) {
          // We ended on an italic, add the final one
          italicChunks.push({ start: italicStart, end: text.getText().length });
        }

        // Work from the end to the front to avoid having to mess with shifting text size as we insert tildes
        for (var z = italicChunks.length - 1; z >= 0; z--) {
          var italicChunk = italicChunks[z];
          text.insertText(italicChunk.end, '~');
          text.insertText(italicChunk.start, '~');
        }
      }
    }
  }
}