/**
 * Callback for rendering the main card.
 * @return {CardService.Card} The card to show the user.
 */
function onHomepage(e) {
  return createSelectionCard();
}

/**
 * Main function to generate the main card.
 * @return {CardService.Card} The card to show to the user.
 */
function createSelectionCard() {
  return CardService
    .newCardBuilder()
    .addSection(
      CardService.newCardSection()
        .setHeader('Document Actions')
        .addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText('The Works')
            .setOnClickAction(CardService.newAction().setFunctionName('formatForGalatea')))
          .addButton(CardService.newTextButton()
            .setText('Format Document')
            .setOnClickAction(CardService.newAction().setFunctionName('setDocumentFormatting')))
          .addButton(CardService.newTextButton()
            .setText('Fix Quotes')
            .setOnClickAction(CardService.newAction().setFunctionName('singleQuotesToDoubleQuotes')))
          .addButton(CardService.newTextButton()
            .setText('Wrap Italics')
            .setOnClickAction(CardService.newAction().setFunctionName('wrapItalicsWithTildes')))))
    .addSection(
      CardService.newCardSection()
        .setHeader('Targeted Actions')
        .addWidget(CardService.newDivider())
        .addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
            .setText('Episode')
            .setOnClickAction(CardService.newAction().setFunctionName('convertToEpisode')))
          .addButton(CardService.newTextButton()
            .setText('Series')
            .setOnClickAction(CardService.newAction().setFunctionName('convertToSeries')))))
    .build();
}

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

var episodeRegex = 'EPISODE: \\d+: ';
function convertToEpisode() {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    var element = cursor.getElement();
    element.asText().insertText(0, 'EPISODE: 0: ');
  }
  var ranges = getEpisodeRanges();

  console.log('ranges', ranges);
  for (var i = 0; i < ranges.length; i++) {
    var range = ranges[i];
    var textEl = range.getElement().asText();

    var correctEpisodeText = `EPISODE: ${i + 1}: `;
    if (textEl.findText(correctEpisodeText) == null) {
      textEl.replaceText(episodeRegex, correctEpisodeText);
    }

    var headingStyle = {};
    headingStyle[DocumentApp.Attribute.HEADING] = DocumentApp.ParagraphHeading.HEADING1;
    textEl.getParent().setAttributes(headingStyle);
  };

  setDocumentFormatting();
}

function getEpisodeRanges() {
  var body = DocumentApp.getActiveDocument().getBody();

  var ranges = [];
  var previousRange = null;
  var finished = false;
  while (!finished) {
    var match = null;
    if (previousRange == null) {
      match = body.findText(episodeRegex);
    } else {
      match = body.findText(episodeRegex, previousRange);
    }

    if (match == null || match == undefined) {
      finished = true;
    } else {
      ranges.push(match);
      previousRange = match;
    }
  }

  return ranges;
}

function convertToSeries() {
  var cursor = DocumentApp.getActiveDocument().getCursor();
  if (cursor) {
    var element = cursor.getElement();
    var textEl = element.asText();
    var oldTitle = textEl.getText();
    var newTitle = "SERIES:" + oldTitle.toLowerCase().replace(new RegExp("\\s", "g"), "_");
    textEl.setText(newTitle);
  }
}
