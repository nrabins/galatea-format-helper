const episodeRegex = 'EPISODE: \\d+: ';

function convertToEpisode(e) {

  console.log(e);

  let startFrom = parseInt(e.formInputs.episodeStartFrom || 1);
  if (isNaN(startFrom)) {
    startFrom = 1;
  }

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

    var correctEpisodeText = `EPISODE: ${i + startFrom}: `;
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