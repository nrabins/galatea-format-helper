/**
 * Main function to generate the main card.
 * @return {CardService.Card} The card to show to the user.
 */
function createMainCard() {
  return CardService
    .newCardBuilder()
    .addSection(buildDocumentActionsSection())
    .addSection(buildTargetedActionsSection())
    .addSection(buildParagraphWordCountSection())
    .build();
}

function buildDocumentActionsSection() {
  return CardService.newCardSection()
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
        .setOnClickAction(CardService.newAction().setFunctionName('wrapItalicsWithTildes'))));
}

function buildTargetedActionsSection() {
  return CardService.newCardSection()
    .setHeader('Targeted Actions')
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Episode')
        .setOnClickAction(CardService.newAction().setFunctionName('convertToEpisode')))
      .addButton(CardService.newTextButton()
        .setText('Series')
        .setOnClickAction(CardService.newAction().setFunctionName('convertToSeries'))));
}

function buildParagraphWordCountSection() {
  const section = CardService.newCardSection()
    .setHeader('Paragraph Counts')
    .addWidget(CardService.newTextButton()
      .setText('Refresh')
      .setOnClickAction(CardService.newAction().setFunctionName('testChapterWordCounts')))
    .addWidget(CardService.newTextParagraph().setText(Math.random().toString()));
  return section;
}