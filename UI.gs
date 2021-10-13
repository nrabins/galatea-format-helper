/**
 * Main function to generate the main card.
 * @return {CardService.Card} The card to show to the user.
 */
function createMainCard() {
  return CardService
    .newCardBuilder()
    .addSection(buildOneOffActionsSection())
    .addSection(buildEpisodesSection())
    .build();
}

function buildOneOffActionsSection() {
  return CardService.newCardSection()
    .setCollapsible(true)
    .setHeader('One-Off Actions')
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
        .setOnClickAction(CardService.newAction().setFunctionName('wrapItalicsWithTildes')))
      .addButton(CardService.newTextButton()
        .setText('Mark Series')
        .setOnClickAction(CardService.newAction().setFunctionName('convertToSeries'))));
}

function buildEpisodesSection() {
  return CardService.newCardSection()
    .setHeader('Episodes')
    .addWidget(CardService.newTextInput()
      .setFieldName('episodeStartFrom')
      .setTitle('Start From')
      .setValue('1'))
    .addWidget(CardService.newDivider())
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Mark Episode')
        .setOnClickAction(CardService.newAction().setFunctionName('convertToEpisode'))));
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