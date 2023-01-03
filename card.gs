function onHomePage(e) {
  var cards = [];
  onOpen(e);
  cards.push(displayFileId(e));
  return cards;
}

function displayFileId(e) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('How to use?'));
  var section = CardService.newCardSection();
  var idtxt = CardService.newTextParagraph().setText('Click on the Menu > Estensioni > Formula Updater [Addon]');
  section.addWidget(idtxt);
  card.addSection(section);
  return card.build();
}