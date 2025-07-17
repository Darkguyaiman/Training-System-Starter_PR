function doGet() {
  return HtmlService.createHtmlOutputFromFile('HelpMe')
    .setTitle('Help Page')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}