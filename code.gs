function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ghibli Farm Sim - Final Phase')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
