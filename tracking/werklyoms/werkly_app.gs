function doGet(e) {
  try {
    const context = buildOrderTemplateContext_();
    const template = HtmlService.createTemplateFromFile('werkly_index');
    template.allOrdersJson = context.allOrdersJson;
    template.maintenanceOrdersJson = context.maintenanceOrdersJson;
    return template.evaluate();
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.message);
  }
}

function buildOrderTemplateContext_() {
  const dataset = loadOrderDataset_();
  return {
    allOrdersJson: stringifyForHtml_(dataset.allOrders),
    maintenanceOrdersJson: stringifyForHtml_(dataset.maintenanceOrders)
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
