function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Каразіна')
    .addItem('Use normal distribution', 'useNormDist')
    .addToUi();
}

function useNormDist() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var range = sheet.getActiveRange();
  var values = range.getValues();

  Logger.log('Input values: ' + JSON.stringify(values));

  var is_vertical = values.length > 1;
  var is_horisontal = values[0].length > 1;

  Logger.log('is_vertical: ' + is_vertical);
  Logger.log('is_horisontal: ' + is_horisontal);

  if (is_vertical) {
    if (values.length < 5) {
      ui.alert('Please select >=5 cells in a single column.');
      return;
    }
  }

  if (is_horisontal) {
    if (values[0].length < 5) {
      ui.alert('Please select >=5 cells in a single row.');
      return;
    }
  }

  var xValue = is_vertical ? values[0][0] : values[0][0];
  Logger.log('xValue: ' + xValue);
  var meanValue = is_vertical ? values[1][0] : values[0][1];
  Logger.log('meanValue: ' + meanValue);
  var stdDevValue = is_vertical ? values[2][0] : values[0][2];
  Logger.log('stdDevValue: ' + stdDevValue);
  var cumulativeValue = is_vertical ? values[3][0] : values[0][3];
  Logger.log('cumulativeValue: ' + cumulativeValue);

  if (xValue === '') {
    ui.alert(`Please select not empty 'x' value`);
    return;
  }

  if (typeof xValue !== 'number') {
    ui.alert(`The 'x' value should be number.`);
    return;
  }

  if (meanValue === '') {
    ui.alert(`Please select not empty 'mean' value`);
    return;
  }

  if (typeof meanValue !== 'number') {
    ui.alert(`The 'mean' value should be number.`);
    return;
  }

  if (stdDevValue === '') {
    ui.alert(`Please select not empty 'stdDev' value`);
    return;
  }

  if (typeof stdDevValue !== 'number') {
    ui.alert(`The 'stdDev' value should be number.`);
    return;
  }

  if (cumulativeValue === '') {
    ui.alert(`Please select not empty 'cumulative' value`);
    return;
  }

  if (typeof cumulativeValue !== 'boolean') {
    ui.alert(`The 'cumulative' value should be TRUE or FALSE.`);
    return;
  }

  if (is_vertical) {
    var lastRow = range.getLastRow();
    var column = range.getColumn();
    var lastCell = sheet.getRange(lastRow, column);
    lastCell.setFormula('=NORM.DIST(' + xValue + ',' + meanValue + ',' + stdDevValue + ',' + cumulativeValue + ')');
  }

  if (is_horisontal) {
    var lastCell = range.getCell(1, range.getNumColumns());
    lastCell.setFormula('=NORM.DIST(' + xValue + ',' + meanValue + ',' + stdDevValue + ',' + cumulativeValue + ')');
  }
}
