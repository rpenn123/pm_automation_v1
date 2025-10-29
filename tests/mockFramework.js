// tests/mockFramework.js

global.MockRange = class MockRange {
  constructor(data, sheet, row, col) {
    this.data = data;
    this.sheet = sheet;
    this.row = row;
    this.col = col;
  }
  getValues() { return this.data; }
  getValue() { return this.data[0] ? this.data[0][0] : ""; }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  setBackground() { return this; } // Chainable
  setFontColor() { return this; } // Chainable

  setValues(values) {
    // This is a simplified implementation for the tests
    const startRow = this.row - 1;
    const startCol = this.col - 1;
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (!this.sheet.data[startRow + i]) {
          this.sheet.data[startRow + i] = [];
        }
        this.sheet.data[startRow + i][startCol + j] = values[i][j];
      }
    }
  }
}

global.MockSheet = class MockSheet {
  constructor(name = "Sheet1") {
    this.name = name;
    this.data = [];
  }
  getName() { return this.name; }
  appendRow(rowData) {
      // Ensure row is wide enough
      const maxCols = this.getMaxColumns();
      if (rowData.length < maxCols) {
          rowData.push(...new Array(maxCols - rowData.length).fill(""));
      }
      this.data.push(rowData);
  }
  getDataRange() { return new global.MockRange(this.data, this, 1, 1); }
  getRange(row, col, numRows = 1, numCols = 1) {
    const rangeData = [];
    for (let r = 0; r < numRows; r++) {
        const currentRow = this.data[row - 1 + r] || [];
        rangeData.push(currentRow.slice(col - 1, col - 1 + numCols));
    }
    return new global.MockRange(rangeData, this, row, col);
  }
  getMaxColumns() {
      return this.data.reduce((max, row) => Math.max(max, row.length), 0);
  }
  getLastRow() { return this.data.length; }
}

global.MockSpreadsheet = class MockSpreadsheet {
  constructor(sheets = []) {
    this.sheets = sheets;
    this.id = "test-ss-id";
  }
  getId() { return this.id; }
  getSheetByName(name) {
    let sheet = this.sheets.find(s => s.getName() === name);
    if (!sheet) {
        sheet = new global.MockSheet(name);
        this.sheets.push(sheet);
    }
    return sheet;
  }
  getName() { return "Test Spreadsheet"; }
}

global.MockSpreadsheetApp = class MockSpreadsheetApp {
  constructor(spreadsheet) {
    this.spreadsheet = spreadsheet || new global.MockSpreadsheet();
  }
  getActiveSpreadsheet() { return this.spreadsheet; }
}
