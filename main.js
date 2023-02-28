const COL_COMMISSION = "佣金"
const COL_COG = "單件COG"
const COL_SETTLEMENT_PRICE = "結算價"
const COL_ORDER_STATUS = "訂單狀態"

const OrderSheetName = "工作表1"
const MapMerchantSheetName = "MAP Merchant"
const MapCogSheetName = "MAP COG"

const OrderStatus = {
  MANUAL_CANCELLED: "已取消(主動)",
  AUTO_CANCELLED: "已取消(自動)",
  FINISHED: "已完成",
  PENDING_FOR_TAKEN: "待取貨",
  PENDING_FOR_SENT: "待發貨"
}

function appendColumns(sheet, columnNames = []) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSheet()
  }

  if (columnNames.length <= 0) {
    return;
  }

  const lastCol = sheet.getLastColumn();

  columnNames.forEach((name, i) => {
    const range = sheet.getRange(1, lastCol + i + 1)
    range.setValue(name)
  })
}

function deleteOrderByStatus(sheet, status) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSheet();
  }

  const data = sheet.getDataRange().getValues()
  const firstRow = data[0]
  let invoiceStatusCellIdx = null;

  for (let i = 0; i < firstRow.length; i++) {
    if (firstRow[i] === COL_ORDER_STATUS) {
      invoiceStatusCellIdx = i
      break
    }
  }

  if (invoiceStatusCellIdx === null) {
    const msg = `Column ${COL_ORDER_STATUS} not found`
    throw new Error(msg)
  }

  let rowsDeleted = 0;

  for (let i = 0; i < data.length; i++) {
    const orderStatus = data[i][invoiceStatusCellIdx]

    if (orderStatus === status) {

      sheet.deleteRow(i + 1 - rowsDeleted)
      rowsDeleted++;
    }
  }
}

function sortInvoiceByStatus(sheet, status) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSheet();
  }
}

function createSheet(sheetName) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadSheet.getSheetByName(sheetName)

  if (sheet) {
    return;
  }

  spreadSheet.insertSheet(sheetName)
}

function main() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = spreadsheet.getSheetByName(OrderSheetName)

  // Add text [佣金] [單件COG] [結算價] at AQ2; AR2; AS2
  appendColumns(orderSheet, [COL_COMMISSION, COL_COG, COL_SETTLEMENT_PRICE])

  // Base on column O data, remove roll with 已取消(自動) 已取消(主動) in column O 
  deleteOrderByStatus(orderSheet, OrderStatus.AUTO_CANCELLED)
  deleteOrderByStatus(orderSheet, OrderStatus.MANUAL_CANCELLED)

  // Create two new tab with name [MAP Merchant] [MAP COG]
  createSheet(MapMerchantSheetName)
  createSheet(MapCogSheetName)
}