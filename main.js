const COL_COMMISSION = "佣金"
const COL_COG = "單件COG"
const COL_SETTLEMENT_PRICE = "結算價"
const COL_ORDER_STATUS = "訂單狀態"

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

function main() {
  const sheet = SpreadsheetApp.getActiveSheet()

  // Add text [佣金] [單件COG] [結算價] at AQ2; AR2; AS2
  appendColumns(sheet, [COL_COMMISSION, COL_COG, COL_SETTLEMENT_PRICE])

  // Base on column O data, remove roll with 已取消(自動) 已取消(主動) in column O 
  deleteOrderByStatus(sheet, OrderStatus.AUTO_CANCELLED)
  deleteOrderByStatus(sheet, OrderStatus.MANUAL_CANCELLED)
}