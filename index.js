const container = document.querySelector("#example");
const button = document.querySelector("#export-file");
const showCell = document.getElementById("show-cell");
const cellPosition = document.getElementById("cell-position");
const exportExcel = document.getElementById("export-excel");
const hot = new Handsontable(container, {
  startRows: 10,
  startCols: 10,
  width: "auto",
  height: "auto",
  dropdownMenu: true,
  rowHeaders: true,
  colHeaders: true,
  selectionMode: "multiple",
  licenseKey: "non-commercial-and-evaluation", // for non-commercial use only
});

const exportPlugin = hot.getPlugin("exportFile");

button.addEventListener("click", () => {
  exportPlugin.downloadFile("csv", {
    bom: false,
    columnDelimiter: ",",
    columnHeaders: false,
    exportHiddenColumns: true,
    exportHiddenRows: true,
    fileExtension: "csv",
    filename: "Handsontable-CSV-file_[YYYY]-[MM]-[DD]",
    mimeType: "text/csv",
    rowDelimiter: "\r\n",
  });
});

const table = document.getElementsByTagName("table")[0];
const cells = table.getElementsByTagName("td");

for (let i = 0; i < cells.length; i++) {
  let cell = cells[i];
  cell.onclick = function (e) {
    let cellValue = e.target.innerHTML;
    if (cellValue == "") {
      cellValue = "Empty";
    }
    let cellIndex = this.cellIndex;
    let rowIndex = this.parentNode.rowIndex;
    cellPosition.innerText = `${cellValue} Cell is in ${cellIndex}, ${rowIndex}.`;
  };
}

const rows = 10;
const cols = 10;

function exportToExcel() {
  for (let i = 1; i <= rows; i++) {
    for (let j = 1; j < cols; j++) {
      const input = table.rows[i].cells[j].querySelector("input");
      if (input) {
        table.rows[i].cells[j].innerText = input.value;
      }
    }
  }
  // table을 clonedTable에 복사
  const clonedTable = table.cloneNode(true);

  // 위쪽 헤더 제거
  clonedTable.deleteRow(0);
  // 왼쪽 헤더 제거
  for (let i = 0; i < rows; i++) {
    clonedTable.rows[i].deleteCell(0);
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(clonedTable);
  XLSX.utils.book_append_sheet(wb, ws, "sheet1");
  XLSX.writeFile(wb, "spreadsheet.xlsx");
}

exportExcel.addEventListener("click", exportToExcel);
