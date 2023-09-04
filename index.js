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
    rowHeaders: true,
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

exportExcel.addEventListener("click", () => {
  /* Create worksheet from HTML DOM TABLE */
  const wb = XLSX.utils.table_to_book(table, { sheet: "sheet-1" });

  /* Export to file (start a download) */
  XLSX.writeFile(wb, "MyTable.xlsx");
});
