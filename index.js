console.log("index.js loaded");

// Global variable for worksheet
let worksheet;

// Initialize the Tableau Extensions API
tableau.extensions.initializeAsync().then(() => {
  console.log("Tableau Extensions API initialized");
  
  // Set worksheet and load data
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  worksheet = dashboard.worksheets[0]; // Use the first worksheet; adjust as needed
  console.log("Worksheet set:", worksheet.name);
  loadTableauData();

  // Set up event listeners after DOM and API are ready
  const refreshButton = document.getElementById("refreshButton");
  if (refreshButton) {
    refreshButton.addEventListener("click", () => {
      console.log("Refresh button clicked");
      worksheet.applyFilterAsync().then(() => {
        console.log("Filters applied, reloading data");
        loadTableauData();
      }).catch(err => console.error("Error applying filters:", err));
    });
  } else {
    console.error("Refresh button not found in DOM");
  }

  const exportButton = document.getElementById("exportButton");
  if (exportButton) {
    exportButton.addEventListener("click", () => {
      console.log("Export button clicked");
      worksheet.getSummaryDataAsync().then(data => {
        console.log("Data fetched for export");
        exportToExcel(data);
      }).catch(err => console.error("Error fetching data for export:", err));
    });
  } else {
    console.error("Export button not found in DOM");
  }
}).catch(err => console.error("Error initializing Tableau Extensions API:", err));

// Load data from Tableau worksheet
function loadTableauData() {
  if (!worksheet) {
    console.error("Worksheet not defined");
    return;
  }
  worksheet.getSummaryDataAsync().then(data => {
    console.log("Rendering table with data");
    renderTable(data);
  }).catch(err => console.error("Error fetching data:", err));
}

// Render the data into the HTML table
function renderTable(data) {
  const tableHeader = document.getElementById("tableHeader").querySelector("tr");
  const tableBody = document.getElementById("tableBody");

  tableHeader.innerHTML = "";
  tableBody.innerHTML = "";

  data.columns.forEach(column => {
    const th = document.createElement("th");
    th.textContent = column.fieldName;
    tableHeader.appendChild(th);
  });

  data.data.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell.formattedValue || cell.value;
      const value = parseFloat(cell.value);
      if (!isNaN(value)) {
        if (value > 1000) td.style.backgroundColor = "#ffcccc"; // Light red
        else if (value > 500) td.style.backgroundColor = "#ffffcc"; // Light yellow
      }
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}

// Export to Excel with formatting and header colors
function exportToExcel(data) {
  console.log("Exporting to Excel...");
  const wb = XLSX.utils.book_new();
  const wsData = [];

  // Add headers
  const headers = data.columns.map(col => col.fieldName);
  wsData.push(headers);

  // Add rows with values
  data.data.forEach(row => {
    const rowData = row.map(cell => cell.formattedValue || cell.value);
    wsData.push(rowData);
  });

  // Create worksheet
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Apply formatting to data cells (pink and yellow)
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = 1; R <= range.e.r; ++R) {
    for (let C = 0; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cellValue = parseFloat(ws[cellAddress]?.v);
      if (!isNaN(cellValue)) {
        ws[cellAddress].s = ws[cellAddress].s || {};
        if (cellValue > 1000) {
          ws[cellAddress].s.fill = { fgColor: { rgb: "FFCCCC" } }; // Light red
        } else if (cellValue > 500) {
          ws[cellAddress].s.fill = { fgColor: { rgb: "FFFFCC" } }; // Light yellow
        }
      }
    }
  }

  // Add styling to headers with colors from Tableau view (customize here)
  for (let C = 0; C <= range.e.c; ++C) {
    const headerCell = XLSX.utils.encode_cell({ r: 0, c: C });
    const column = data.columns[C];
    let headerColor;

    // Define header colors based on your Tableau view
    switch (column.fieldName.toLowerCase()) { // Adjust to match your viz
      case "sales":
        headerColor = "D3D3D3"; // Darker gray
        break;
      case "profit":
        headerColor = "CCFFCC"; // Light green
        break;
      case "quantity":
        headerColor = "CCE5FF"; // Light blue
        break;
      default:
        headerColor = "F2F2F2"; // Default light gray
    }

    ws[headerCell].s = {
      font: { bold: true },
      fill: { fgColor: { rgb: headerColor } },
      border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
    };
  }

  // Set column widths
  ws["!cols"] = headers.map(() => ({ wpx: 100 }));

  // Append worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, "TableauExport");

  // Export the file (Blob workaround for Tableau Desktop)
  const fileData = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
  const blob = new Blob([s2ab(fileData)], { type: "application/octet-stream" });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "TableauViewExport.xlsx";
  console.log("Triggering download...");
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
  console.log("Download complete");
}

// Utility function for binary string
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}
