// Initialize the Tableau Extensions API
tableau.extensions.initializeAsync().then(() => {
  console.log("Tableau Extensions API initialized");
  loadTableauData(); // Load data on startup
}).catch(err => {
  console.error("Error initializing Tableau Extensions API:", err);
});

// Global variables
let worksheet;

// Load data from Tableau worksheet
function loadTableauData() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  worksheet = dashboard.worksheets[0]; // Use the first worksheet; adjust as needed
  worksheet.getSummaryDataAsync().then(data => {
    renderTable(data);
  }).catch(err => {
    console.error("Error fetching data:", err);
  });
}

// Render the data into the HTML table
function renderTable(data) {
  const tableHeader = document.getElementById("tableHeader").querySelector("tr");
  const tableBody = document.getElementById("tableBody");

  // Clear existing content
  tableHeader.innerHTML = "";
  tableBody.innerHTML = "";

  // Populate headers
  data.columns.forEach(column => {
    const th = document.createElement("th");
    th.textContent = column.fieldName;
    tableHeader.appendChild(th);
  });

  // Populate rows with data and apply conditional formatting
  data.data.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach((cell, index) => {
      const td = document.createElement("td");
      td.textContent = cell.formattedValue || cell.value;
      
      // Example: Apply color based on value (customize this logic)
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

// Refresh button: Apply filters and reload data
document.getElementById("refreshButton").addEventListener("click", () => {
  worksheet.applyFilterAsync().then(() => {
    loadTableauData();
  }).catch(err => {
    console.error("Error applying filters:", err);
  });
});

// Export button: Export table to Excel with formatting
document.getElementById("exportButton").addEventListener("click", () => {
  worksheet.getSummaryDataAsync().then(data => {
    exportToExcel(data);
  }).catch(err => {
    console.error("Error fetching data for export:", err);
  });
});

// Export to Excel with formatting
function exportToExcel(data) {
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

  // Apply formatting (e.g., colors based on values)
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = 1; R <= range.e.r; ++R) { // Skip header row (R=0)
    for (let C = 0; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cellValue = parseFloat(ws[cellAddress]?.v);

      if (!isNaN(cellValue)) {
        ws[cellAddress].s = ws[cellAddress].s || {}; // Initialize style object
        if (cellValue > 1000) {
          ws[cellAddress].s.fill = { fgColor: { rgb: "FFCCCC" } }; // Light red
        } else if (cellValue > 500) {
          ws[cellAddress].s.fill = { fgColor: { rgb: "FFFFCC" } }; // Light yellow
        }
      }
    }
  }

  // Add styling to headers
  for (let C = 0; C <= range.e.c; ++C) {
    const headerCell = XLSX.utils.encode_cell({ r: 0, c: C });
    ws[headerCell].s = {
      font: { bold: true },
      fill: { fgColor: { rgb: "F2F2F2" } }, // Light gray background
      border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } }
    };
  }

  // Set column widths (approximate)
  ws["!cols"] = headers.map(() => ({ wpx: 100 }));

  // Append worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, "TableauExport");

  // Export the file
  XLSX.writeFile(wb, "TableauViewExport.xlsx");
}

// Utility function for binary string (if needed for Blob export)
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}
