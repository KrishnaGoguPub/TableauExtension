console.log("index.js loaded");

// Global variable for worksheet
let worksheet;

// Initialize the Tableau Extensions API
tableau.extensions.initializeAsync().then(() => {
  console.log("Tableau Extensions API initialized");
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  worksheet = dashboard.worksheets[0];
  console.log("Worksheet:", worksheet.name);
  loadTableauData();

  const refreshButton = document.getElementById("refreshButton");
  if (refreshButton) {
    refreshButton.addEventListener("click", () => {
      console.log("Refresh button clicked");

      // Step 1: Apply filters
      worksheet.getFiltersAsync()
        .then(filters => {
          if (filters.length > 0) {
            const filter = filters[0];
            console.log("Filter details:", JSON.stringify(filter, null, 2));

            if (filter.filterType === tableau.FilterType.CATEGORICAL && filter.values && Array.isArray(filter.values)) {
              return worksheet.applyFilterAsync(
                filter.fieldName,
                filter.values,
                tableau.FilterUpdateType.REPLACE
              ).then(() => console.log("Categorical filter reapplied"));
            } else if (filter.filterType === tableau.FilterType.RANGE && (filter.minValue !== undefined || filter.maxValue !== undefined)) {
              const rangeOptions = {};
              if (filter.minValue !== undefined) rangeOptions.min = filter.minValue;
              if (filter.maxValue !== undefined) rangeOptions.max = filter.maxValue;
              return worksheet.applyRangeFilterAsync(filter.fieldName, rangeOptions)
                .then(() => console.log("Range filter reapplied"));
            } else {
              console.log("Skipping unsupported or invalid filter type:", filter.filterType);
              return Promise.resolve(); // Continue without applying
            }
          } else {
            console.log("No filters found");
            return Promise.resolve();
          }
        })
        .then(() => {
          // Step 2: Apply parameter after filters
          return worksheet.getParametersAsync().then(params => {
            const param = params.find(p => p.name === "Selected Metric"); // Replace with your parameter name
            if (param) {
              console.log("Parameter details:", JSON.stringify(param, null, 2));
              const newValue = param.currentValue.value === "Sales" ? "Profit" : "Sales"; // Customize values
              return worksheet.changeParameterValueAsync(param.name, newValue)
                .then(() => console.log(`Parameter ${param.name} updated to ${newValue}`));
            } else {
              console.log("Parameter 'Selected Metric' not found");
              return Promise.resolve();
            }
          });
        })
        .then(() => {
          // Step 3: Refresh table after filters and parameter
          loadTableauData();
        })
        .catch(err => console.error("Error in refresh process:", err));
    });
  } else {
    console.error("Refresh button not found");
  }

  const exportButton = document.getElementById("exportButton");
  if (exportButton) {
    exportButton.addEventListener("click", () => {
      console.log("Export button clicked");
      worksheet.getSummaryDataAsync().then(data => {
        console.log("Data fetched for export");
        exportToExcel(data);
      }).catch(err => console.error("Error fetching data:", err));
    });
  } else {
    console.error("Export button not found");
  }
}).catch(err => console.error("Error initializing API:", err));

// Load data from Tableau worksheet
function loadTableauData() {
  if (!worksheet) {
    console.error("Worksheet not defined");
    return;
  }
  worksheet.getSummaryDataAsync().then(data => {
    console.log("Rendering table");
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

  const headers = data.columns.map(col => col.fieldName);
  wsData.push(headers);

  data.data.forEach(row => {
    const rowData = row.map(cell => cell.formattedValue || cell.value);
    wsData.push(rowData);
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Data cell formatting
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = 1; R <= range.e.r; ++R) {
    for (let C = 0; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[cellAddress];
      if (cell && cell.v !== undefined) {
        const cellValue = parseFloat(cell.v.toString().replace(/[^0-9.-]+/g, ""));
        if (!isNaN(cellValue)) {
          cell.s = cell.s || {};
          if (cellValue > 1000) {
            cell.s.fill = { patternType: "solid", fgColor: { rgb: "FFCCCC" } };
            console.log(`Applied pink to ${cellAddress}: ${cellValue}`);
          } else if (cellValue > 500) {
            cell.s.fill = { patternType: "solid", fgColor: { rgb: "FFFFCC" } };
            console.log(`Applied yellow to ${cellAddress}: ${cellValue}`);
          }
        }
      }
    }
  }

  // Header colors (customize to match your Tableau view)
  for (let C = 0; C <= range.e.c; ++C) {
    const headerCell = XLSX.utils.encode_cell({ r: 0, c: C });
    const column = data.columns[C];
    let headerColor;

    switch (column.fieldName.toLowerCase()) { // Customize here
      case "sales": headerColor = "D3D3D3"; break; // Darker gray
      case "profit": headerColor = "CCFFCC"; break; // Light green
      case "quantity": headerColor = "CCE5FF"; break; // Light blue
      default: headerColor = "F2F2F2"; // Default light gray
    }

    ws[headerCell].s = {
      font: { bold: true },
      fill: { patternType: "solid", fgColor: { rgb: headerColor } },
      border: { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" 
