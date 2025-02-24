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

      // Apply filters safely
      worksheet.getFiltersAsync().then(filters => {
        if (filters.length > 0) {
          const filter = filters[0];
          console.log("Filter details:", JSON.stringify(filter, null, 2));

          if (filter.filterType === tableau.FilterType.CATEGORICAL && filter.values && Array.isArray(filter.values)) {
            worksheet.applyFilterAsync(
              filter.fieldName,
              filter.values,
              tableau.FilterUpdateType.REPLACE
            ).then(() => console.log("Categorical filter reapplied"));
          } else if (filter.filterType === tableau.FilterType.RANGE) {
            const rangeOptions = {};
            if (filter.minValue !== undefined) rangeOptions.min = filter.minValue;
            if (filter.maxValue !== undefined) rangeOptions.max = filter.maxValue;
            worksheet.applyRangeFilterAsync(filter.fieldName, rangeOptions)
              .then(() => console.log("Range filter reapplied"));
          } else {
            console.log("Unsupported filter type:", filter.filterType);
          }
        } else {
          console.log("No filters found");
        }
      }).catch(err => console.error("Error applying filters:", err));

      // Apply parameter (example: "Selected Metric")
      worksheet.getParametersAsync().then(params => {
        const param = params.find(p => p.name === "Selected Metric"); // Replace with your parameter name
        if (param) {
          console.log("Parameter details:", JSON.stringify(param, null, 2));
          // Example: Toggle or set a value (adjust to your parameter’s allowable values)
          const newValue = param.currentValue.value === "Sales" ? "Profit" : "Sales";
          worksheet.changeParameterValueAsync(param.name, newValue).then(() => {
            console.log(`Parameter ${param.name} updated to ${newValue}`);
            loadTableauData(); // Refresh table after parameter change
          }).catch(err => console.error("Error updating parameter:", err));
        } else {
          console.log("Parameter 'Selected Metric' not found; refreshing data only");
          loadTableauData();
        }
      }).catch(err => console.error("Error getting parameters:", err));
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
        const cellValu
