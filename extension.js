tableau.extensions.dashboardContent.dashboard.worksheets[0]
  .getUnderlyingTableDataAsync()
  .then(dataTable => {
    console.log("✅ Got data:", dataTable);
  })
  .catch(err => {
    console.error("❌ Failed to get underlying data:", err);
  });