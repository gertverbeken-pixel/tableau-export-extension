tableau.extensions.initializeAsync().then(() => {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const worksheet = dashboard.worksheets[0];

  worksheet.getUnderlyingTableDataAsync()
    .then(dataTable => {
      console.log("✅ Got data:", dataTable);
    })
    .catch(err => {
      console.error("❌ Failed to get underlying data:", err);
    });

}).catch(err => {
  console.error("❌ Failed to initialize Tableau:", err);
});