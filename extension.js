document.getElementById('exportBtn').addEventListener('click', async () => {
  try {
    console.log("🧪 Initializing Tableau Extension");
    await tableau.extensions.initializeAsync();

    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const worksheetNames = dashboard.worksheets.map(ws => ws.name);
    console.log("📄 Worksheets found:", worksheetNames);

    alert("Found worksheets:\n" + worksheetNames.join("\n"));

  } catch (err) {
    console.error("❌ Initialization or fetch failed:", err);
    alert("Failed to fetch worksheets. See console.");
  }
});