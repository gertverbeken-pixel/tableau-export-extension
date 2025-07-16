
document.getElementById('exportBtn').addEventListener('click', async () => {
  try {
    console.log("🔁 Button clicked, initializing Tableau...");
    await tableau.extensions.initializeAsync();

    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const worksheetNames = dashboard.worksheets.map(ws => ws.name);
    console.log("✅ Worksheets:", worksheetNames);

    alert("Found worksheets:\n" + worksheetNames.join("\n"));
  } catch (err) {
    console.log("❌ Initialization or fetch failed.");
    console.log("🔎 Error:", JSON.stringify(err, Object.getOwnPropertyNames(err)));
    alert("Something went wrong. Check the console.");
  }
});
