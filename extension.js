document.getElementById('exportBtn').addEventListener('click', async () => {
  try {
    console.log("🔁 Export button clicked");

    // Tableau SDK check
    if (typeof tableau === 'undefined') {
      throw new Error('Tableau Extensions API not loaded.');
    }

    await tableau.extensions.initializeAsync();

    const worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
    const worksheet = worksheets[0]; // Pick first worksheet (customize as needed)

    const dataTable = await worksheet.getUnderlyingTableDataAsync();
    console.log("✅ Data received:", dataTable);

    alert(`Export succeeded. ${dataTable.data.length} rows fetched.`);

  } catch (err) {
    console.error("❌ Error in export process:", err);
    alert("Export failed. Check console.");
  }
});