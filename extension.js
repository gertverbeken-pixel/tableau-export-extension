document.getElementById("exportBtn").addEventListener("click", async () => {
  try {
    console.log("🔁 Button clicked, initializing Tableau...");
    await tableau.extensions.initializeAsync();

    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const worksheet = dashboard.worksheets[0]; // change if you need a specific worksheet

    const dataTable = await worksheet.getUnderlyingDataAsync();

    const exportInfo = {
      user: tableau.extensions.environment.userName,
      dashboard: dashboard.name,
      worksheet: worksheet.name,
      timestamp: new Date().toISOString(),
      rowCount: dataTable.data.length
    };

    await fetch("https://script.google.com/macros/s/AKfycbxrgR7BqnSoL01D-LzC_kBkqto579BRl9YBxJcYABhItWxvR-pbOCai6BJz2uCeewampw/exec", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(exportInfo)
    });

    alert("✅ Export logged to Google Sheets!");
  } catch (err) {
    console.error("❌ Error in export process:", err);
  }
});