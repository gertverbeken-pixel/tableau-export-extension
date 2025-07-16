// extension.js

(async function () {
  console.log("🔁 Button clicked, initializing Tableau...");

  try {
    await tableau.extensions.initializeAsync();
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    const worksheet = dashboard.worksheets[0];

    const summaryData = await worksheet.getSummaryDataAsync();
    const data = summaryData.data;
    const columns = summaryData.columns.map(col => col.fieldName);

    // Prepare CSV
    let csvContent = columns.join(",") + "\n";
    data.forEach(row => {
      const rowData = row.map(cell => cell.formattedValue);
      csvContent += rowData.join(",") + "\n";
    });

    // Optional: Show the user
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.setAttribute("href", URL.createObjectURL(blob));
    link.setAttribute("download", "summary_export.csv");
    link.click();

    // 🚀 Send metadata to webhook
    await fetch("https://script.google.com/macros/s/AKfycbxrgR7BqnSoL01D-LzC_kBkqto579BRl9YBxJcYABhItWxvR-pbOCai6BJz2uCeewampw/exec", {
      method: "POST",
      mode: "no-cors",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        user: tableau.extensions.environment.username,
        dashboard: dashboard.name,
        worksheet: worksheet.name,
        rowCount: data.length,
        timestamp: new Date().toISOString()
      })
    });

    console.log("✅ Export + ingest successful");

  } catch (err) {
    console.error("❌ Error in export process:", err);
  }
})();