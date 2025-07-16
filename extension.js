document.getElementById("exportBtn").addEventListener("click", async function () {
  console.log("✅ Export button clicked");

  try {
    await tableau.extensions.initializeAsync();
    console.log("✅ Extension initialized");

    const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    if (!worksheet) {
      console.error("❌ No worksheet found");
      alert("No worksheet available in the dashboard.");
      return;
    }

    console.log("✅ Worksheet found:", worksheet.name);

    const dataTable = await worksheet.getUnderlyingTableDataAsync();
    const filters = await worksheet.getFiltersAsync();

    const filterSummary = filters.map(f =>
      `${f.fieldName}=${f.appliedValues.map(v => v.value).join(",")}`
    ).join("; ");

    const rowCount = dataTable.totalRowCount;

    console.log("📦 Sending data to webhook:", {
      user: tableau.extensions.environment.username,
      view: worksheet.name,
      filters: filterSummary,
      rowCount: rowCount
    });

    await fetch("https://script.google.com/macros/s/AKfycbxrgR7BqnSoL01D-LzC_kBkqto579BRl9YBxJcYABhItWxvR-pbOCai6BJz2uCeewampw/exec", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user: tableau.extensions.environment.username,
        view: worksheet.name,
        filters: filterSummary,
        rowCount: rowCount
      })
    });

    alert("✅ Export tracked and sent to webhook!");
  } catch (err) {
    console.error("❌ Error in export process:", err);
    alert("Something went wrong — check the console.");
  }
});