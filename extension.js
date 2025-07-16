document.getElementById("exportBtn").addEventListener("click", async function () {
  await tableau.extensions.initializeAsync();
  const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
  const user = tableau.extensions.environment.username;

  try {
    const dataTable = await worksheet.getUnderlyingTableDataAsync();
    const filters = await worksheet.getFiltersAsync();

    const filterSummary = filters.map(f =>
      \`\${f.fieldName}=\${f.appliedValues.map(v => v.value).join(",")}\`
    ).join("; ");

    const rowCount = dataTable.totalRowCount;

    // 🔗 Send to webhook
    await fetch("https://script.google.com/macros/s/AKfycbxrgR7BqnSoL01D-LzC_kBkqto579BRl9YBxJcYABhItWxvR-pbOCai6BJz2uCeewampw/exec", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        user: user,
        view: worksheet.name,
        filters: filterSummary,
        rowCount: rowCount
      })
    });

    alert("Export tracked! (This version doesn't generate a file — tracking only.)");
  } catch (err) {
    console.error("Export error:", err);
    alert("Something went wrong. Check console.");
  }
});
