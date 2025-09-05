function main(workbook: ExcelScript.Workbook) {
  // FACT_FINANCE
  const fact = workbook.getWorksheet("Fact_Finance");
  const factTable = fact.getTables()[0];

  ensureColWithFormula(factTable, "SegmentID", "=UPPER(LEFT([@Segment],3))");
  ensureColWithFormula(factTable, "CountryID", "=UPPER(LEFT([@Country],2))");
  ensureColWithFormula(factTable, "ProductID", "=UPPER(LEFT([@Product],2))");

  // PRODUCT (nếu có)
  const prodSheet = workbook.getWorksheet("Product");
  if (prodSheet && prodSheet.getTables().length > 0) {
    const prodTable = prodSheet.getTables()[0];
    ensureColWithFormula(prodTable, "ProductID", "=UPPER(LEFT([@Product],2))");
  }
}

function ensureColWithFormula(table: ExcelScript.Table, name: string, formula: string) {
  let col = table.getColumnByName(name);
  if (!col) {
    table.addColumn(-1, { name });
    col = table.getColumnByName(name);
  }
  // setFormula cho structured reference (English)
  col.getRangeBetweenHeaderAndTotal().setFormula(formula);
}
