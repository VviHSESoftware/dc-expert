/* global Excel */

/**
 * Reads the currently selected range and returns it as a formatted string (CSV-like)
 * for use as AI context.
 */
export async function getSelectionContext(): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "address"]);
      await context.sync();

      const values = range.values;
      if (!values || values.length === 0) {
        return "No data selected.";
      }

      const csvContent = range.values
        .map((row) => row.map((cell) => (cell === null || cell === undefined ? "" : cell)).join("\t"))
        .join("\n");

      return `Selected data (Range ${range.address}):\n${csvContent}`;
    });
  } catch (error) {
    console.error("Error reading selection:", error);
    return "Error reading selection context.";
  }
}

/**
 * Inserts content into the currently selected cell or directly below it.
 */
export async function insertAtSelection(content: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[content]];
      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting content:", error);
  }
}
