/**
 * Thin wrapper around Excel.run for consistent error handling.
 * @param {function(Excel.RequestContext): Promise<void>} callback
 */
export async function runExcel(callback) {
  try {
    await Excel.run(callback);
  } catch (error) {
    console.error("[Excel Error]", error);
    throw error;
  }
}

/**
 * Initialize Office.js global settings.
 */
export function initializeOffice() {
  Office.context.document.settings.set("initialized", true);
  Office.context.document.settings.saveAsync();
}
