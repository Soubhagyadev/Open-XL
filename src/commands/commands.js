/**
 * Ribbon command handlers.
 * Each function must call event.completed() when done.
 */

Office.onReady(() => {});

/**
 * @param {Office.AddinCommands.Event} event
 */
export function openTaskpane(event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

// Register commands
if (typeof Office !== "undefined") {
  Office.actions.associate("openTaskpane", openTaskpane);
}
