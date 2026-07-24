/* global Office */

Office.onReady(() => {
  // Commands are registered here for ribbon button actions
});

/**
 * Shows the AI Copilot taskpane.
 * This is called when the user clicks the ribbon button.
 */
function showTaskpane(event: Office.AddinCommands.Event) {
  Office.addin.showAsTaskpane();
  event.completed();
}

// Register the command functions globally
(globalThis as any).showTaskpane = showTaskpane;
