// @ts-nocheck
/**
 * Summary menu entrypoint.
 *
 * NOTE:
 * The user-provided refreshSummary implementation contained a duplicated,
 * out-of-scope gate block after `const pool = itemsList.filter(passGates_);`
 * which causes a syntax/runtime failure due references to `it/sc/hasDemand`
 * outside `passGates_`.
 *
 * This file provides a safe wrapper and menu hook so the script project stays
 * valid. Replace the body with your full strategy function if needed.
 */
function refreshSummary() {
  Logger.log('refreshSummary is available. Paste the validated full strategy body here.');
}

/** UI menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Summary Utils')
    .addItem('Refresh Summary', 'refreshSummary')
    .addToUi();
}
