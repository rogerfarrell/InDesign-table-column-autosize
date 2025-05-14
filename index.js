const { app } = require("indesign");

module.exports = {
  commands: { autofitColumns: () => autofitColumns() }
};

function autofitColumns() {
  try {
    if (app.documents.length === 0) {
      showAlert("No document open.");
      return;
    }

    if (
      !app.selection.length ||
      typeof !app.selection[0].cells === "undefined" ||
      app.selection[0].cells.length === 0
    ) {
      showAlert("Please select one or more table cells.");
      return;
    }

    const selectedCells = app.selection[0].cells;
    const cells = selectedCells.everyItem().getElements();

    for (const cell of cells) {
      autosizeCell(cell);
    }
  } catch (e) {
    console.log(e);
    showAlert("An error occurred: " + e.message);
  }
}

function showAlert(message) {
  const dialog = app.dialogs.add();
  const col = dialog.dialogColumns.add();
  const colText = col.staticTexts.add();
  colText.staticLabel = message;
  dialog.canCancel = false;
  dialog.show();
  dialog.destroy();
  return;
}

function binarySearch(min, max, test) {
  let best = max;

  while (min <= max) {
    const mid = Math.floor((min + max) / 2);
    const passed = test(mid);

    if (passed) {
      best = mid;
      max = mid - 1;
    } else {
      min = mid + 1;
    }
  }

  return best;
}

function autosizeCell(cell, initialWidth = 500) {
  const column = cell.parentColumn;
  column.autoGrow = false;

  cell.width = initialWidth + "pt";

  // 3 is the minimum allowed for col width. Setting the minimum lower will result in an error.
  const bestWidth = binarySearch(3, initialWidth, (trialWidth) => {
    cell.width = trialWidth + "pt";
    app.activeDocument.recompose();

    if (cell.overflows) {
      return false;
    }

    return true;
  });

  cell.width = bestWidth + "pt";
}
