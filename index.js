const { app } = require("indesign");

module.exports =
  {
    commands: { autofitColumns: () => autofitColumns() }
  };


///////////////////////////////////////////////////////////////////////////
// COMMANDS
///////////////////////////////////////////////////////////////////////////

function autofitColumns()
{
  try
  {
    console.log("command called");
    if (app.documents.length === 0)
    {
      showAlert("No document open.");
      return;
    }

    if
    (
      !app.selection.length ||
      typeof !app.selection[0].cells === "undefined" ||
      app.selection[0].cells.length === 0
    )
    {
      showAlert("Please select one or more table cells.");
      return;
    }

    const selectedCells = app.selection[0].cells;
    const cells = selectedCells.everyItem().getElements();

    autosizeCell(cells);
  }
  catch (e)
  {
    console.log(e);
    showAlert("An error occurred: " + e.message);
  }
}


///////////////////////////////////////////////////////////////////////////
// HELPER FUNCTIONS
///////////////////////////////////////////////////////////////////////////

function showAlert(message)
{
  const dialog = app.dialogs.add();
  const col = dialog.dialogColumns.add();
  const colText = col.staticTexts.add();
  colText.staticLabel = message;
  dialog.canCancel = false;
  dialog.show();
  dialog.destroy();
  return;
}

function binarySearch(min, max, test)
{
  console.log(`binarySearch called, min: ${min} max ${max}`);

  let best = max;

  while (min <= max)
  {
    const mid = Math.floor((min + max) / 2);
    const passed = test(mid);

    if (!passed)
      { min = mid + 1; }
    else
      { best = mid; max = mid - 1; }
  }

  console.log("best: " + best);
  return best;
}

function isTrulyEmpty(cell)
{
    // cell.contents returns "" if all content is overflowed
    return cell.contents == "" && !cell.overflows;
}

function autosizeCell(cells)
{
  console.log("autosizeCell called");

  const startingMaxWidth = 500;

  let columns = [];
  let spanningCells = [];

  for (const cell of cells)
  {
    if ( cell.columnSpan > 1 )
    {
      spanningCells.push(cell);
      continue;
    }

    console.log(`cell contents: ${cell.contents}`);

    if ( isTrulyEmpty(cell) ) continue;

    const parentColumn = cell.parentColumn;
    parentColumn.autoGrow = false;

    const matchedColumn = () =>
      columns.find(column => parentColumn.index == column.index);

    // 3 is the minimum allowed for col width.
    // Setting the minimum lower will result in an error.
    if ( !matchedColumn() ) columns.push( { index: parentColumn.index, minWidth: 3 } );

    const startingMinWidth = matchedColumn().minWidth;
    console.log(`matchedColumn().width: ${matchedColumn().minWidth}`);

    cell.width = startingMaxWidth + "pt";

    const bestWidth =
      binarySearch(
        startingMinWidth,
        startingMaxWidth,
        (trialWidth) =>
          {
            cell.width = trialWidth + "pt";
            app.activeDocument.recompose();

            if (cell.overflows) return false;

            return true;
          }
      );

    cell.width = bestWidth + "pt";
    matchedColumn().minWidth = bestWidth;
  }

  for ( const spanningCell of spanningCells )
  {
    if ( isTrulyEmpty(spanningCell) ) continue;

    const parentColumn = spanningCell.parentColumn;
    parentColumn.autoGrow = false;

    const startingMinWidth = spanningCell.width;

    const bestWidth =
      binarySearch(
        startingMinWidth,
        startingMaxWidth,
        (trialWidth) =>
          {
            spanningCell.width = trialWidth + "pt";
            app.activeDocument.recompose();

            if (spanningCell.overflows) return false;

            return true;
          }
      );

    spanningCell.width = bestWidth + "pt";
  }
}
