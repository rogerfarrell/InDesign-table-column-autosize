const { app } = require("indesign");

module.exports = {
  commands: { autofitColumns: () => autofitColumns() } };


function autofitColumns()
{
  try
  {
    if (app.documents.length === 0)
    {
      showAlert("No document open.");
    }
    else if (!isCellInTableSelected())
    {
      showAlert("Please select a cell inside a table.");
    }
    else
    {
      const table = app.selection[0].parent;
      for (let colIndex = 0; colIndex < table.columns.length; colIndex++)
      {
        shrinkColumnToFit(table, colIndex);
      }
    }
  }
  catch (e)
  {
    console.log(e);
    showAlert("An error occurred: " + e.message);
  }
}


function showAlert(message)
{
  const dialog = app.dialogs.add();
  const col = dialog.dialogColumns.add();
  const colText = col.staticTexts.add();
  colText.staticLabel = message;;

  dialog.canCancel = false;
  dialog.show();
  dialog.destroy();
  return;
}

function isCellInTableSelected()
{
  const sel = app.selection;
  return ( sel.length > 0 &&
           sel[0].parent &&
           sel[0].parent.constructor.name === "Table" );
}

function binarySearch(min, max, test)
{
  let best = max;

  while (min <= max)
  {
    const mid = Math.floor((min + max) / 2);
    const passed = test(mid);

    if (passed)
    {
      best = mid;
      max = mid - 1;
    }
    else
    {
      min = mid + 1;
    }
  }

  return best;
}

function shrinkColumnToFit(table, colIndex, initialWidth = 500)
{
  const column = table.columns.item(colIndex);
  column.autoGrow = false;
  column.width = initialWidth + "pt";

  const bestWidth = binarySearch(1, initialWidth, (trialWidth) =>
    {
      column.width = trialWidth + "pt";
      app.activeDocument.recompose();

      for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++)
      {
        const cell = table.rows.item(rowIndex).cells.item(colIndex);
        if (cell.overflows)
        {
          return false;
        }
      }
      return true;
    });

  column.width = bestWidth + "pt";
}
