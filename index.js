///////////////////////////////////////////////////////////////////////////
//
//   InDesign table column autosize plugin
//   (Size table columns to content in InDesign)
//   Author: Roger Farrell
//
///////////////////////////////////////////////////////////////////////////


const { app } = require("indesign");


module.exports =
  {
    commands: { autosizeColumns: () => autosizeColumns() }
  };


const autosizeColumns =
  () =>
  {
    try
    {
      if ( invalidSelection(app) ) throw new Error("Please select one or more table cells.");

      const minColumnWidth = 3;    // This is the minimum allowed by InDesign.
      const maxColumnWidth = 500;

      const selectedColumns       = getSelectedColumns(app.selection[0]);
      const selectedCells         = app.selection[0].cells.everyItem().getElements();
      const selectedNormalCells   = selectedCells.filter( cell => cell.columnSpan == 1 );
      const selectedSpanningCells = selectedCells.filter( cell => cell.columnSpan > 1 );

      selectedColumns.forEach(
        column =>
        {
          const cells =
            [
              ...selectedNormalCells
                .filter( cell => cell.parentColumn.index == column.index ),
              ...selectedSpanningCells
                .filter( cell => ((cell.parentColumn.index + cell.columnSpan) - 1) == column.index )
            ];

          column.autoGrow = false;

          const bestWidth =
            binarySearch(
              minColumnWidth,
              maxColumnWidth,
              trialWidth =>
              {
                column.width = trialWidth + "pt";
                column.recompose();

                if ( cellsOverflow(cells) ) return false;
                                            return true;
              }
            );

          column.width = bestWidth + "pt";
        }
      );
    }
    catch (e)
    {
      console.log(e);
      showAlert(`Error: ${e.message}`);
    }
  };


///////////////////////////////////////////////////////////////////////////
// HELPER FUNCTIONS
///////////////////////////////////////////////////////////////////////////

const invalidSelection =
  app =>
  {
    return app.documents.length === 0
           || !app.selection.length
           || app.selection[0].cells === undefined
           || app.selection[0].cells.length === 0;
  };

const showAlert =
  message =>
  {
    const dialog        = app.dialogs.add();
    const col           = dialog.dialogColumns.add();
    const colText       = col.staticTexts.add();
    colText.staticLabel = message;
    dialog.canCancel    = false;
    dialog.show();
    dialog.destroy();
    return;
  };

const getSelectedColumns =
  selection =>
  {
    const selectedColumns = selection.columns.everyItem().getElements();

    const table                  = getParentTable( selection.columns.firstItem() );
    const lastSelectedColumn     = selection.columns.lastItem();
    const lastColumnCells        = lastSelectedColumn.cells.everyItem().getElements();
    const maxLastColumnSpan      = lastColumnCells.reduce( (max, cell) => Math.max(max, cell.columnSpan), 1);
    const lastSpannedColumnIndex = (lastSelectedColumn.index + maxLastColumnSpan) - 1;
    const extraSpannedColumns    = table.columns.everyItem().getElements()
                                     .filter(
                                       column =>
                                         column.index > lastSelectedColumn.index
                                         && column.index <= lastSpannedColumnIndex
                                     );

    return [...selectedColumns, ...extraSpannedColumns];
  };

const getParentTable =
  object =>
  {
    if ( object.constructor.name === "Table" ) return object;
                                               return getParentTable(object.parent);
  };

const cellsOverflow =
  cells =>
    cells.some( cell => cell.overflows == true );

const binarySearch =
  (min, max, test) =>
  {
    // This is a generic binary search that takes its test
    // condition function as an argument.

    const search =
      (low, high, best) =>
      {
        const mid = Math.floor((low + high) / 2);

        if ( low > high ) return best;
        if ( test(mid) )  return search(low, mid - 1, mid);    // try smaller range
                          return search(mid + 1, high, best);  // try larger range
      };

    return search(min, max, max); // initial best is max fallback
  };
