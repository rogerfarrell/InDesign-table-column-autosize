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
    commands: { autofitColumns: () => autofitColumns() }
  };


const autofitColumns =
  () =>
  {
    try
    {
      if (app.documents.length === 0)
      {
        showAlert("No document open.");
        return;
      }

      const noCellSelected =
        !app.selection.length
        || app.selection[0].cells === undefined
        || app.selection[0].cells.length === 0;

      if ( noCellSelected )
      {
        showAlert("Please select one or more table cells.");
        return;
      }

      const selectedCells = app.selection[0].cells;
      const cells = selectedCells.everyItem().getElements();

      autosizeCells(cells);
    }
    catch (e)
    {
      console.log(e);
      showAlert("An error occurred: " + e.message);
    }
  };


///////////////////////////////////////////////////////////////////////////
// HELPER FUNCTIONS
///////////////////////////////////////////////////////////////////////////

const showAlert =
  (message) =>
  {
    const dialog = app.dialogs.add();
    const col = dialog.dialogColumns.add();
    const colText = col.staticTexts.add();
    colText.staticLabel = message;
    dialog.canCancel = false;
    dialog.show();
    dialog.destroy();
    return;
  };

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

const isTrulyEmpty =
  (cell) =>
  {
      return cell.contents == "" && !cell.overflows;
      // cell.contents returns "" if all content is overflowed.
  };

const testCell =
  (cell, min, max) =>
  {
    const bestWidth =
      binarySearch( min, max,
        (trialWidth) =>
        {
          cell.width = trialWidth;
          app.activeDocument.recompose();

          if ( cell.overflows ) return false;
                                return true;
        }
      );

    cell.width = bestWidth;
    // This is the final width of the cell.
  };

const autosizeCells =
  (cells) =>
  {
    const maxWidth = 500;
    const minWidth = 3;
    // 3 is the minimum allowed for col width.
    // Setting the minimum lower will result in an error.
    // (InDesign interprets these units as points.)

    const nonEmptyCells = cells.filter( cell => !isTrulyEmpty(cell) );
    nonEmptyCells.forEach( cell => cell.parentColumn.autoGrow = false );
    // autoGrow must be false for cell contents to overflow predictably.

    const columnIndices = [...new Set( nonEmptyCells.map( cell => cell.parentColumn.index ) )];
    const columns = columnIndices.map( index => ({ index: index, minWidth: minWidth }) );

    const normalCells   = nonEmptyCells.filter( cell => cell.columnSpan == 1 );
    const spanningCells = nonEmptyCells.filter( cell => cell.columnSpan >  1 );
    // Spanning (merged) cells are handled separately at the end.
    // This side-steps complicated column width calculations. 

    normalCells.forEach(
      (normalCell) =>
      {
        const currentColumn =
          columns.find( column => column.index == normalCell.parentColumn.index );
        const minWidth = currentColumn.minWidth;

        testCell(normalCell, minWidth, maxWidth);

        if ( normalCell.width >= currentColumn.minWidth )
        {
          currentColumn.minWidth = normalCell.width;
        }
      }
    );

    spanningCells.forEach(
      (spanningCell) =>
      {
        const firstParentColumnIndex = spanningCell.parentColumn.index;
        const lastParentColumnIndex = (firstParentColumnIndex + spanningCell.columnSpan) - 1;
        const lastParentColumnExists =
          columns.some( (index) => index == lastParentColumnIndex );
        const inParentRange =
          (index) =>
            (index >= firstParentColumnIndex)
            && (index <= lastParentColumnIndex);
        // I pulled this out of the filter below to make it read easier.

        if ( !lastParentColumnExists )
        {
          const minWidth =
            columns
              .filter( (column) => inParentRange(column.index) )
              .reduce( (sum, column) => sum + column.minWidth, 0 );
          // minWidth is the sum of all columns spanned.

          return testCell(spanningCell, minWidth, maxWidth);
        }

        const minWidth = spanningCell.width;
        return testCell(spanningCell, minWidth, maxWidth);
      }
    );
  };
