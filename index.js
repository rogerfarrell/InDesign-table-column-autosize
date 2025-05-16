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

const getParentTable = (cell) => {
  let parent = cell;
  while (parent.constructor.name !== "Table") {
    parent = parent.parent;
  }
  return parent;
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
    console.log(`columnIndices: ${columnIndices}`);
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

        const bestWidth =
          binarySearch( minWidth, maxWidth,
            (trialWidth) =>
            {
              normalCell.width = trialWidth;
              app.activeDocument.recompose();

              if ( normalCell.overflows ) return false;
                                    return true;
            }
          );

        normalCell.width = bestWidth;
        // This is the final width of the cell.

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
        const parentColumnIndices =
          Array.from(
            { length: lastParentColumnIndex - firstParentColumnIndex + 1 },
            (_, i) => firstParentColumnIndex + i);
        const parentTable = getParentTable(spanningCell);

        parentColumnIndices.forEach(
          (parentColumnIndex) =>
          {
            const parentColumnIsEmpty =
              !columns.some( (column) => parentColumnIndex == column.index );
            console.log(parentColumnIsEmpty);
            const parentColumn =
              parentTable
                .columns
                .item(parentColumnIndex);
            console.log(`parent column: ${parentColumn}`);
            console.log(spanningCell.parent.constructor.name); // should be "Table"

            if ( parentColumnIsEmpty )
            {
              const minWidth = 3;

              const bestWidth =
                binarySearch( minWidth, maxWidth,
                  (trialWidth) =>
                  {
                    parentColumn.width = trialWidth;
                    app.activeDocument.recompose();

                    if ( spanningCell.overflows ) return false;
                                                  return true;
                  }
                );
              
              return;
            }

            const minWidth =
              columns
                .find( column => column.index == parentColumnIndex )
                .minWidth;

            const bestWidth =
              binarySearch( minWidth, maxWidth,
                (trialWidth) =>
                {
                  parentColumn.width = trialWidth;
                  app.activeDocument.recompose();

                  if ( spanningCell.overflows ) return false;
                                                return true;
                }
              );

            parentColumn.width = bestWidth;
          }
        );
      }
    );
  };
