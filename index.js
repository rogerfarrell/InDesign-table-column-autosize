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
      if ( invalidSelection(app) )
        throw new Error("Please select one or more table cells.");

      const selectedCells = app.selection[0].cells;
      const cells = selectedCells.everyItem().getElements();

      autosizeCells(cells);
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
  cell =>
  {
      return cell.contents == "" && !cell.overflows;
      // cell.contents returns "" if all content is overflowed.
  };

const getParentTable =
  object =>
  {
    if ( object.constructor.name === "Table" ) return object;
                                               return getParentTable(object.parent);
  };

const resize = 
  (elementToResize, elementToTest, min, max) =>
  {
     const bestWidth =
       binarySearch( min, max,
         trialWidth =>
         {
           elementToResize.width = trialWidth;
           app.activeDocument.recompose();

           if ( elementToTest.overflows ) return false;
                                          return true;
         }
       );

       elementToResize.width = bestWidth;
  };

const autosizeCells =
  cells =>
  {
    const maxWidth = 500;
    const minWidth = 3;
    // 3 is the minimum allowed for col width.
    // Setting the minimum lower will result in an error.
    // (InDesign interprets these units as points.)

    const nonEmptyCells = cells.filter( cell => !isTrulyEmpty(cell) );
    nonEmptyCells.forEach( cell => cell.parentColumn.autoGrow = false );
    // autoGrow must be false for cell contents to overflow predictably.

    const columnIndices =
      [...new Set( nonEmptyCells.map( cell => cell.parentColumn.index ) )];
    const columns =
      columnIndices.map( index => ({ index: index, minWidth: minWidth }) );

    const normalCells   = nonEmptyCells.filter( cell => cell.columnSpan == 1 );
    const spanningCells = nonEmptyCells.filter( cell => cell.columnSpan >  1 );
    // Spanning (merged) cells are handled separately at the end.
    // This side-steps complicated column width calculations. 

    normalCells.forEach(
      normalCell =>
      {
        const currentColumn =
          columns.find( column => column.index == normalCell.parentColumn.index );
        const minWidth = currentColumn.minWidth;

        resize(normalCell, normalCell, minWidth, maxWidth);

        if ( normalCell.width >= currentColumn.minWidth )
        {
          currentColumn.minWidth = normalCell.width;
        }
      }
    );

    spanningCells.forEach(
      spanningCell =>
      {
        const firstParentColumnIndex = spanningCell.parentColumn.index;
        const lastParentColumnIndex = (firstParentColumnIndex + spanningCell.columnSpan) - 1;
        const parentColumnIndices =
          Array.from(
            { length: lastParentColumnIndex - firstParentColumnIndex + 1 },
            (_, i) => firstParentColumnIndex + i
          );
        const parentTable = getParentTable(spanningCell);
        const parentColumnAt = index => parentTable.columns.item(index);

        parentColumnIndices .forEach(
          parentColumnIndex =>
          {
            parentColumnAt(parentColumnIndex).width = maxWidth;
          }
            // Sets all the parent columns to maxWidth to start the calc.
        );

        parentColumnIndices.forEach(
          parentColumnIndex =>
          {
            const parentColumnIsEmpty =
              !columns.some( (column) => parentColumnIndex == column.index );

            const minWidth =
              parentColumnIsEmpty
                ? 3
                : columns
                    .find( column => column.index == parentColumnIndex )
                    .minWidth;
            //TODO I would like to replace this with something more elegant.

            resize(parentColumnAt(parentColumnIndex), spanningCell, minWidth, maxWidth);
          }
        );
      }
    );
  };
