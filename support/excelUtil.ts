import XLSX from 'xlsx';
import ExcelJS from 'exceljs';

export const readExcelCell = (
  filePath: string,
  columnName: string,
  rowNumber: number,
  sheetName?: string,
): any => {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[sheetName || workbook.SheetNames[0]];

  const data: any[] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
  });
  const headers = data[1]; // second row
  const rows = data.slice(2); // from third row onward

  const colIndex = headers.indexOf(columnName);
  return colIndex !== -1 ? rows[rowNumber - 1]?.[colIndex] : undefined;
};

export function readExcelSheet(filePath: string, sheetName?: string): any[][] {
  const workbook = XLSX.readFile(filePath);
  const targetSheet = sheetName || workbook.SheetNames[0];
  const sheet = workbook.Sheets[targetSheet];
  const data: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  return data.filter((row) => row.length > 0);
}

/**
 * Clears all cell colors (fills) from a workbook except header rows,
 * and clears all data from a specific column except headers.
 *
 * @param filePath Path to Excel file
 * @param headerRowsToKeep Number of top rows to keep unchanged (default: 2)
 * @param clearColumn Column letter (e.g. "E") to clear data from (default: M)
 */
// export async function clearAllColorsAndColumnData(
//   filePath: string,
//   headerRowsToKeep = 2,
//   clearColumn = "M"  // Comment column
// ) {
//   const workbook = new ExcelJS.Workbook();
//   await workbook.xlsx.readFile(filePath);

//   workbook.worksheets.forEach((worksheet) => {
//     worksheet.eachRow((row, rowNumber) => {
//       // Skip header rows
//       if (rowNumber <= headerRowsToKeep) return;

//       row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
//         // Safely clear fill color (TypeScript-safe)
//         (cell as any).fill = undefined;

//         // Clear data & comment in the specific column
//         if (
//           clearColumn &&
//           (typeof clearColumn === "string"
//             ? columnLetterToNumber(clearColumn) === colNumber
//             : clearColumn === colNumber)
//         ) {
//           cell.value = undefined;
//           if (typeof (cell as any).note !== "undefined") {
//             (cell as any).note = undefined; // remove Excel note/comment if present
//           }
//         }
//       });

//       row.commit();
//     });
//   });

//   await workbook.xlsx.writeFile(filePath);

//   console.log(
//     `Cleared all colors (except first ${headerRowsToKeep} header rows) and cleared all data from column ${clearColumn}, except headers.`
//   );
// }

export async function clearAllColorsAndColumnData(
  filePath: string,
  headerRowsToKeep = 2,
  clearColumns: string[] = ['M', 'N', 'O', 'P', 'Q'], // Columns to clear text and comments
  colorColumns: string[] = ['E'], // Columns to remove color (only E)
) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Map columns to numbers
  const clearColumnNumbers = clearColumns.map((col) =>
    columnLetterToNumber(col),
  );
  const colorColumnNumbers = colorColumns.map((col) =>
    columnLetterToNumber(col),
  );

  workbook.worksheets.forEach((worksheet) => {
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRowsToKeep) return; // Skip header rows

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // Clear fill color for specified color columns (E)
        if (colorColumnNumbers.includes(colNumber)) {
          (cell as any).fill = undefined;
        }

        // Clear data and comments for specified columns (M-Q)
        if (clearColumnNumbers.includes(colNumber)) {
          cell.value = undefined; // Clear data

          // Remove comments if present
          if (typeof (cell as any).note !== 'undefined') {
            (cell as any).note = undefined;
          }
        }
      });

      row.commit();
    });
  });

  await workbook.xlsx.writeFile(filePath);

  console.log(
    `Cleared all colors (except first ${headerRowsToKeep} header rows), and cleared all data from columns ${clearColumns.join(
      ', ',
    )}, except headers.`,
  );
}

/**
 * Colors a given row based on test status and appends a comment to a specific column cell.
 */
export async function updateResultinExcel(
  filePath: string,
  rowNumber: number,
  commentColumn: string = 'M', // default parameter Comment column
  commentText: string = '', // default value
  isBold: boolean = false,
  isPriceCorrect: boolean = false,
) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];

  console.log('Comment column index name -', commentColumn);
  const colCount = Math.max(
    worksheet.columnCount || 0,
    columnLetterToNumber(commentColumn),
  );
  console.log('Column number -', colCount);
  console.log('Row number -', rowNumber);

  // if (!isPriceCorrect) {
  //   console.log('----------coloring cell WHITE');
  //   worksheet.getCell('H3').style = {};
  // }

  const commentCellAddress = `${commentColumn}${rowNumber}`;
  const commentCell = worksheet.getCell(commentCellAddress);
  // console.log(commentCell);

  if (isBold) {
    let ecomCellColumn = 'E';
    let ecomCellRow = rowNumber;
    let ecomCellAddress = `${ecomCellColumn}${ecomCellRow}`;
    console.log('Ecom name bold-', ecomCellAddress);

    const ecomCell = worksheet.getCell(ecomCellAddress);
    ecomCell.style = { font: { bold: true } };
    console.log('Boldening');
    await workbook.xlsx.writeFile(filePath);
    return;
  }

  // Handle ExcelJS .note or fallback
  let existingComment = '';

  if (
    typeof (commentCell as any).note !== 'undefined' &&
    (commentCell as any).note
  ) {
    // ExcelJS note format
    const note = (commentCell as any).note;
    if (typeof note === 'string') {
      existingComment = note;
    } else if (Array.isArray(note.texts)) {
      existingComment = note.texts.map((t: any) => t.text).join('');
    }
  } else if (commentCell.value && typeof commentCell.value === 'string') {
    // Fallback: read from cell value if used for comments
    existingComment = commentCell.value;
  }

  const newComment =
    existingComment.trim().length > 0
      ? `${existingComment}\n---\n${commentText}`
      : commentText;

  // Write updated comment
  if (typeof (commentCell as any).note !== 'undefined') {
    (commentCell as any).note = {
      texts: [{ text: newComment }],
      author: 'Automated',
    };
  } else {
    commentCell.value = newComment;
  }

  // if (isPriceCorrect) {
  //   console.log('coloring cell');

  //   let priceCellColumn = 'H';
  //   let priceCellRow = rowNumber;
  //   let priceCellAddress = `${priceCellColumn}${priceCellRow}`;

  //   const priceCell = worksheet.getCell(priceCellAddress);
  //   // Add the Border property for a thick border befor filling color
  //   priceCell.style = {
  //     border: {
  //       top: { style: 'thick', color: { argb: 'FF000000' } }, // Black color
  //       left: { style: 'thick', color: { argb: 'FF000000' } },
  //       bottom: { style: 'thick', color: { argb: 'FF000000' } },
  //       right: { style: 'thick', color: { argb: 'FF000000' } },
  //     },
  //   };

  //   priceCell.fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: 'FF00FF00' },
  //   };
  // }
  let priceCellColumn = 'H';
  let priceCellRow = rowNumber;
  let priceCellAddress = `${priceCellColumn}${priceCellRow}`;
  const priceCell = worksheet.getCell(priceCellAddress);
  // ✅ Always apply border (both cases)
  priceCell.style = {
    border: {
      top: { style: 'thick', color: { argb: 'FF000000' } },
      left: { style: 'thick', color: { argb: 'FF000000' } },
      bottom: { style: 'thick', color: { argb: 'FF000000' } },
      right: { style: 'thick', color: { argb: 'FF000000' } },
    },
  };
  if (isPriceCorrect === true) {
    console.log('✅ coloring cell GREEN');
    priceCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00FF00' },
    };
  } else {
    console.log('✅ removing cell color');
    // ✅ THIS IS THE CORRECT WAY TO REMOVE COLOR
    priceCell.style = {};
  }

  console.log('Saving data to excel file...'); //udpated the M2 column with newComment->message from the test

  await workbook.xlsx.writeFile(filePath);
  console.log(`Row ${rowNumber} updated with comment in ${commentCellAddress}`);
}

/**
 * Utility: convert column letter (e.g. "E") -> index (5)
 */
function columnLetterToNumber(letter: string): number {
  let num = 0;
  for (let i = 0; i < letter.length; i++) {
    const charCode = letter.charCodeAt(i);
    if (charCode < 65 || charCode > 90) {
      // 'A' to 'Z'
      throw new Error('Invalid column letter');
    }
    num = num * 26 + (charCode - 64);
  }
  return Math.min(num, 16384); // cap at Excel max column
}
