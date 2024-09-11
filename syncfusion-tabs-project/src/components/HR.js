import React, { useRef, useEffect } from 'react';
import { SpreadsheetComponent, SheetsDirective, SheetDirective, RowsDirective, RowDirective, ColumnsDirective, ColumnDirective } from '@syncfusion/ej2-react-spreadsheet';
import { registerLicense } from '@syncfusion/ej2-base';
import '@syncfusion/ej2-base/styles/material.css';
import '@syncfusion/ej2-inputs/styles/material.css';
import '@syncfusion/ej2-buttons/styles/material.css';
import '@syncfusion/ej2-splitbuttons/styles/material.css';
import '@syncfusion/ej2-lists/styles/material.css';
import '@syncfusion/ej2-navigations/styles/material.css';
import '@syncfusion/ej2-popups/styles/material.css';
import '@syncfusion/ej2-dropdowns/styles/material.css';
import '@syncfusion/ej2-grids/styles/material.css';
import '@syncfusion/ej2-react-spreadsheet/styles/material.css';

// Register the Syncfusion license key
registerLicense('Ngo9BigBOggjHTQxAR8/V1NCaF1cVGhIfEx1RHxQdld5ZFRHallYTnNWUj0eQnxTdEFjUX5acXRRQmJYVEJ1Ww==');

const HR = ({ pressure, temperature, distance, time, velocity, weight }) => {
  const spreadsheetRef = useRef(null);
  const editableRanges = [
   'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B20', 'B21', 'B22', 'B23', 'B24', 'B25', 'B26','D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11', 'D12', 'D13', 'D14', 'D15', 'D16', 'D17', 'D18', 'D19', 'D20', 'D21', 'D22', 'D23', 'D24', 'D25', 'D26', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12', 'F13', 'F14', 'F15', 'F16', 'F17', 'F18', 'F19', 'F20', 'F21', 'F22', 'F23', 'F24', 'F25', 'F26', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11', 'H12', 'H13', 'H14', 'H15', 'H16', 'H17', 'H18', 'H19', 'H20', 'H21', 'H22', 'H23', 'H24', 'H25', 'H26', 'J3', 'J4', 'J5', 'J6', 'J7', 'J8', 'J9', 'J10', 'J11', 'J12', 'J13', 'J14', 'J15', 'J16', 'J17', 'J18', 'J19', 'J20', 'J21', 'J22', 'J23', 'J24', 'J25', 'J26', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L10', 'L11', 'L12', 'L13', 'L14', 'L15', 'L16', 'L17', 'L18', 'L19', 'L20', 'L21', 'L22', 'L23', 'L24', 'L25', 'L26'
   ];


  const onBeforeSelect = (args) => {
    // Check if indexes are defined and have the required elements
    if (args.indexes && args.indexes.length >= 2) {
      const { indexes } = args;
      const rowIndex = indexes[0];
      const colIndex = indexes[1];
      const cell = String.fromCharCode(65 + colIndex) + (rowIndex + 1);

      if (!editableRanges.includes(cell)) {
        args.cancel = true; // Cancel selection if cell is not in editableRanges
      }
    }
  };

  const onCellSelected = (args) => {
    const spreadsheet = spreadsheetRef.current;
    const { rowIndex, colIndex } = args;
    const cell = String.fromCharCode(65 + colIndex) + (rowIndex + 1);

    if (!editableRanges.includes(cell)) {
      args.cancel = true; // Prevent selection
      let nextEditableCell = findNextEditableCell(rowIndex, colIndex);
      if (nextEditableCell) {
        setTimeout(() => spreadsheet.selectRange(nextEditableCell), 0);
      }
    }
  };

  const onKeyDown = (e) => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet && (e.key === 'Tab' || e.key.startsWith('Arrow'))) {
      e.preventDefault();
      const activeCell = spreadsheet.getActiveSheet().activeCell;
      let [col, row] = activeCell.match(/[A-Z]+|[0-9]+/g);
      row = parseInt(row, 10) - 1;
      col = col.charCodeAt(0) - 65;

      let nextEditableCell;
      if (e.key === 'Tab' || e.key === 'ArrowRight') {
        nextEditableCell = findNextEditableCell(row, col + 0);
      } else if (e.key === 'ArrowLeft') {
        nextEditableCell = findPreviousEditableCell(row, col - 0);
      } else if (e.key === 'ArrowDown') {
        nextEditableCell = findNextEditableCell(row + 0, col);
      } else if (e.key === 'ArrowUp') {
        nextEditableCell = findPreviousEditableCell(row - 0, col);
      }

      if (nextEditableCell) {
        setTimeout(() => spreadsheet.selectRange(nextEditableCell), 0);
      }
    }
  };

  const findNextEditableCell = (rowIndex, colIndex) => {
    const flatEditableRanges = editableRanges.map(range => {
      const col = range.charCodeAt(0) - 65;
      const row = parseInt(range.substring(1)) - 1;
      return { row, col, cell: range };
    }).sort((a, b) => (a.row - b.row) || (a.col - b.col));

    for (let i = 0; i < flatEditableRanges.length; i++) {
      const { row, col, cell } = flatEditableRanges[i];
      if (row > rowIndex || (row === rowIndex && col > colIndex)) {
        return cell;
      }
    }

    return flatEditableRanges[0].cell; // Loop back to the first editable cell
  };

  const findPreviousEditableCell = (rowIndex, colIndex) => {
    const flatEditableRanges = editableRanges.map(range => {
      const col = range.charCodeAt(0) - 65;
      const row = parseInt(range.substring(1)) - 1;
      return { row, col, cell: range };
    }).sort((a, b) => (a.row - b.row) || (a.col - b.col));

    for (let i = flatEditableRanges.length - 1; i >= 0; i--) {
      const { row, col, cell } = flatEditableRanges[i];
      if (row < rowIndex || (row === rowIndex && col < colIndex)) {
        return cell;
      }
    }

    return flatEditableRanges[flatEditableRanges.length - 1].cell; // Loop back to the last editable cell
  };

  // const setCellBackground = () => {
  //   // Select the specific cell (A1 in this case)
  //   const cell = document.querySelector('td[aria-colindex="1"][aria-rowindex="1"]'); // Adjust selectors as needed
  
  //   if (cell) {
  //     cell.style.position = 'relative';
  //     cell.style.backgroundImage = 'url("https://i.postimg.cc/gj2yXqDH/picture-picture.png")';
  //     cell.style.backgroundSize = 'cover';
  //     cell.style.backgroundPosition = 'center';
  //     cell.style.backgroundRepeat = 'no-repeat';
  //   }
  // };
  
  // setCellBackground();
  
  const insertImage = () => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet) {
      
      const image1 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 1200,  // Adjust width to fit within B10:B14
        height: 65, // Adjust height to fit within B10:B14
        opacity: 0
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image1], 'A1'); // Start from the upper-left corner of the range

      const image2 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 100,  // Adjust width to fit within B10:B14
        height: 480, // Adjust height to fit within B10:B14
        opacity: 0
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image2], 'A3'); // Start from the upper-left corner of the range
    
      const image3 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 100,  // Adjust width to fit within B10:B14
        height: 480, // Adjust height to fit within B10:B14
        opacity: 0
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image3], 'D3'); // Start from the upper-left corner of the range
    
      const image4 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 104,  // Adjust width to fit within B10:B14
        height: 480, // Adjust height to fit within B10:B14
        opacity: 0
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image4], 'G3'); // Start from the upper-left corner of the range
    
      const image5 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 112,  // Adjust width to fit within B10:B14
        height: 480 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image5], 'J3'); // Start from the upper-left corner of the range
    
      const image6 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 120,  // Adjust width to fit within B10:B14
        height: 480, // Adjust height to fit within B10:B14
        opacity: 0
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image6], 'M3'); // Start from the upper-left corner of the range
    
      const image7 = {
        backgroundImage: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 120,  // Adjust width to fit within B10:B14
        height: 480, // Adjust height to fit within B10:B14
        opacity: 0
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image7], 'P3'); // Start from the upper-left corner of the range
    
    }
  };
  
  const customizeSpreadsheet = () => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet) {
      try {
        for (let i = 1; i <= 24; i++) {
          spreadsheet.updateCell({
            value: i,
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: 'rgb(248, 248, 200)'
            }
          }, `A${i + 2}`); // Start from A3
        }
        for (let i = 25; i <= 48; i++) {
          spreadsheet.updateCell({
            value: i,
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: 'rgb(248, 248, 200)'
            }
          }, `C${i - 22}`); // Start from C3 (i.e., 25 - 22 = 3)
        }
        for (let i = 49; i <= 72; i++) {
          spreadsheet.updateCell({
            value: i,
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: 'rgb(248, 248, 200)'
            }
          }, `E${i - 46}`); // Start from E3 (i.e., 49 - 46 = 3)
        }
        for (let i = 73; i <= 96; i++) {
          spreadsheet.updateCell({
            value: i,
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: 'rgb(248, 248, 200)'
            }
          }, `G${i - 70}`); // Start from G3 (i.e., 73 - 70 = 3)
        }
  
        // Add numbers 97 to 120 in column I starting from I3
        for (let i = 97; i <= 120; i++) {
          spreadsheet.updateCell({
            value: i,
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: 'rgb(248, 248, 200)'
            }
          }, `I${i - 94}`); // Start from I3 (i.e., 97 - 94 = 3)
        }
        for (let i = 121; i <= 144; i++) {
          spreadsheet.updateCell({
            value: i,
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: 'rgb(248, 248, 200)'
            }
          }, `K${i - 118}`); // Start from L3 (i.e., 121 - 118 = 3)
        }

        for (let i = 1; i <= 24; i++) {
          spreadsheet.updateCell({
            value: '',
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: '#f7cfd5'
            }
          }, `B${i + 2}`); // Start from A3
        }
        for (let i = 25; i <= 48; i++) {
          spreadsheet.updateCell({
            value: '',
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: '#f7cfd5'
            }
          }, `D${i - 22}`); // Start from C3 (i.e., 25 - 22 = 3)
        }
        for (let i = 49; i <= 72; i++) {
          spreadsheet.updateCell({
            value: '',
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: '#f7cfd5'
            }
          }, `F${i - 46}`); // Start from E3 (i.e., 49 - 46 = 3)
        }
        for (let i = 73; i <= 96; i++) {
          spreadsheet.updateCell({
            value: '',
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: '#f7cfd5'
            }
          }, `H${i - 70}`); // Start from G3 (i.e., 73 - 70 = 3)
        }
  
        // Add numbers 97 to 120 in column I starting from I3
        for (let i = 97; i <= 120; i++) {
          spreadsheet.updateCell({
            value: '',
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: '#f7cfd5'
            }
          }, `J${i - 94}`); // Start from I3 (i.e., 97 - 94 = 3)
        }
        for (let i = 121; i <= 144; i++) {
          spreadsheet.updateCell({
            value: '',
            style: {
              textAlign: 'center',
              verticalAlign: 'middle',
              fontSize: '10pt',
              border: '1px solid black',
              backgroundColor: '#f7cfd5'
            }
          }, `L${i - 118}`); // Start from L3 (i.e., 121 - 118 = 3)
        }
        // Update and style the header cell
        spreadsheet.updateCell({
          value: 'Hot Runner Controller Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(163, 152, 236)',
            color: 'white',
            border: '1.5px solid black'
          }
        }, 'A1');
        spreadsheet.updateCell({
          value: 'Zone no.',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A2');
        spreadsheet.updateCell({
          value: 'Zone no.',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C2');
        spreadsheet.updateCell({
          value: 'Zone no.',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E2');
        spreadsheet.updateCell({
          value: 'Zone no.',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'G2');
        spreadsheet.updateCell({
          value: 'Zone no.',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'I2');
        spreadsheet.updateCell({
          value: 'Zone no.',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'K2');
        spreadsheet.updateCell({
          value: 'Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B2');
        spreadsheet.updateCell({
          value: 'Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D2');
        spreadsheet.updateCell({
          value: 'Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F2');
        spreadsheet.updateCell({
          value: 'Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H2');
        spreadsheet.updateCell({
          value: 'Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J2');
        spreadsheet.updateCell({
          value: 'Settings',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(181, 185, 240)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L2');
        spreadsheet.updateCell({
          value: '',
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent' 
          }
        }, 'A28');
        spreadsheet.updateCell({
          value: '',
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',    
            borderRight: 'none', 
            borderBottom: 'none',          
            color: 'transparent' 
          }
        }, 'A27');
        spreadsheet.updateCell({
          value: '',
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',              
            color: 'transparent' 
          }
        }, 'N1');
        spreadsheet.updateCell({
          value: '',
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',              
            color: 'transparent' 
          }
        }, 'O1');
        spreadsheet.updateCell({
          value: '',
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',              
            borderRight: 'none', 
            borderBottom: 'none', 
            color: 'transparent' 
          }
        }, 'M1');
        
      
        
        // Merge specified cell ranges
        spreadsheet.merge('A1:L1');
        spreadsheet.merge('A28:Z100');
        spreadsheet.merge('A27:AZ27');
        spreadsheet.merge('N1:N26');
        spreadsheet.merge('O1:AZ26');
        spreadsheet.merge('M1:M26');

        spreadsheet.setColumnsWidth(100, ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']);
        // spreadsheet.setColumnsWidth(100, ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']);
        // spreadsheet.setColumnsWidth(0, Array.from({ length: 163840 }, (_, i) => String.fromCharCode(84 + i)));
        spreadsheet.setColumnsHeight(150, ['1']);
        
      } catch (error) {
        console.error('Error updating spreadsheet:', error);
      }
    }
  };

  useEffect(() => {
    if (spreadsheetRef.current) {
      const spreadsheet = spreadsheetRef.current;

      spreadsheet.protectSheet(0, {
        selectCells: true,
        formatCells: false,
        formatRows: false,
        formatColumns: false,
        insertLink: false,
        insertShape: true,
        insertImage: true
      });

      spreadsheet.lockCells('A1:AB30', true); // Lock all cells initially
      editableRanges.forEach(range => {
        spreadsheet.lockCells(range, false); // Unlock the editable ranges
      });

      // Add beforeSelect event handler
      spreadsheet.beforeSelect = onBeforeSelect;

      // Add event listeners
      spreadsheet.cellSelected = onCellSelected;
      spreadsheet.element.addEventListener('keydown', onKeyDown);

      // Customize spreadsheet after it's created
      spreadsheet.created = () => {
        insertImage(); // Insert image into spreadsheet
        customizeSpreadsheet(); // Apply additional customizations
      };
    }

    // Cleanup event listener on component unmount
    return () => {
      if (spreadsheetRef.current) {
        const spreadsheet = spreadsheetRef.current;
        spreadsheet.element.removeEventListener('keydown', onKeyDown);
      }
    };
  }, []);

  return (
    <div style={{ height: '480px', width: '99%', backgroundColor: '#f0f8ff', padding: '5px' }}>
      <SpreadsheetComponent ref={spreadsheetRef} allowScrolling={true} showRibbon={false} showFormulaBar={true} showSheetTabs={false} style={{ height: '150px', width: '830px' }}>
        <SheetsDirective>
          <SheetDirective name="HR" showHeaders={false}>
            <RowsDirective>
              <RowDirective index={0} height={45}/>
              <RowDirective index={1} height={20}/>
            </RowsDirective>
            <ColumnsDirective>
              <ColumnDirective width={70}/>
              <ColumnDirective width={70}/>
            </ColumnsDirective>
          </SheetDirective>
        </SheetsDirective>
      </SpreadsheetComponent>

      {/* Fixed Bottom Pane */}
      <div style={{
        position: 'fixed',
        bottom: 0,
        width: '95vw', // Adjusted width to match container width
        backgroundColor: '#f0f8ff', // Light blue background color
        borderTop: '1px solid #ccc',
        padding: '8px',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        backgroundColor: '#bfd5e9',
        boxShadow: '0px -1px 5px rgba(0,0,0,0.2)' // Added shadow for better visibility
      }}>
        <div style={{ display: 'flex', gap: '15px' }}>
          <span>Pressure: {pressure}</span>
          <span>Temp: {temperature}</span>
          <span>Distance: {distance}</span>
          <span>Time: {time}</span>
          <span>Velocity: {velocity}</span>
          <span>Weight: {weight}</span>
        </div>
        <div style={{ display: 'flex', gap: '10px' }}>
          <button style={{ padding: '5px 10px', backgroundColor: '#e0e0e0', border: 'none', borderRadius: '4px' }} onClick={() => window.print()}>Print</button>
          <button style={{ padding: '5px 10px', backgroundColor: '#e0e0e0', border: 'none', borderRadius: '4px' }} onClick={() => alert('Save functionality not implemented.')}>Save</button>
          <button style={{ padding: '5px 10px', backgroundColor: '#e0e0e0', border: 'none', borderRadius: '4px' }} onClick={() => alert('Close functionality not implemented.')}>Close</button>
        </div>
      </div>
    </div>
  );
};

export default HR;
