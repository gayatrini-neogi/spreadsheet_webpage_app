import React, { useRef, useEffect } from 'react';
import { SpreadsheetComponent, SheetsDirective, SheetDirective } from '@syncfusion/ej2-react-spreadsheet';
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

const Inputs = ({ pressure, temperature, distance, time, velocity, weight }) => {
  const spreadsheetRef = useRef(null);
  const editableRanges = [
   'A2', 'A3', 'A4', 'C2', 'C3', 'C4', 'D2', 'D3', 'D4', 
    'F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'L3', 'M3', 'N3', 'O3', 'P3', 'A15', 'B15', 'C15', 'D15', 'E15', 'F15', 'A17', 'B17', 'C17', 'D17', 'E17', 'F17', 'A19', 'B19', 'C19', 'D19', 'E19', 'F19',
    'S2', 'S3', 'C10', 'F11', 'H11', 'J11', 'L11', 'N11', 'P11', 'R11', 'H15', 'J15', 'L15', 'N15', 'P15', 'R15', 'H17', 'J17', 'L17', 'N17', 'P17', 'R17' 
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

  const insertImage = () => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet) {
      // Image details
      const image1 = {
        src: 'https://i.postimg.cc/y8NYvCTk/block-1.png', // Use the provided image URL
        width: 320,  // Adjust width to fit within B10:B14
        height: 95 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image1], 'F6'); // Start from the upper-left corner of the range
      const image2 = {
        src: 'https://i.postimg.cc/SNXkPs7B/Screenshot-177.png', // Use the provided image URL
        width: 1450,  // Adjust width to fit within B10:B14
        height: 100 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image2], 'K5'); // Start from the upper-left corner of the range
      
      const image3 = {
        src: 'https://i.postimg.cc/sDdYNB58/brackets-1.png', // Use the provided image URL
        width: 280,  // Adjust width to fit within B10:B14
        height: 20 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image3], 'N13'); // Start from the upper-left corner of the range

       const image4 = {
         src: 'https://i.postimg.cc/sDdYNB58/brackets-1.png', // Use the provided image URL
         width: 240,  // Adjust width to fit within B10:B14
         height: 20 // Adjust height to fit within B10:B14
       };
 
       // Insert the image into the specified range B10:B14
       spreadsheet.insertImage([image4], 'R13'); // Start from the upper-left corner of the range

       const image5 = {
         src: 'https://i.postimg.cc/sDdYNB58/brackets-1.png', // Use the provided image URL
         width: 240,  // Adjust width to fit within B10:B14
         height: 20 // Adjust height to fit within B10:B14
       };
 
       // Insert the image into the specified range B10:B14
       spreadsheet.insertImage([image5], 'U13'); // Start from the upper-left corner of the range

       const image6 = {
        src: 'https://i.postimg.cc/sDdYNB58/brackets-1.png', // Use the provided image URL
        width: 240,  // Adjust width to fit within B10:B14
        height: 20 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image6], 'X13'); // Start from the upper-left corner of the range

      const image7 = {
        src: 'https://i.postimg.cc/sDdYNB58/brackets-1.png', // Use the provided image URL
        width: 240,  // Adjust width to fit within B10:B14
        height: 20 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image7], 'AA13'); // Start from the upper-left corner of the range

      const image8 = {
        src: 'https://i.postimg.cc/sDdYNB58/brackets-1.png', // Use the provided image URL
        width: 240,  // Adjust width to fit within B10:B14
        height: 20 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image8], 'AD13'); // Start from the upper-left corner of the range

      const image9 = {
        src: 'https://i.postimg.cc/RhvyBv8x/download.png', // Use the provided image URL
        width: 980,  // Adjust width to fit within B10:B14
        height: 10 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image9], 'A14'); // Start from the upper-left corner of the range

      const image10 = {
        src: 'https://i.postimg.cc/RhvyBv8x/download.png', // Use the provided image URL
        width: 960,  // Adjust width to fit within B10:B14
        height: 10 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image10], 'Q14'); // Start from the upper-left corner of the range
      
      const image11 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 480,  // Adjust width to fit within B10:B14
        height: 70 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image11], 'F12'); // Start from the upper-left corner of the range

      const image12 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 320,  // Adjust width to fit within B10:B14
        height: 210 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image12], 'A5'); // Start from the upper-left corner of the range
      
      const image13 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 2300,  // Adjust width to fit within B10:B14
        height: 19 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image13], 'A15'); // Start from the upper-left corner of the range
      
      const image14 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 2300,  // Adjust width to fit within B10:B14
        height: 21 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image14], 'A1'); // Start from the upper-left corner of the range
      
      const image15 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 340,  // Adjust width to fit within B10:B14
        height: 19 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image15], 'F5'); // Start from the upper-left corner of the range
      
      const image16 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 2300,  // Adjust width to fit within B10:B14
        height: 21 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image16], 'A17'); // Start from the upper-left corner of the range
      
      const image17 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 1850,  // Adjust width to fit within B10:B14
        height: 37 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image17], 'H10'); // Start from the upper-left corner of the range
      
      const image18 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 1450,  // Adjust width to fit within B10:B14
        height: 60 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image18], 'K4'); // Start from the upper-left corner of the range
      
      const image19 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 2300,  // Adjust width to fit within B10:B14
        height: 23 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image19], 'A19'); // Start from the upper-left corner of the range

      const image20 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 1520,  // Adjust width to fit within B10:B14
        height: 20 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image20], 'K2'); // Start from the upper-left corner of the range

      const image21 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 2300,  // Adjust width to fit within B10:B14
        height: 23 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image21], 'A21'); // Start from the upper-left corner of the range
      
      const image22 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 160,  // Adjust width to fit within B10:B14
        height: 23 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image22], 'K3'); // Start from the upper-left corner of the range
      
      const image23 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 193,  // Adjust width to fit within B10:B14
        height: 23 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image23], 'C3'); // Start from the upper-left corner of the range
      
      const image24 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 193,  // Adjust width to fit within B10:B14
        height: 23 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image24], 'C2'); // Start from the upper-left corner of the range
      
      const image25 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 193,  // Adjust width to fit within B10:B14
        height: 23 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image25], 'C4'); // Start from the upper-left corner of the range
      
      const image26 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 1500,  // Adjust width to fit within B10:B14
        height: 40 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image26], 'L13'); // Start from the upper-left corner of the range

      const image27 = {
        src: 'https://i.postimg.cc/gj2yXqDH/picture-picture.png', // Use the provided image URL
        width: 159,  // Adjust width to fit within B10:B14
        height: 100 // Adjust height to fit within B10:B14
      };

      // Insert the image into the specified range B10:B14
      spreadsheet.insertImage([image27], 'K10'); // Start from the upper-left corner of the range
      
    }
  };
  
  const customizeSpreadsheet = () => {
    const spreadsheet = spreadsheetRef.current;
    if (spreadsheet) {
      try {
        // Update and style the header cell
        spreadsheet.updateCell({
          value: 'Mold Temp Notes',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A1');
        spreadsheet.updateCell({
          value: 'Sections',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B1');
        spreadsheet.updateCell({
          value: 'Mold Temp Inputs',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C1');
        
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A2');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A4');
        spreadsheet.updateCell({
          value: '3',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B2');
        spreadsheet.updateCell({
          value: '2',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B3');
        spreadsheet.updateCell({
          value: '1',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B4');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C2');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C4');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D2');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D4');
        spreadsheet.updateCell({
          value: 'Nozzle',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            border: '1.5px solid black'
          }
        }, 'F2');
        spreadsheet.updateCell({
          value: 'H1',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'G2');
        spreadsheet.updateCell({
          value: 'H2',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H2');
        spreadsheet.updateCell({
          value: 'H3',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'I2');
        spreadsheet.updateCell({
          value: 'H4',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J2');
        spreadsheet.updateCell({
          value: 'H5',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'K2');
        spreadsheet.updateCell({
          value: 'H6',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L2');
        spreadsheet.updateCell({
          value: 'H7',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'M2');
        spreadsheet.updateCell({
          value: 'H8',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N2');
        spreadsheet.updateCell({
          value: 'H9',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          
          }
        }, 'O2');
        spreadsheet.updateCell({
          value: 'H10',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P2');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'G3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'I3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'K3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'M3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'O3');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P3');
        spreadsheet.updateCell({
          value: 'B-Side',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C5');
        spreadsheet.updateCell({
          value: 'A-Side',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D5');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
            
          }
        }, 'C10');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
            
          }
        }, 'F11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
            
          }
        }, 'H11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
            
          }
        }, 'J11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
            
          }
        }, 'L11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R11');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'S2');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'S3');
        spreadsheet.updateCell({
          value: 'Cooling Time',
          style: {
            textAlign: 'left',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            borderTop: '1.5px solid black',
            borderBottom: '1.5px solid black',
            borderLeft: '1.5px solid black',
            borderRight: '1.5px solid black'
          }
        }, 'B10');
        spreadsheet.updateCell({
          value: 'Sec',
          style: {
            textAlign: 'left',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D10');
        spreadsheet.updateCell({
          value: 'Cushion*',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white', 
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F10');
        spreadsheet.updateCell({
          value: 'Transfer*',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H10');
        spreadsheet.updateCell({
          value: 'Posn 4',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J10');
        spreadsheet.updateCell({
          value: 'Posn 3',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L10');
        spreadsheet.updateCell({
          value: 'Posn 2',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N10');
        spreadsheet.updateCell({
          value: 'Posn 1',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P10');
        spreadsheet.updateCell({
          value: 'Shot Size*',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R10');
        spreadsheet.updateCell({
          value: 'Decompression',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R1');
        spreadsheet.updateCell({
          value: 'Distance:',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R2');
        spreadsheet.updateCell({
          value: 'Speed:',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R3');
        spreadsheet.updateCell({
          value: 'Pack & Hold Phase',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A14');
        spreadsheet.updateCell({
          value: 'Pack & Hold Phase',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black',
            borderTop: '1.5px solid black'
          }
        }, 'B14');
        spreadsheet.updateCell({
          value: 'Pack & Hold Phase',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            borderTop: '1.5px solid black'
          }
        }, 'C14');
        spreadsheet.updateCell({
          value: 'Pack & Hold Phase',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D14');
        spreadsheet.updateCell({
          value: 'Pack & Hold Phase',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E14');
        spreadsheet.updateCell({
          value: 'Pack & Hold Phase',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'I14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'K14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'M14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'O14');
        spreadsheet.updateCell({
          value: 'Injection Phase (Max. Inj. Speed)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P14');
        
        spreadsheet.updateCell({
          value: 'Barrel Settings Temperatures (image not to scale)',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F1');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F15');
  
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A19');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B19');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C19');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D19');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E19');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F19');
  
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P15');
      
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R17');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            borderTop: '1.5px solid black'
          }
        }, 'S15');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'rgb(241, 217, 201)',
            color: 'black',
            borderTop: '1.5px solid black'
          }
        }, 'R15');
        spreadsheet.updateCell({
          value: 'P/H Vel 6',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A16');
        spreadsheet.updateCell({
          value: 'P/H Vel 5',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B16');
        spreadsheet.updateCell({
          value: 'P/H Vel 4',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C16');
        spreadsheet.updateCell({
          value: 'P/H Vel 3',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D16');
        spreadsheet.updateCell({
          value: 'P/H Vel 2',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E16');
        spreadsheet.updateCell({
          value: 'P/H Vel 1',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F16');
        spreadsheet.updateCell({
          value: 'P/H Time 6',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A20');
        spreadsheet.updateCell({
          value: 'P/H Time 5',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B20');
        spreadsheet.updateCell({
          value: 'P/H Time 4',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C20');
        spreadsheet.updateCell({
          value: 'P/H Time 3',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D20');
        spreadsheet.updateCell({
          value: 'P/H Time 2',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E20');
        spreadsheet.updateCell({
          value: 'P/H Time 1',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F20');
  
        spreadsheet.updateCell({
          value: 'Inj Vel 5',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H16');
        spreadsheet.updateCell({
          value: 'Inj Vel 4',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J16');
        spreadsheet.updateCell({
          value: 'Inj Vel 3',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L16');
        spreadsheet.updateCell({
          value: 'Inj Vel 2',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N16');
        spreadsheet.updateCell({
          value: 'Inj Vel 1',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P16');
        spreadsheet.updateCell({
          value: 'Screw Rot Speed*',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1px solid black'
          }
        }, 'R16');
        spreadsheet.updateCell({
          value: 'Inj Press 5',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'H18');
        spreadsheet.updateCell({
          value: 'Inj Press 4',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'J18');
        spreadsheet.updateCell({
          value: 'Inj Press 3',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'L18');
        spreadsheet.updateCell({
          value: 'Inj Press 2',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'N18');
        spreadsheet.updateCell({
          value: 'Inj Press 1',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'P18');
        spreadsheet.updateCell({
          value: 'Back Pressure*',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'R18');
  
        spreadsheet.updateCell({
          value: 'P/H Press 6',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'A18');
        spreadsheet.updateCell({
          value: 'P/H Press 5',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'B18');
        spreadsheet.updateCell({
          value: 'P/H Press 4',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'C18');
        spreadsheet.updateCell({
          value: 'P/H Press 3',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'D18');
        spreadsheet.updateCell({
          value: 'P/H Press 2',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'E18');
        spreadsheet.updateCell({
          value: 'P/H Press 1',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'F18');
        spreadsheet.updateCell({
          value: '',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'G10');
        spreadsheet.updateCell({
          value: '',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'I10');
        spreadsheet.updateCell({
          value: '',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'K10');
        spreadsheet.updateCell({
          value: '',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'M10');
        spreadsheet.updateCell({
          value: '',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'O10');
        spreadsheet.updateCell({
          value: '',
          style: {
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            border: '1.5px solid black'
          }
        }, 'Q10');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            borderBottom: '1.5px solid black'
          }
        }, 'I18');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            borderBottom: '1.5px solid black'
          }
        }, 'K18');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            borderBottom: '1.5px solid black'
          }
        }, 'M18');
        spreadsheet.updateCell({
          value: '',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '12pt',
            backgroundColor: 'white',
            color: 'black',
            borderBottom: '1.5px solid black'
          }
        }, 'O18');
        spreadsheet.updateCell({
          value: 'Process Inputs',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '20pt',
            backgroundColor: 'white',
            color: 'black',
            border: 'none', 
            
          }
        }, 'A6');
        spreadsheet.updateCell({
          value: 'The Fields marked with * are mandatory',
          style: {
            fontWeight: 'bold',
            textAlign: 'center',
            verticalAlign: 'middle',
            fontSize: '11pt',
            backgroundColor: 'white',
            color: 'black',
            border: 'none', 
            
          }
        }, 'A7');

        //WHITE BOARD BACKGROUND FOR CELLS
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A5');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'B5');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A8');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'B8');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'A9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A11');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A10');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'B11');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'C11');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'D11');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E11');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderTop: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E10');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderTop: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderTop: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E4');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'F4');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'F9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'G9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'H9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'I9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'J9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'K9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'L9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'M9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'N9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'O9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'P9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E1');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E2');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E3');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'Q1');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'Q2');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'Q3');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S5');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S6');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S7');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S8');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'B12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'C12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'D12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'E12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'F12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'B13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'C13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'D13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'E13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'F13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'G13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'H13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'I13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'J13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'K13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'L13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'M13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'N13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'O13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'P13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'Q13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'R13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S13');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'G4');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S4');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'I5');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'Q9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'I6');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'I7');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'I8');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderLeft: 'none',           // Remove borders
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'R9');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A22');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S10');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S11');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'R12');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderLeft: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S14');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'S14');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'Q14');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'Q16');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'Q17');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'Q18');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderRight: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'Q15');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'I15');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'I16');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'I17');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'I18');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'k15');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'k16');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'k17');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'k18');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'M15');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'M16');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'M17');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'M18');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'O15');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'O16');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'O17');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'O18');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'G14');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'G15');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'G16');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'G17');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            borderTop: 'none',           // Remove borders
            borderBottom: 'none',
            color: 'transparent'      // Hide text color
          }
        }, 'G18');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'G21');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'G19');
        spreadsheet.updateCell({
          style: {
            backgroundColor: 'white', // Match background color
            border: 'none',           // Remove borders
            color: 'transparent'      // Hide text color
          }
        }, 'A21');
    
        
        // Merge specified cell ranges
        spreadsheet.merge('I8:R8');
        spreadsheet.merge('A21:F21');
        spreadsheet.merge('G21:S21');
        spreadsheet.merge('A22:S100');
        spreadsheet.merge('I7:R7');
        spreadsheet.merge('I6:R6');
        spreadsheet.merge('I5:R5');
        spreadsheet.merge('G4:R4');
        spreadsheet.merge('C1:D1');
          spreadsheet.merge('F1:P1');
          spreadsheet.merge('R1:S1');
          spreadsheet.merge('G10:G11');
          spreadsheet.merge('I10:I11');
          spreadsheet.merge('K10:K11');
          spreadsheet.merge('M10:M11');
          spreadsheet.merge('O10:O11');
          spreadsheet.merge('Q10:Q11');
          spreadsheet.merge('A14:F14');
          spreadsheet.merge('H14:P14');
          spreadsheet.merge('R16:S16');
          spreadsheet.merge('R18:S18');
          spreadsheet.merge('R15:S15');
          spreadsheet.merge('R17:S17');
          spreadsheet.merge('A6:B6');
          spreadsheet.merge('A7:B7');
          spreadsheet.merge('F5:F8');
          spreadsheet.merge('G19:S20');
          spreadsheet.merge('C6:C8');

        spreadsheet.setColumnsWidth(160, ['A', 'B', 'C', 'D', 'E', 'F']);
        spreadsheet.setColumnsWidth(100, ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']);
        spreadsheet.setColumnsWidth(0, Array.from({ length: 163840 }, (_, i) => String.fromCharCode(84 + i)));
        spreadsheet.setColumnsHeight(150, ['1', '2', '3', '4', '5']);

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
    <div style={{ height: '81vh', width: '97vw' }}>
      <SpreadsheetComponent
        ref={spreadsheetRef}
        style={{ height: '90%', width: '90%' }}
        allowEditing={false} // Disable editing by default
        showSheetTabs={false}
        showFormulaBar={true}
        showRibbon={false}
        showHeaders={false}
      >
        <SheetsDirective>
          <SheetDirective showHeaders={false}>
            {/* Define your cells and other sheet content here */}
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

export default Inputs;
