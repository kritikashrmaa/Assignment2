const ExcelJS = require( 'exceljs' );

// Create a new ExcelJS workbook
const workbook = new ExcelJS.Workbook();

// Load the Excel file
workbook.xlsx.readFile( 'Assignment_Timecard.xlsx' )
    .then( () => {
        // Assuming data is in the first worksheet
        const worksheet = workbook.getWorksheet( 1 );

        // Create a map to track the consecutive working days for each employee
        const consecutiveDaysMap = new Map();

        // Iterate through rows
        worksheet.eachRow( ( row, rowNumber ) => {
            const employeeName = row.getCell( 'H' ).text; // Assuming employee name is in column H
            const position = row.getCell( 'A' ).text; // Assuming position is in column A
            const dateCell = row.getCell( 'C' ); // Assuming date is in column C
            const dateValue = new Date( dateCell.text );

            // Check if the date is valid and not a weekend (adjust this logic as needed)
            if ( !isNaN( dateValue ) && dateValue.getDay() !== 0 && dateValue.getDay() !== 6 )
            {
                // Initialize or increment consecutive working days for each employee
                const consecutiveDays = consecutiveDaysMap.get( employeeName ) || 0;
                consecutiveDaysMap.set( employeeName, consecutiveDays + 1 );
            } else
            {
                // Reset consecutive working days if it's a weekend or invalid date
                consecutiveDaysMap.set( employeeName, 0 );
            }

            // Check if an employee has worked for 7 consecutive days
            if ( consecutiveDaysMap.get( employeeName ) === 7 )
            {
                console.log( `${employeeName} (Position: ${position}) has worked for 7 consecutive days.` );
            }
        } );
    } )
    .catch( error => {
        console.error( 'Error reading the Excel file:', error );
    } );
