const XLSX = require( 'xlsx' );


// Load the Excel file
const workbook = XLSX.readFile( 'Assignment_Timecard.xlsx' );
const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
const worksheet = workbook.Sheets[sheetName];


// converts the time to hours format for comparison
function timeToHoursDecimal ( timeStr ) {
    const [hoursStr, minutesStr] = timeStr.split( ':' );
    const hours = parseInt( hoursStr, 10 );
    const minutes = parseInt( minutesStr, 10 );
    const hoursDecimal = hours + minutes / 60;
    return hoursDecimal;
}


// Process the data
const rows = XLSX.utils.sheet_to_json( worksheet );


for ( const record of rows )
{

    const employeeName = record['Employee Name'];

    // Calculate time difference in hours
    const timeDiffHours = record['Timecard Hours (as Time)'];
    const timeInHours = timeToHoursDecimal( timeDiffHours );




    if ( timeInHours > 1 && timeInHours < 10 )
    {

        console.log( `${employeeName} has less than 10 hours between shifts.` );
    }

    if ( timeInHours > 14 )
    {
        console.log( `${employeeName} worked for more than 14 hours in a single shift.` );
    }


}
