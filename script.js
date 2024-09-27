let extractedData = []; // Array to hold extracted data

document.getElementById('file-input').addEventListener('change', handleFile, false);
document.getElementById('export-btn').addEventListener('click', exportToExcel, false);

function handleFile(event) {
    const fileInput = event.target;
    const file = fileInput.files[0]; // Get the selected file
    if (!file) {
        return;
    }

    const reader = new FileReader(); // Create a FileReader instance
    reader.onload = function (e) {
        const arrayBuffer = e.target.result; // Get the result from FileReader

        // Read the workbook
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Get the first sheet name
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName]; // Access the first sheet

        // Convert the sheet to JSON format
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });

        // Extract only the specified columns
        extractedData = jsonData.map(row => ({
            tktnbr: row["tktnbr"],
            pax: row["pax"],
            pnr: row["pnr"],
            Issuer: row["Issuer"],
            agent: row["agent"],
            IssueDate: row["Issue Date"],
            ScheduledFlightDate: row["Scheduled Flight Date"],
            OperatedFlightDate: row["Operated Flight Date"],
            origin: row["origin"],
            amount: row["cons_amount"],
            YQ: row["YQ"]
        })).filter(row => Object.keys(row).length > 0); // Filter out empty rows

        // Display the extracted data in the output element
        document.getElementById('output').textContent = JSON.stringify(extractedData, null, 2);

        // Show the export button
        document.getElementById('export-btn').style.display = 'block';

        // Send extracted data to the API (optional)
        axios.post('http://localhost:3000/api/endpoint', extractedData)
            .then(response => {
                console.log('Data sent successfully:', response.data);
            })
            .catch(error => {
                console.error('Error sending data:', error);
            });
    };

    reader.readAsArrayBuffer(file); // Read the file as an ArrayBuffer
}

function exportToExcel() {
    if (extractedData.length === 0) {
        alert('No data to export!');
        return;
    }

    // Convert JSON data to a worksheet
    const worksheet = XLSX.utils.json_to_sheet(extractedData);
    
    // Create a new workbook and append the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Extracted Data');

    // Export the workbook as an Excel file
    XLSX.writeFile(workbook, 'extracted_data.xlsx');
}
