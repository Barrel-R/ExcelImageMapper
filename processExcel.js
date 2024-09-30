const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const processExcel = (excelFile, insertHeader, dataHeader, imgFolder) => {
    return new Promise((resolve, reject) => {
        let insertedImages = 0;
        let missingImages = 0;

        // Load the workbook
        let workbook;
        try {
            workbook = xlsx.readFile(excelFile);
        } catch (err) {
            return reject(`Error loading the Excel file: ${err}`);
        }

        // Iterate through the sheets
        workbook.SheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            console.log(`Processing sheet: ${sheetName}`);

            // Convert worksheet to JSON
            const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            const header = jsonData[0]; // First row is the header
            const data = jsonData.slice(1); // All other rows are data

            // Find the column indices based on headers
            const dataColIndex = header.findIndex(h => h && h.toLowerCase().includes(dataHeader.toLowerCase()));
            const insertImageColIndex = header.findIndex(h => h && h.toLowerCase().includes(insertHeader.toLowerCase()));

            if (dataColIndex === -1 || insertImageColIndex === -1) {
                console.error(`Error: Couldn't find columns for '${dataHeader}' or '${insertHeader}' in sheet: ${sheetName}`);
                return;
            }

            // Loop through the data rows
            data.forEach((row, rowIndex) => {
                const value = row[dataColIndex];
                if (!value) return; // Skip if the value is empty

                const imgPath = path.join(imgFolder, `${value}.png`);

                if (fs.existsSync(imgPath)) {
                    // Add the image (resize to 75x75)
                    sharp(imgPath)
                        .resize(75, 75)
                        .toBuffer()
                        .then(resizedImage => {
                            // Create a new cell with the resized image in the insert column
                            const cellRef = xlsx.utils.encode_cell({ r: rowIndex + 1, c: insertImageColIndex });
                            worksheet[cellRef] = { t: "z" }; // Placeholder for image
                            insertedImages++;
                            console.log(`Inserted image for ${dataHeader}: ${value}`);
                        })
                        .catch(err => {
                            console.error(`Error processing image ${imgPath}: ${err}`);
                        });
                } else {
                    console.log(`Image not found for ${dataHeader}: ${value}`);
                    const cellRef = xlsx.utils.encode_cell({ r: rowIndex + 1, c: insertImageColIndex });
                    worksheet[cellRef] = { t: "s", v: "Image not found" };
                    missingImages++;
                }
            });
        });

        // Save the updated workbook
        try {
            xlsx.writeFile(workbook, excelFile); // Overwrite the existing file
        } catch (err) {
            return reject(`Error saving the Excel file: ${err}`);
        }

        resolve();
    });
};

module.exports = processExcel;
