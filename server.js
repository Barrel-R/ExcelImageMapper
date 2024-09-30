const express = require('express');
const multer = require('multer');
const processExcel = require('./processExcel'); // Assuming your logic is in processExcel.js

const app = express();
const upload = multer({ dest: 'uploads/' }); // Temporary file storage

app.use(express.static('public')); // Serve the frontend HTML/JS

// Handle form submission
app.post('/process', upload.any(), (req, res) => {
    const folder = req.files.filter(f => f.fieldname === 'folder[]');
    const excelFile = req.files.find(f => f.fieldname === 'file');
    const insertHeader = req.body.insertHeader;
    const dataHeader = req.body.dataHeader;

    // Verify the file and folder exist before proceeding
    if (!excelFile || folder.length === 0) {
        return res.status(400).json({ status: 'error', message: 'Missing file or folder' });
    }

    // Directory where images are stored (just an example, adjust accordingly)
    const imgFolder = folder[0].path;

    // Call your Node.js processing function
    processExcel(excelFile.path, insertHeader, dataHeader, imgFolder)
        .then(() => {
            res.json({ status: 'success' });
        })
        .catch((error) => {
            console.error('Error during processing:', error);
            res.status(500).json({ status: 'error', message: error.message });
        });
});

app.listen(3000, () => {
    console.log('Server running on http://localhost:3000');
});
