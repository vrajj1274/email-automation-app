// app.js - Main application file
const express = require('express');
const multer = require('multer');
const path = require('path');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Set up EJS as the view engine
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const filetypes = /xlsx|xls/;
        const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
        if (extname) {
            return cb(null, true);
        } else {
            cb('Error: Excel files only!');
        }
    }
});

// Ensure uploads directory exists
if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}

// Home page
app.get('/', (req, res) => {
    res.render('index');
});

// Handle file upload and preview first row
app.post('/preview', upload.single('excelFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded');
    }

    const { messageTemplate, emailSubject } = req.body;

    if (!messageTemplate) {
        return res.status(400).send('Message template is required');
    }

    try {
        // Read and parse Excel file
        const workbook = XLSX.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        if (data.length === 0) {
            fs.unlinkSync(req.file.path); // Clean up file
            return res.status(400).send('Excel file has no data');
        }

        // Get first row for preview
        const firstRow = data[0];

        // Store file path in session or temporary storage for later use
        const tempData = {
            filePath: req.file.path,
            firstRow: firstRow,
            allColumns: Object.keys(firstRow),
            hasEmailColumn: firstRow.hasOwnProperty('email'),
            messageTemplate: messageTemplate,
            emailSubject: emailSubject || 'Personalized Message'
        };

        // Create personalized message for preview
        let previewMessage = messageTemplate;
        let previewSubject = emailSubject || 'Personalized Message';

        Object.keys(firstRow).forEach(key => {
            const placeholder = new RegExp(`{{${key}}}`, 'g');
            previewMessage = previewMessage.replace(placeholder, firstRow[key]);
            previewSubject = previewSubject.replace(placeholder, firstRow[key]);
        });

        tempData.previewMessage = previewMessage;
        tempData.previewSubject = previewSubject;

        // Render preview page
        res.render('preview', tempData);
    } catch (error) {
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path); // Clean up file
        }
        res.status(500).send(`Error processing file: ${error.message}`);
    }
});

// Handle file upload and form submission
app.post('/send-emails', upload.single('excelFile'), async (req, res) => {
    // Check if this is a continuation from preview (no new file)
    let filePath = req.body.existingFilePath;

    // If no existing file path or it doesn't exist, use the new upload
    if (!filePath || !fs.existsSync(filePath)) {
        if (!req.file) {
            return res.status(400).send('No file uploaded');
        }
        filePath = req.file.path;
    }

    const { email, appPassword, messageTemplate, emailSubject } = req.body;
    const subject = emailSubject || 'Personalized Message';

    if (!email || !appPassword || !messageTemplate) {
        return res.status(400).send('All fields are required');
    }

    try {
        // Read and parse Excel file
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // Setup email transporter
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: email,
                pass: appPassword
            }
        });

        // Track results
        const results = [];

        // Process each row in the Excel file
        for (const row of data) {
            // Skip if no email in this row
            if (!row.email) {
                results.push({ recipient: 'Unknown', status: 'Failed', error: 'No email address found' });
                continue;
            }

            // Create personalized message by replacing placeholders
            let personalizedMessage = messageTemplate;
            let personalizedSubject = subject;

            Object.keys(row).forEach(key => {
                const placeholder = new RegExp(`{{${key}}}`, 'g');
                personalizedMessage = personalizedMessage.replace(placeholder, row[key]);
                personalizedSubject = personalizedSubject.replace(placeholder, row[key]);
            });

            // Email options
            const mailOptions = {
                from: email,
                to: row.email,
                subject: personalizedSubject,
                text: personalizedMessage
            };

            // Send email
            try {
                await transporter.sendMail(mailOptions);
                results.push({ recipient: row.email, status: 'Success' });
            } catch (error) {
                results.push({ recipient: row.email, status: 'Failed', error: error.message });
            }
        }

        // Clean up - remove the uploaded file
        fs.unlinkSync(filePath);

        // Render results page
        res.render('results', { results });
    } catch (error) {
        // Clean up file in case of error
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
        }
        res.status(500).send(`Error processing request: ${error.message}`);
    }
});

app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});