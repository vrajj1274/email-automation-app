// app.js - Main application file for serverless environment
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const session = require('express-session');

const app = express();
const PORT = process.env.PORT || 3000;

// Set up EJS as the view engine
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Set up session middleware
app.use(session({
    secret: 'email-automation-secret',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: process.env.NODE_ENV === 'production' }
}));

// Configure multer for in-memory file uploads
const upload = multer({
    storage: multer.memoryStorage(),
    fileFilter: (req, file, cb) => {
        const filetypes = /xlsx|xls/;
        const extname = filetypes.test(
            file.originalname.split('.').pop().toLowerCase()
        );
        if (extname) {
            return cb(null, true);
        } else {
            cb('Error: Excel files only!');
        }
    }
});

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
        // Read and parse Excel file from buffer
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        if (data.length === 0) {
            return res.status(400).send('Excel file has no data');
        }

        // Get first row for preview
        const firstRow = data[0];

        // Store data in session for later use
        req.session.emailData = {
            workbookData: data,
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

        // Render preview page
        res.render('preview', {
            firstRow: firstRow,
            allColumns: Object.keys(firstRow),
            hasEmailColumn: firstRow.hasOwnProperty('email'),
            messageTemplate: messageTemplate,
            emailSubject: emailSubject || 'Personalized Message',
            previewMessage: previewMessage,
            previewSubject: previewSubject
        });
    } catch (error) {
        res.status(500).send(`Error processing file: ${error.message}`);
    }
});

// Handle email sending
app.post('/send-emails', upload.single('excelFile'), async (req, res) => {
    const { email, appPassword, messageTemplate, emailSubject } = req.body;
    const subject = emailSubject || 'Personalized Message';

    if (!email || !appPassword || !messageTemplate) {
        return res.status(400).send('All fields are required');
    }

    try {
        let data;

        // If we have a new file upload, process it
        if (req.file) {
            const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            data = XLSX.utils.sheet_to_json(worksheet);
        }
        // Otherwise use data from session (from preview)
        else if (req.session && req.session.emailData && req.session.emailData.workbookData) {
            data = req.session.emailData.workbookData;
        }
        else {
            return res.status(400).send('No data found. Please upload an Excel file.');
        }

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

        // Clear session data
        if (req.session) {
            req.session.emailData = null;
        }

        // Render results page
        res.render('results', { results });
    } catch (error) {
        res.status(500).send(`Error processing request: ${error.message}`);
    }
});

// Export the express app
module.exports = app;

// Only listen if not being imported
if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`Server running at http://localhost:${PORT}`);
    });
}