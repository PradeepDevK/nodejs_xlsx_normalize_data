const express = require('express');
const app = express();
const router = express.Router();
const expressSession = require('express-session');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx')


app.use(cors());

app.use(express.urlencoded({extended: true}));
app.use(express.json());

app.use(function(req, res, next) {
    res.header('Access-Control-Allow-Credentials', true);
    res.header('Access-Control-Allow-Origin', req.headers.origin);
    res.header('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE');
    res.header('Access-Control-Allow-Headers', 'X-Requested-With, X-HTTP-Method-Override, Content-Type, Accept');
    if ('OPTIONS' == req.method) {
      res.send(200);
    } else {
        next();
    }
});

// Specify the directory for storing uploaded files
const uploadDirectory = path.join(__dirname, 'uploads');

// Ensure the directory exists
if (!fs.existsSync(uploadDirectory)){
    fs.mkdirSync(uploadDirectory, { recursive: true });
}

// Set up Multer for file uploads with dynamic folder creation based on user_email
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const username = req.params.username;
        if (!username) {
            return cb(new Error('username is required'), null);
        }

        const userUploadDirectory = path.join(uploadDirectory, username);

        // Ensure the user-specific directory exists
        if (!fs.existsSync(userUploadDirectory)){
            fs.mkdirSync(userUploadDirectory, { recursive: true });
        }

        cb(null, userUploadDirectory);
    },
    filename: (req, file, cb) => {
        // Use the original filename
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

// Test Endpoint
app.get('/', (req, res) => {
    res.send("Hello World!");
});

// File upload endpoint
app.post('/upload/:username', upload.single('file'), (req, res) => {

    if (!req.params.username) {
        return res.status(400).send('username is required....');
    }

    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    res.send(`File uploaded successfully: ${req.file.path}`);
});

app.post('/process-files', (req, res) => {
    const { columns, username } = req.body;
    const jsonResults = [];
    let serialNumber = 1; // Initialize the serial number counter

    const userUploadDirectory = path.join(uploadDirectory, username);

    if (!fs.existsSync(userUploadDirectory)){
        return res.status(500).send('Unable to scan directory: ' + err);
    }

    fs.readdir(userUploadDirectory, (err, files) => {
        if (err) {
            return res.status(500).send('Unable to scan directory: ' + err);
        }

        console.log(`Files found: ${files.length}`);
        if (files.length === 0) {
            return res.status(404).send('No files found in the uploads directory.');
        }

        files.forEach((file) => {
            console.log(`Processing file: ${file}`);
            const filePath = path.join(userUploadDirectory, file);
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const json = xlsx.utils.sheet_to_json(sheet);

            if (json.length) {
                const finalArray = json.map(item => {
                    let filteredItem = { SerialNumber: serialNumber++ }; // Add serial number
                    columns.forEach(key => {
                      if (item.hasOwnProperty(key)) {
                        filteredItem[key] = item[key];
                      } else {
                        filteredItem[key] = '';
                      }
                    });
                    return filteredItem;
                });
                jsonResults.push(...finalArray);
            }
        });

        if (jsonResults.length) {
            // Create a new workbook and add a worksheet
            const workbook = xlsx.utils.book_new();
            const worksheet = xlsx.utils.json_to_sheet(jsonResults);
            xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");

            const filePath = path.join(uploadDirectory, 'finalData.xlsx');
            // Ensure the directory exists
            if (!fs.existsSync(uploadDirectory)){
                fs.mkdirSync(uploadDirectory, { recursive: true });
            }

            // Write the workbook to a file
            xlsx.writeFile(workbook, filePath);

            // Send the Excel file back to the client
            res.download(filePath, 'finalData.xlsx', (err) => {
                if (err) {
                    res.status(500).send('Error sending the file: ' + err);
                }

                // Optionally, delete the file after sending it
                fs.unlink(filePath, (err) => {
                    if (err) {
                        console.error('Error deleting the file: ' + err);
                    }
                });
            });
        } else {
            res.json("No Data Found!");
        }
    });
});

app.post('/deleteDirectory', async (req, res) => {
    const { username } = req.body;

    const userUploadDirectory = path.join(uploadDirectory, username);

    if (!fs.existsSync(userUploadDirectory)){
        return res.status(500).send('Unable to scan directory: ' + err);
    }

    try {
        await fs.rm(userUploadDirectory, { recursive: true, force: true });
        return res.send(`Successfully deleted ${userUploadDirectory}`);
    } catch (error) {
        return res.status(500).send(`Error while deleting ${userUploadDirectory}.`, error);
    }
});


const PORT = process.env.PORT || 8000;

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});