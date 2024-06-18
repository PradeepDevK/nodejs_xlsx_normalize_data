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

// Set up Multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadDirectory);
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
app.post('/upload', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    res.send(`File uploaded successfully: ${req.file.path}`);
});

app.post('/process-files', (req, res) => {
    const {columns} = req.body;
    const jsonResults = [];

    fs.readdir(uploadDirectory, (err, files) => {
        if (err) {
            return res.status(500).send('Unable to scan directory: ' + err);
        }

        console.log(`Files found: ${files.length}`);
        if (files.length === 0) {
            return res.status(404).send('No files found in the uploads directory.');
        }

        files.forEach((file) => {
            console.log(`Processing file: ${file}`);
            const filePath = path.join(uploadDirectory, file);
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const json = xlsx.utils.sheet_to_json(sheet);

            if (json.length) {
                const finalArray = json.map(item => {
                    let filteredItem = {};
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

            res.json(jsonResults);
        } else {
            res.json("No Data Found!");
        }
    });
});




const PORT = process.env.PORT || 8000;

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});