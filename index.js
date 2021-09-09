const express = require('express');
const multer = require('multer');
const path = require("path");

const xlsxReader = require('read-excel-file/node');

const app = express();

//Exception Handling
process.on("uncaughtException", (err) => {
    console.log("UNCAUGHT EXCEPTION, APP SHUTTING NOW!!");
    console.log(err.message, err.name);
    process.exit(1);
});

//body parser
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

//File Uploading Configuration
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, "uploads");
    },
    filename: function (req, file, cb) {
        cb(null, file.originalname);
    }
});

const maxFileSize = 1 * 1024 * 1024;

const upload = multer({
    storage,
    limits: { fileSize: maxFileSize },
    fileFilter: function (req, file, cb) {
        var filetypes = /xlsx|xls/;
        var extname = filetypes.test(path.extname(file.originalname).toLocaleLowerCase());
        
        if (extname) {
            return cb(null, true);
        }
        cb("Error. File upload only supports these filetypes: " + filetypes);
    }
}).single("extract_sheet");

app.get("/", (req, res) => {
    res.json(
        {
            message: "Server is running on port 4000",
            status: 'UP'
        });
});

app.post("/upload-excel-sheet", async (req, res) => {

    await upload(req, res, function (err) {
        if (err) {
            res.status(500).json({ error: err });
        }
        else {
            res.status(200).json({ message: "Upload Success", status: "OK" });
        }
    });
});

app.get("/read-excel-sheet",async (req,res)=>{

     xlsxReader("./uploads/file_example_XLSX_50.xlsx")
     .then(rows=>{
         rows.forEach(element => {
             console.log(element);
         });
        res.json(rows);
     });
});

const PORT = 4000 || process.env.PORT;

app.listen(PORT, () => {
    console.log(`Server is running on Port ${PORT}`);
});
