const express = require("express");
const XlsxPopulate = require("xlsx-populate");
// const fileUpload = require("express-fileupload");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const { parse, stringify, toJSON, fromJSON } = require("flatted");

const app = express();
// app.use(fileUpload());


const storage = multer.diskStorage({
  destination:(req, res, cb) =>{
    const tempDir = 'uploads';
    fs.mkdirSync(tempDir, { recursive: true });
    cb(null, tempDir);
  },
  filename: (req, file, cb) => {
    const uniqueFileName =  file.filename  + path.extname(file.originalname);
    cb(null, uniqueFileName);
  },
});

const upload = multer({ storage });

app.post("/api/change",upload.single('file'), async (req, res) => {
  try {
   
    const excelBuffer = req.file.path; 
    console.log(excelBuffer);  
    const workbook = XlsxPopulate.fromFileAsync(excelBuffer);
   const name = excelBuffer.split('/')
   console.log(name)
    workbook.then((book) => {
      book.toFileAsync('here.xlsx', { password: "123" });
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=protected_excel.xlsx"
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      const file = path.join(__dirname, "here.xlsx");
      res.status(200).sendFile(file);
    });
  } catch (error) {
    console.log(error)
    res.status(500).send(error.message);
  }
});

app.listen(3000, (req, res) => {
  console.log("listening on port 3000");
});

// ------------------------------------main logic ---------------------------------------------
// const workbook = await XlsxPopulate.fromBlankAsync("./here.xlsx");

// // workbook
// //   .then((workbook) => {
// //     console.log(workbook);

//   await    workbook.toFileAsync("protectFiles.xlsx", { password: "qwer" });
//     return workbook;
//   // })
//   // .catch((err) => {
//   //   console.log(err);
//   // });
