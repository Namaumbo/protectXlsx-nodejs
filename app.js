const express = require("express");
const XlsxPopulate = require("xlsx-populate");
const fileUpload = require("express-fileupload");
const app = express();
app.use(fileUpload());

// app.post("/api/change", async (req, res) => {
//   try {
//     const excelBuffer = req.body;
//     const workbook = await XlsxPopulate.fromBlankAsync(excelBuffer);
//     workbook
//       .then((workbook) => {
//         workbook.toFileAsync("./example.xlsx", { password: "123456" });
//       })
//       .catch((err) => {
//         console.log(err);
//       });

//     const protectedExcelBuffer = await workbook.outputAsync();

//     res.setHeader(
//       "Content-Disposition",
//       "attachment; filename=protected_excel.xlsx"
//     );
//     res.setHeader(
//       "Content-Type",
//       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
//     );
//     res.send(protectedExcelBuffer);
//   } catch (error) {
//     console.error("Error protecting the Excel file:", error);
//     res.status(500).send("Internal Server Error");
//   }

//   // workbook
//   //   .then((workbook) => {
//   //     res.setHeader(
//   //       "Content-Disposition",
//   //       "attachment; filename=protected_excel.xlsx"
//   //     );
//   //     res.setHeader(
//   //       "Content-Type",
//   //       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
//   //     );
//   //     res.send(workbook.toFileAsync("./example.xlsx", { password: "123456" }));
//   //   })
//   //   .catch((err) => {
//   //     console.log("error processing workbook");
//   //   });
// });

// app.listen(3000, (req, res) => {
//   console.log("listening on port 3000");
// });

// =-----------------main logic --------------------
const workbook = XlsxPopulate.fromBlankAsync("./example.xlsx");

workbook
  .then((workbook) => {
    workbook.toFileAsync("./protectFile.xlsx", { password: "qwer" });
    return workbook;
  })
  .catch((err) => {
    console.log(err);
  });
