/* ===============================
    IMPORTS & INITIAL SETUP
   =============================== */

const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const puppeteer = require("puppeteer");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const ExcelJS = require("exceljs");

const app = express();
const PORT = 3000;

/* ===============================
   VIEW ENGINE & STATIC FILES
   =============================== */

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use("/static", express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));

/* ===============================
   FILE UPLOAD HANDLING (MULTER)
   =============================== */

const upload = multer({ dest: "tmp/" });

const MAX_ATT_SECTIONS = 20;
let attendanceFields = [];

for (let i = 0; i < MAX_ATT_SECTIONS; i++) {
  attendanceFields.push({
    name: `attendanceFiles${i}[]`,
    maxCount: 50
  });
}

const cpUpload = upload.fields([
  { name: "invitation[]", maxCount: 50 },
  { name: "poster[]", maxCount: 50 },
  { name: "resource[]", maxCount: 10 },
  { name: "photos[]", maxCount: 50 },
  { name: "feedback[]", maxCount: 50 },
  ...attendanceFields
]);

/* ===============================
   HELPERS
   =============================== */

// Convert uploaded files to base64
function convertToBase64(files) {
  if (!files || files.length === 0) return [];
  return files.map(f => {
    const data = fs.readFileSync(f.path);
    return `data:${f.mimetype};base64,${data.toString("base64")}`;
  });
}

/* 
  Build all report data from the request
  Used by PDF, Preview, DOCX, Excel — single source.
*/
function buildTemplateData(req) {
  const logo1 = fs.readFileSync(path.join(__dirname, "public/logo1.png")).toString("base64");
  const logo2 = fs.readFileSync(path.join(__dirname, "public/logo2.png")).toString("base64");

  const data = {
    activityName: req.body.activityName || "",
    coordinator: req.body.coordinator || "",
    activityDate: req.body.activityDate || "",
    duration: req.body.duration || "",
    po: req.body.po || "",
    programLine: req.body.programLine || "",
    resourceText: req.body.resourceText || "",
    sessionName: req.body.sessionName || "",
    sessionResourcePerson: req.body.sessionResourcePerson || "",
    sessionCoordinators: req.body.sessionCoordinators || "",
    sessionStartDate: req.body.sessionStartDate || "",
    sessionStartTime: req.body.sessionStartTime || "",
    sessionEndDate: req.body.sessionEndDate || "",
    sessionEndTime: req.body.sessionEndTime || "",
    sessionParticipants: req.body.sessionParticipants || "",
    sessionActivityTitle: req.body.sessionActivityTitle || "",
    sessionPreamble: req.body.sessionPreamble || "",
    sessionSummary: req.body.sessionSummary || "",
    academicYear: "2024-25",
    headerLeft: `data:image/png;base64,${logo1}`,
    headerRight: `data:image/png;base64,${logo2}`,
  };

  // Table of contents
  let toc = req.body["tocRows[]"] || req.body["tocRows"] || [];
  if (!Array.isArray(toc)) toc = [toc];
  toc = toc.filter(Boolean);

  // Attendance sections
  let titles = req.body["attendanceTitles[]"] || req.body["attendanceTitles"] || [];
  if (!Array.isArray(titles)) titles = [titles];

  const attendanceSections = titles.map((title, i) => ({
    title,
    images: convertToBase64(req.files[`attendanceFiles${i}[]`] || []),
  }));

  return {
    data,
    tocRows: toc,
    invitationUrls: convertToBase64(req.files["invitation[]"] || []),
    posterUrls: convertToBase64(req.files["poster[]"] || []),
    resourceUrls: convertToBase64(req.files["resource[]"] || []),
    attendanceSections,
    photosUrls: convertToBase64(req.files["photos[]"] || []),
    feedbackUrls: convertToBase64(req.files["feedback[]"] || []),
  };
}

/* ===============================
   ROUTES — UI FORM
   =============================== */

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "views", "form.html"));
});
/* ===============================
   ROUTES — PDF PREVIEW
   =============================== */

app.post("/preview", cpUpload, async (req, res) => {
  try {
    const templateData = buildTemplateData(req);

    const html = await new Promise((resolve, reject) => {
      app.render("report", templateData, (err, out) => {
        if (err) reject(err);
        else resolve(out);
      });
    });

    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox"]
    });

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: { top: 0, bottom: 0, left: 0, right: 0 }
    });

    await browser.close();
    res.setHeader("Content-Type", "application/pdf");
    res.send(pdf);

  } catch (err) {
    res.status(500).send("Preview error: " + err.message);
  }
});

/* ===============================
   ROUTES — PDF DOWNLOAD
   =============================== */

app.post("/generate", cpUpload, async (req, res) => {
  try {
    const templateData = buildTemplateData(req);

    const html = await new Promise((resolve, reject) => {
      app.render("report", templateData, (err, out) => {
        if (err) reject(err);
        else resolve(out);
      });
    });

    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox"]
    });

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: { top: 0, bottom: 0, left: 0, right: 0 }
    });

    await browser.close();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "attachment; filename=report.pdf");
    res.send(pdf);

  } catch (err) {
    res.status(500).send("PDF error: " + err.message);
  }
});

/* ===============================
   ROUTES — DOCX DOWNLOAD
   =============================== */

app.post("/generate-docx", cpUpload, async (req, res) => {
  try {
    const form = req.body;
    const rawToc = form["tocRows[]"] || form["tocRows"] || [];
    const tocRows = Array.isArray(rawToc) ? rawToc : [rawToc];

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [ new TextRun({ text: "ACTIVITY CONDUCTED REPORT", bold: true, size: 28 }) ],
              spacing: { after: 200 }
            }),

            new Paragraph(`Activity Name: ${form.activityName || ""}`),
            new Paragraph(`Co-ordinator: ${form.coordinator || ""}`),
            new Paragraph(`Date: ${form.activityDate || ""}`),
            new Paragraph(`Duration: ${form.duration || ""}`),
            new Paragraph(`PO & POs: ${form.po || ""}`),
            new Paragraph(`Program Line: ${form.programLine || ""}`),
            new Paragraph(""),

            new Paragraph({ children: [ new TextRun({ text: "TABLE OF CONTENTS", bold: true }) ] }),
            ...tocRows.map((r, i) => new Paragraph(`${i + 1}. ${r}`)),

            new Paragraph(""),
            new Paragraph({ children: [ new TextRun({ text: "SESSION REPORT", bold: true }) ] }),

            new Paragraph(`Session Name: ${form.sessionName || ""}`),
            new Paragraph(`Resource Person: ${form.sessionResourcePerson || ""}`),
            new Paragraph(`Co-ordinators: ${form.sessionCoordinators || ""}`),
            new Paragraph(`Start Date: ${form.sessionStartDate || ""} Time: ${form.sessionStartTime || ""}`),
            new Paragraph(`End Date: ${form.sessionEndDate || ""} Time: ${form.sessionEndTime || ""}`),
            new Paragraph(`Participants: ${form.sessionParticipants || ""}`),
            new Paragraph(`Activity Title: ${form.sessionActivityTitle || ""}`),
            new Paragraph(`Preamble: ${form.sessionPreamble || ""}`),
            new Paragraph("Summary:"),
            new Paragraph((form.sessionSummary || "").substring(0, 5000)),
          ]
        }
      ]
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Disposition", "attachment; filename=report.docx");
    res.send(buffer);

  } catch (err) {
    res.status(500).send("DOCX error: " + err.message);
  }
});

/* ===============================
   ROUTES — EXCEL DOWNLOAD (Matches PDF)
   =============================== */

app.post("/generate-excel", cpUpload, async (req, res) => {
  try {
    const {
      data, tocRows,
      invitationUrls, posterUrls, resourceUrls,
      attendanceSections, photosUrls, feedbackUrls
    } = buildTemplateData(req);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Report");

    // Helper to insert image
    function insertImage(url) {
      const [meta, base64] = url.split(",");
      const buf = Buffer.from(base64, "base64");
      const ext = meta.includes("jpeg") ? "jpg" : "png";

      const id = workbook.addImage({ buffer: buf, extension: ext });

      const startRow = sheet.lastRow.number + 1;
      sheet.addRow(["(Image Below)"]);
      sheet.addImage(id, {
        tl: { col: 1, row: startRow },
        br: { col: 6, row: startRow + 15 }
      });
      sheet.addRow([]);
    }

    /* PAGE 1 */
    sheet.addRow(["ACTIVITY CONDUCTED REPORT"]).font = { bold: true, size: 16 };
    sheet.addRow([]);

    [
      ["Activity Name", data.activityName],
      ["Co-ordinator", data.coordinator],
      ["Date", data.activityDate],
      ["Duration", data.duration],
      ["PO & POs", data.po],
      ["Program Line", data.programLine],
      ["Academic Year", data.academicYear]
    ].forEach(r => sheet.addRow(r));

    sheet.addRow([]);

    /* TOC */
    sheet.addRow(["TABLE OF CONTENTS"]).font = { bold: true, size: 14 };
    sheet.addRow(["Sl. No", "Content"]).font = { bold: true };

    tocRows.forEach((item, i) => sheet.addRow([i + 1, item]));
    sheet.addRow([]);

    /* INVITATION */
    if (invitationUrls.length) {
      sheet.addRow(["INVITATION"]).font = { bold: true, size: 14 };
      invitationUrls.forEach(insertImage);
      sheet.addRow([]);
    }

    /* POSTER */
    if (posterUrls.length) {
      sheet.addRow(["POSTER"]).font = { bold: true, size: 14 };
      posterUrls.forEach(insertImage);
      sheet.addRow([]);
    }

    /* RESOURCE PERSON */
    sheet.addRow(["RESOURCE PERSON DETAILS"]).font = { bold: true, size: 14 };
    if (resourceUrls.length) insertImage(resourceUrls[0]);
    sheet.addRow(["Description"]);
    sheet.addRow([data.resourceText]);
    sheet.addRow([]);

    /* SESSION REPORT */
    sheet.addRow(["SESSION REPORT"]).font = { bold: true, size: 14 };

    [
      ["Session Name", data.sessionName],
      ["Resource Person", data.sessionResourcePerson],
      ["Co-ordinator(s)", data.sessionCoordinators],
      ["Start Date", data.sessionStartDate],
      ["Start Time", data.sessionStartTime],
      ["End Date", data.sessionEndDate],
      ["End Time", data.sessionEndTime],
      ["Participants", data.sessionParticipants],
      ["Activity Title", data.sessionActivityTitle],
      ["Preamble", data.sessionPreamble],
      ["Summary", data.sessionSummary],
    ].forEach(r => sheet.addRow(r));

    sheet.addRow([]);

    /* ATTENDANCE */
    if (attendanceSections.length) {
      sheet.addRow(["ATTENDANCE"]).font = { bold: true, size: 14 };
      attendanceSections.forEach(sec => {
        sheet.addRow([sec.title]).font = { bold: true };
        sec.images.forEach(insertImage);
      });
      sheet.addRow([]);
    }

    /* PHOTOS */
    if (photosUrls.length) {
      sheet.addRow(["PHOTOS"]).font = { bold: true, size: 14 };
      photosUrls.forEach(insertImage);
      sheet.addRow([]);
    }

    /* FEEDBACK */
    if (feedbackUrls.length) {
      sheet.addRow(["FEEDBACK"]).font = { bold: true, size: 14 };
      feedbackUrls.forEach(insertImage);
      sheet.addRow([]);
    }

    // Auto width
    sheet.columns.forEach(col => col.width = 40);

    // Return Excel
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Disposition", "attachment; filename=report.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);

  } catch (err) {
    res.status(500).send("Excel error: " + err.message);
  }
});

/* ===============================
   START SERVER
   =============================== */

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
