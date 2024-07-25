// server.js
const express = require("express");
const mongoose = require("mongoose");
const multer = require("multer");
const nodemailer = require("nodemailer");
const ExcelModel = require("./models/excelModel.cjs");

const excelService = require("./services/excelService.cjs");
const dotenv = require("dotenv");

dotenv.config();

const app = express();
const PORT = 3000;

const DATABASE_NAME = "Hello";
const MONGODB_URI = `mongodb://localhost:27017/{DATABASE_NAME}`;

mongoose.connect(MONGODB_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const db = mongoose.connection;

db.on("error", (err) => {
  console.error(`MongoDB connection error: ${err}`);
});

db.once("open", () => {
  console.log(`Connected to MongoDB: ${DATABASE_NAME}`);
});

app.use(express.json());
app.set("view engine", "ejs");
app.set("views", __dirname + "/views");

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.get("/", (req, res) => {
  res.render("upload", { message: null });
});

app.post("/upload", upload.single("excelFile"), async (req, res) => {
  try {
    await excelService.uploadExcelData(req.file.buffer);
    res.redirect("/email-page");
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

app.get("/email-page", (req, res) => {
  res.render("email-page");
});

app.post("/send-emails", async (req, res) => {
  try {
    await excelService.sendEmails();
    res.send("Emails sent successfully!");
  } catch (error) {
    console.error(error);
    res.status(500).send("Internal Server Error");
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
