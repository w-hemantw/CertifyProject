// services/excelService.js
const xlsx = require('xlsx');
const fs = require('fs');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const ExcelModel = require('../models/excelModel.cjs');

async function uploadExcelData(buffer) {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    const certificates = data.map((row) => ({
        name: row.name,
        email: row.email,
        mobile: row.mobile,
        amount: row.amount,
        numberoftree: row.numberoftree,
    }));

    const existingData = await ExcelModel.find({ email: { $in: certificates.map(cert => cert.email) } });
    const newData = certificates.filter(cert => !existingData.some(existingCert => existingCert.email === cert.email));

    if (newData.length > 0) {
        await ExcelModel.create(newData);
        console.log("Uploaded");
    } else {
        console.log("No new data to upload");
    }
}

async function sendEmails() {
    const mongoData = await ExcelModel.find();

    // Set up nodemailer transporter
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASSWORD,
        },
    });

    // Iterate over MongoDB data and send emails
    for (const userData of mongoData) {
        const pdfDoc = new PDFDocument({ margin: 10, size: [1273, 771] });

        // Add PDF content and styling
        const backgroundImagePath = './public/custom.jpg';
        pdfDoc.image(backgroundImagePath, 0, 0, { width: 1273, height: 771 });
        
        // Add name to the PDF
        pdfDoc.font('./public/segoesc.ttf').fontSize(50);
        const nameX = 600; // Adjust this value to position the name correctly
        const nameY = 300; // Adjust this value to position the name correctly
        pdfDoc.text(`${userData.name}`, nameX, nameY);

        // Add number of trees to the PDF
        pdfDoc.fontSize(20);
        const treesX = 658; // Adjust this value to position the number of trees correctly
        const treesY = 380; // Adjust this value to position the number of trees correctly
        pdfDoc.text(` ${userData.numberoftree}`, treesX, treesY);
        
        pdfDoc.fontSize(12).font('Helvetica');

        // Generate PDF file
        const pdfFilePath = `./public/certificates_${userData.name}.pdf`;
        pdfDoc.pipe(fs.createWriteStream(pdfFilePath));
        pdfDoc.end();

        // Convert PDF to buffer
        const pdfBuffer = await new Promise((resolve) => {
            const chunks = [];
            pdfDoc.on('data', (chunk) => chunks.push(chunk));
            pdfDoc.on('end', () => resolve(Buffer.concat(chunks)));
        });

        // Set up email options
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: userData.email,
            subject: 'Tree Planting Report',
            text: `Dear ${userData.name},\n\nAttached is your Tree Planting Report.\n\nRegards,\nYour Organization`,
            attachments: [
                {
                    filename: 'Tree_Planting_Report.pdf',
                    content: pdfBuffer,
                },
            ],
        };

        // Send email
        await transporter.sendMail(mailOptions);
    }
}

module.exports = { uploadExcelData, sendEmails };
