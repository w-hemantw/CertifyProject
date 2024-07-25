// models/excelModel.js
const mongoose = require('mongoose');

const excelSchema = new mongoose.Schema({
    name: String,
    email: String,
    mobile: String,
    amount: Number,
    numberoftree: Number,
});

const ExcelModel = mongoose.model('ExcelModel', excelSchema);


module.exports = ExcelModel;
