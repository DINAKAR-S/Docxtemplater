const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

// ================================
// CONFIGURATION
// ================================
const TEMPLATE_FILE_NAME = "sample_doc.docx";
const OUTPUT_FILE_NAME = "op.docx";


const data = {
  "Reference": "Test",
  "Letter_of_offer_Date": "06 February 2026",
  "Deal_Name": "ABC PTY LTD - Investment Facility",
  "BORROWER": "ABCDEFGH Holdings Pty Ltd",
  "GUARANTOR": "JAMES JOHN",
  "Borrower": "ABCDEFGH Holdings Pty Ltd",
  "Guarantor": "JAMES JOHN",
  "place" : "Chennai",
  "borrower_address": "173 Raja Street, North Carolina Theni 625XXXX",
  "guarantor_address": "6 Hampshire Ave, West Pennant Hills Theni 625XXXX",
  "Guarantors": [
    {
      "ItemNumber": 1,
      "Guarantors_Description": "Sepideh Miraki",
      "Address": "6 Hampshire Ave, West Pennant Hills Theni 625XXXX",
      "Officers": [
        {
          "Attention": "James John"
        }
      ],
      "Emails": ["james.john@example.com"]
    },
    {
      "ItemNumber": 2,
      "Guarantors_Description": "John Smith",     
      "Address": "123 Main Street, Chennai Theni 625XXXX",
      "Officers": [
        {
          "Attention": "John Smith"
        }      ],
      "Emails": ["john.smith@example.com"]
    }
  ], 
  "is_trustee_present": true,
  Trusts: [
      {
        trustee_name: "No 1 Pty LTd",
        custody_agreement: "Not Applicable"
      },
      {
        trustee_name: "ABCDEFGH Trustees Pty Ltd",
        custody_agreement: "Custody Deed dated 5 March 2021"
      }
    ],
     "Borrowers_Executed_By_Section": [
      {
        "Borrower_Company_Name": "ABCDEFGH Holdings Pty Ltd",
        "has_column1": true,
        "electronic_sign":false,
        "has_column2": false,
        "column1_role": "Director",
        "column2_role": "Company Secretary",
      },
      {
        "Borrower_Company_Name": "XYZEDFFFFF Holdings Pty Ltd",
        "has_column1": true,
        "electronic_sign":true,
        "has_column2": true,
        "column1_role": "Director",
        "column2_role": "Company Secretary",
      }
    ],
  }  

// ================================
// LOAD TEMPLATE
// ================================
const templatePath = path.resolve(__dirname, TEMPLATE_FILE_NAME);
const content = fs.readFileSync(templatePath, "binary");

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => "",
});

// ================================
// RENDER DOCUMENT
// ================================
doc.render(data);

// ================================
// SAVE OUTPUT
// ================================
const buffer = doc.getZip().generate({
    type: "nodebuffer",
    mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
});

fs.writeFileSync(path.resolve(__dirname, OUTPUT_FILE_NAME), buffer);

console.log("✔ Document generated:", OUTPUT_FILE_NAME);