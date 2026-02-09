# DOCX Generator using Docxtemplater (Node.js)

This project generates dynamic **.docx** documents using **Docxtemplater** and **PizZip** in Node.js.  
It is designed for **long legal documents** where preserving the original **DOCX formatting** (tables, headers, styles, pagination, etc.) is critical.

The input data is injected into a Word template (`.docx`) and the final rendered document is saved as a new `.docx` file.

---

## ✨ Features

- Uses **Docxtemplater** for advanced templating logic
- Preserves original **DOCX formatting**
- Supports:
  - Variables
  - Conditional logic
  - Loops (arrays)
  - Nested objects
- Ideal for **legal documents, contracts, agreements**
- Simple Node.js setup

---

## 📂 Project Structure

```text
.
├── sample_doc.docx      # Input Word template
├── op.docx              # Generated output file
├── server.js             # Main Node.js script
├── package.json
└── README.md
```

---

## 🧩 Technologies Used

* Node.js
* Docxtemplater
* PizZip
* fs (File System)
* path

---

## 📦 Prerequisites

Make sure you have **Node.js** installed.

👉 Download from: [https://nodejs.org/](https://nodejs.org/)

---

## ⚙️ Installation

1. Clone the repository:

```bash
git clone https://github.com/DINAKAR-S/Docxtemplater.git
cd Docxtemplater
```

2. Install dependencies:

```bash
npm install docxtemplater pizzip
```

---

## ▶️ Usage

1. Place your Word template file in the project directory
   (default name: `sample_doc.docx`)

2. Update the `data` object inside `server.js` with your required values.

3. Run the script:

```bash
node ./server.js
```

4. The generated document will be created as:

```text
op.docx
```

---

## 📝 Docxtemplater Tag Types Reference

Docxtemplater uses various types of tags enclosed in curly braces `{}` to manage template data. Here are the main tag types:

### 🔹 Placeholder Tags
Simple variable replacement:
```docx
Hello {name}!
```
Data: `{"name": "John"}` → Output: `Hello John!`

### 🔹 Loop Tags
Iterate over arrays:
```docx
{#products}
{name}, {price} €
{/products}
```
Data: `{"products": [{"name": "Windows", "price": 100}]}` → Creates multiple rows

### 🔹 Condition Tags
Show/hide content based on boolean values:
```docx
{#hasKitty}Cat’s name: {kitty}{/hasKitty}
{#hasDog}Dog’s name: {dog}{/hasDog}
```

### 🔹 Sections
Sections handle both loops and conditions based on data type:

| Type of Value | Section Behavior |
|---------------|------------------|
| falsy or empty array | Never shown |
| non-empty array | Each element |
| object | Once, with object as scope |
| truthy value | Once, unchanged scope |

### 🔹 Inverted Sections
Opposite of regular sections (render when value is false/empty):
```docx
{^repo}No repos :({/repo}
```

### 🔹 Table Rows with Loops
Create tables with dynamic rows:
```docx
| Name | Age | Phone Number |
{#users}{name}|{age}|{phone}{/}
```

### 🔹 Raw XML Tags
Insert raw XML content:
```docx
{@rawXml}
```
Advanced feature for complex formatting.

---

## 📝 Example Data Structure

The script supports complex nested data like:

* Borrowers
* Guarantors
* Trustees
* Conditional sections
* Repeating sections

Example (simplified):

```js
{
  Reference: "Test",
  Borrower: "ABCDEFGH Holdings Pty Ltd",
  Guarantors: [
    {
      Guarantors_Description: "John Smith",
      Emails: ["john.smith@example.com"]
    }
  ],
  is_trustee_present: true
}
```

---

## 🔧 Configuration Options

### Paragraph Loop Option
To avoid unwanted line breaks in loops:

```js
const doc = new Docxtemplater(zip, {
  paragraphLoop: true  // Removes outer section newlines
});
```

### Custom Delimiters
Change delimiter characters if needed:

```js
const doc = new Docxtemplater(zip, {
  syntax: {
    changeDelimiterPrefix: "$"  // Use $ instead of {= =} syntax
  }
});
```

---

## 🧠 Notes

* All document logic (loops, conditions, placeholders) is handled **inside the DOCX template** using Docxtemplater syntax.
* This approach ensures **legal document formatting is not broken**.
* `nullGetter` is used to avoid template errors when values are missing.

---

## 🚀 Use Cases

* Legal agreements
* Loan documents
* Contracts
* Offer letters
* Compliance documents
* Any long-form DOCX generation

---

## 📄 License

MIT License
Feel free to use, modify, and distribute.

---

## 🤝 Contributions

Pull requests are welcome.
If you find a bug or want to improve the template logic, feel free to contribute.

---

## 📬 Questions?

Open an issue or reach out if you need help with Docxtemplater logic or complex DOCX templates.

## 🔗 Additional Resources

- [Docxtemplater Official Documentation](https://docxtemplater.com/docs/tag-types/)
- [Docxtemplater Demo](https://docxtemplater.com/demo)
