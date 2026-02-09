const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { BlobServiceClient, BlobSASPermissions } = require("@azure/storage-blob");

module.exports = async function (context, req) {
    try {
        const data = req.body;
        const templateFile = req.body?.templateFile;

        if (!data || !templateFile) {
            context.res = {
                status: 400,
                body: { success: false, message: "Body and templateFile are required" }
            };
            return;
        }

        // Load template
        const templatePath = path.resolve(__dirname, templateFile);
        if (!fs.existsSync(templatePath)) {
            throw new Error(`Template not found: ${templateFile}`);
        }

        const content = fs.readFileSync(templatePath, "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: () => ""
        });

        // Render document
        doc.render(data);

        const buffer = doc.getZip().generate({
            type: "nodebuffer",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });

        // Upload to Azure Blob
        const blobServiceClient = BlobServiceClient.fromConnectionString(
            process.env.AzureWebJobsStorage
        );

        const containerClient = blobServiceClient.getContainerClient("generated-documents");
        await containerClient.createIfNotExists();

        const fileName = `${(data.Deal_Name || "Document").replace(/\s+/g, "_")}_${Date.now()}.docx`;
        const blockBlobClient = containerClient.getBlockBlobClient(fileName);

        await blockBlobClient.upload(buffer, buffer.length);

        // Generate SAS URL (1 hour)
        const sasUrl = await blockBlobClient.generateSasUrl({
            permissions: BlobSASPermissions.parse("r"),
            expiresOn: new Date(Date.now() + 60 * 60 * 1000),
        });

        context.res = {
            status: 200,
            body: {
                success: true,
                fileName,
                downloadUrl: sasUrl
            }
        };

    } catch (error) {
        context.res = {
            status: 500,
            body: {
                success: false,
                error: error.message
            }
        };
    }
};
