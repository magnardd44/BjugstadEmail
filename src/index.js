const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fsPromises = require("fs").promises;
const nodemailer = require("nodemailer");

const fillHtml = require("./html");

module.exports = async function (context, req) {
  if (req.method === "GET") {
    context.res = {
      status: 200,
      body: __dirname,
    };
  } else {
    if (
      req.body &&
      req.body.navn &&
      req.body.maskinNr &&
      req.body.maskinType &&
      req.body.dato &&
      req.body.userEmail
    ) {
      const navn = req.body.navn;
      const maskinNr = req.body.maskinNr;
      const maskinType = req.body.maskinType;
      const dato = req.body.dato;
      const userEmail = req.body.userEmail;

      let ccEmail = "";

      if (req.body.ccEmail) {
        ccEmail = req.body.ccEmail;
      }

      const transporter = nodemailer.createTransport({
        host: "smtp.office365.com",
        port: 587,
        secure: false,
        auth: {
          user: "",
          pass: "",
        },
      });

      const inputDocxPath = `${context.executionContext.functionDirectory}/mal.docx`;

      const createFile = async () => {
        try {
          // Load the docx file as binary content
          const content = await fsPromises.readFile(inputDocxPath);

          const zip = new PizZip(content);

          const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
          });

          doc.render({
            name: navn,
            maskinNr: maskinNr,
            maskinType: maskinType,
            dato: dato,
          });

          const docxBuffer = doc.getZip().generate({ type: "nodebuffer" });

          return docxBuffer;
        } catch (error) {
          console.error(error);
        }
      };

      const sendEmail = async () => {
        const html = fillHtml(navn, maskinType);

        const docxBuffer = await createFile();

        const msg = {
          to: userEmail,
          cc: [ccEmail, "opplearing@bjugstad.no"],
          from: '"Bjugstad Utleie" <opplearing@bjugstad.no>',
          subject: `Bevis på dokumentert opplæring - ${navn}`,
          html: html,
          attachments: [
            {
              filename: "bevis.docx",
              content: docxBuffer,
              encoding: "base64",
            },
          ],
        };

        transporter.sendMail(msg, (error, info) => {
          if (error) {
            console.error("Error sending email:", error);
          } else {
            console.log("Email sent:", info.response);
          }
        });
      };

      sendEmail().catch((err) => console.error(err));

      context.res = {
        status: 202,
        body: "Suksess! Filen ble sendt.",
      };
    } else {
      context.res = {
        status: 400,
        body: "Feil! Vennligst ha med både navn, maskinNr, maskintype, dato og email.",
      };
    }
  }
};
