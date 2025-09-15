import express, { raw } from "express";
import dotenv from "dotenv";
import cors from "cors";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

dotenv.config();
const app = express();
const PORT = process.env.PORT || 3000;

app.use(
  cors({
    origin: "*",
    // credentials: true,
  })
);
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Microsoft Graph Auth Setup
const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

const graphClient = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await credential.getToken(
        "https://graph.microsoft.com/.default"
      );
      return token.token;
    },
  },
});

app.get("/email-api", (req, res) => {
  res.send("Hello World!");
});

app.post("/send-text-email", async (req, res) => {
  const { email, subject, body } = req.body;
  if (!email || !subject || !body) {
    return res.status(400).json({ error: "Missing required fields" });
  }

  try {
    await graphClient.api(`/users/${process.env.SENDER_EMAIL}/sendMail`).post({
      message: {
        subject,
        body: {
          contentType: "Text",
          content: body,
        },
        toRecipients: [
          {
            emailAddress: {
              address: email,
            },
          },
        ],
      },
      saveToSentItems: "true",
    });

    res.status(200).json({ message: "Email sent successfully!" });
  } catch (error) {
    console.error("Error sending email:", error);
    res.status(500).json({ error: "Failed to send email" });
  }
});

app.post("/send-email", async (req, res) => {
  const { toName, toEmail, subject, textBody, htmlBody } = req.body;
  if (!toName || !toEmail || !subject || !textBody || !htmlBody) {
    return res.status(400).json({
      error: `Missing required fields  
      ${!toName && " To Name,"} ${!toEmail && " To Email,"} 
      ${!subject && " Subject,"} ${!textBody && " Text Body,"} 
      ${!htmlBody && " HTML Body,"}`,
    });
  }
  // const mimeMessage = `
  //   MIME-Version: 1.0
  //   Content-Type: multipart/alternative; boundary="boundary123"
  //   Subject: ${subject}
  //   From: TesseractApps <${process.env.SENDER_EMAIL}>
  //   To: ${toName} <${toEmail}>

  //   --boundary123
  //   Content-Type: text/plain; charset=utf-8

  //   ${textBody}

  //   --boundary123
  //   Content-Type: text/html; charset=utf-8

  //   <!DOCTYPE html>
  //   <html lang="en">
  //     <head>
  //       <meta charset="UTF-8">
  //       <meta name="viewport" content="width=device-width, initial-scale=1.0">
  //       <title>TesseractApps</title>
  //     </head>
  //     <body style="font-family: Roboto, sans-serif;">
  //       ${htmlBody}
  //     </body>
  //   </html>

  //   --boundary123--
  // `;

  // const encodedMessage = Buffer.from(mimeMessage)
  //   .toString("base64")
  //   .replace(/\+/g, "-")
  //   .replace(/\//g, "_")
  //   .replace(/=+$/, "");
  const mail = {
    message: {
      subject: subject,
      body: {
        contentType: "HTML", // "Text" or "HTML"
        content: htmlBody, // use your htmlBody directly
      },
      toRecipients: [
        {
          emailAddress: {
            address: toEmail,
            name: toName,
          },
        },
      ],
      from: {
        emailAddress: {
          address: process.env.SENDER_EMAIL,
        },
      },
    },
    saveToSentItems: true,
  };

  try {
    await graphClient
      .api(`/users/${process.env.SENDER_EMAIL}/sendMail`)
      .post(mail);
    res.status(200).json({ message: "Email sent successfully!" });
  } catch (error) {
    console.error("Error sending email:", error);
    res.status(500).json({ error: "Failed to send email" });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
