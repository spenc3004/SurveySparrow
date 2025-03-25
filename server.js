const express = require('express')
const fs = require('fs')
const path = require('path')
const { exec } = require('child_process');
const nodemailer = require('nodemailer')

const app = express()
require('dotenv').config()

const PORT = process.env.PORT || 3001

app.use(express.json())

const OUTPUT_DIR = path.join(__dirname, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

app.get("/", (req, res) => {
    res.send('Hello from server.js')
})


app.post("/ss", async (req, res) => {
    // #region Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const htmlContent = generateHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'output.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'output.docx');

        fs.writeFileSync(htmlFilePath, htmlContent);
        await convertToDocx(htmlFilePath, docxFilePath);
        await sendEmail(docxFilePath);

        fs.unlinkSync(htmlFilePath);
        fs.unlinkSync(docxFilePath);

        res.status(200).send('File processed and email sent.');
    }

    catch (error) {
        console.error(error);
        res.status(500).send('Internal Server Error');
    }
    // #endregion
});

function generateHTML(data) {
    // #region insert data dynamically to html
    return `<!DOCTYPE html>
    <html lang="en">
    
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>HVAC Brief</title>
    </head>
    
    <body>
            <p>We have a new HVAC client doing a "Insert Product and Size"</p>
            <p><strong><u>DESIGN:</u></strong> Specific design? ${data.design} Design based on look of website? ${data.designBasedOnWeb}</p>
            <p><strong><u>LOCATION INFORMATION:</u></strong><br>
            Address: <br>
            ${data.companyName}<br>
            ${data.companyAddress}<br>
            ${data.companyAddress2}<br>
            ${data.companyCity}, ${data.companyState} ${data.companyZip}<br>
            “Insert Call to Action” Based on Q4<br>
            Call Today to Schedule Your Appointment! 
            OR Call Today or Conveniently Schedule Online! 
            OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment! (Q4)
            Phone: ${data.companyPhone}<br>
            Website: ${data.website}<br>
            Hours: <br>
            License #: ${data.license}
            </p>
            <p><strong><u>PHOTOS TO USE:</u></strong> ${data.photos}</p>
            <p><strong><u>SERVICES:</u></strong> ${data.services}</p>
            <p><strong><u>You Can Trust Us to Do the Job for You:</u></strong> ${data.pricing}, ${data.warranties}, ${data.technicians}, ${data.financing}</p>
            <p><strong><u>TAGLINES:</u></strong> ${data.hasTaglines}<br>
            (Side #1)- ${data.taglines.tagline1}<br>
            (Side #2)- ${data.taglines.tagline2}
            </p>
            <p><strong><u>RATINGS:</u></strong><br>
            Google: ${data.stars.google}<br>
            ${data.stars.other1}<br>
            ${data.stars.other2}<br>
            ${data.stars.other3}<br>
            ${data.stars.other4}<br>
            <p><strong><u>LOGOs to Use:</u></strong> Logo: ${data.logo}<br>
            BBB: ${data.bbb}<br>
            Logo or Award: ${data.awards}<br>
            Other Awards, Affiliations, or Organizations: ${data.otherAwards}
            </p>
            <p><strong><u>OTHER NOTES:</u></strong> <br>
            Applicable: ${data.applicables}<br>
            Radius Offers: <br>
            <table>
                <tr>
                    <th>Coupon</th>
                    <th>Disclaimer</th>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon1}</td>
                    <td>${data.radiusOffers.disclaimer1}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon2}</td>
                    <td>${data.radiusOffers.disclaimer2}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon3}</td>
                    <td>${data.radiusOffers.disclaimer3}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon4}</td>
                    <td>${data.radiusOffers.disclaimer4}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon5}</td>
                    <td>${data.radiusOffers.disclaimer5}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon6}</td>
                    <td>${data.radiusOffers.disclaimer6}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon7}</td>
                    <td>${data.radiusOffers.disclaimer7}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon8}</td>
                    <td>${data.radiusOffers.disclaimer8}</td>
                </tr>
            </table>
            ${JSON.parse(data.homeOwner) ? `
               New Homeowner Mailings:</p>
            <p><strong><u>COUPONS(There are "4" coupons):</u></strong><br>           <table>
                <tr>
                    <th>Coupon</th>
                    <th>Disclaimer</th>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon1}</td>
                    <td>${data.radiusOffers.disclaimer1}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon2}</td>
                    <td>${data.radiusOffers.disclaimer2}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon3}</td>
                    <td>${data.radiusOffers.disclaimer3}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon4}</td>
                    <td>${data.radiusOffers.disclaimer4}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon5}</td>
                    <td>${data.radiusOffers.disclaimer5}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon6}</td>
                    <td>${data.radiusOffers.disclaimer6}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon7}</td>
                    <td>${data.radiusOffers.disclaimer7}</td>
                </tr>
                <tr>
                    <td>${data.radiusOffers.coupon8}</td>
                    <td>${data.radiusOffers.disclaimer8}</td>
                </tr>
            </table></p> 
                
            ` : ``}
        </body>
    </body>
    
    </html>`
    // #endregion 
}

`            `
function convertToDocx(htmlFilePath, docxFilePath) {
    // #region Convert output.html to a docx
    return new Promise((resolve, reject) => {
        exec(`pandoc ${htmlFilePath} -o ${docxFilePath}`, (error, stdout, stderr) => {
            if (error) {
                console.error(`Error converting HTML to DOCX: ${stderr}`);
                reject(error);
            } else {
                resolve();
            }
        });
    });
    // #endregion
}

async function sendEmail(attachmentPath) {
    // #region Send output.docx in email

    let transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: { user: process.env.user, pass: process.env.pass },
        tls: {
            ciphers: "SSLv3" // Helps avoid connection issues
        }

    });

    let mailOptions = {
        from: process.env.user,
        to: process.env.recipient,
        subject: 'New HVAC Survey Submitted',
        text: 'Please see the attached document.',
        attachments: [{ path: attachmentPath }]
    };

    await transporter.sendMail(mailOptions);
    // #endregion
}


app.listen(PORT, () => {
    console.log(`This is running on port http://localhost:${PORT}`)
});