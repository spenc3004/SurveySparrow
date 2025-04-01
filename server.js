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
        const company = data.companyName
        const htmlContent = generateHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'HVAC_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'HVAC_Brief.docx');

        fs.writeFileSync(htmlFilePath, htmlContent);
        await convertToDocx(htmlFilePath, docxFilePath);
        await sendEmail(docxFilePath, company);

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
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create a tags for logo photos
    const logoLinks = data.logo
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Logo</i></u></a>`)
        .join("<br>");
    // dynamically create a tags for awar or logo photos
    const awardOrLogoLinks = data.awards
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Logo</i></u></a>`)
        .join("<br>");
    // dynamically create a tags for bbb logo photos
    const bbbLinks = data.bbb
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View BBB Logo</i></u></a>`)
        .join("<br>");

    const phone = `(${data.companyPhone.slice(3, 6)}) ${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    function numCoupons() {
        counter = 0
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        return counter;
    }


    return `<!DOCTYPE html>
    <html lang="en">
    
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    
    <body>
            <h4>We have a new HVAC client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p> ` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p> ` : ``}

            <h5><u>LOCATION INFORMATION:</u></h5>
            ${data.companyName}<br><br>
            ${phone}<br><br>
            ${data.website}<br><br>
            ${data.license}
            ${data.onlineService === "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                OR Call Today or Conveniently Schedule Online! <br>
                OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${photoLinks}
            <h5><u>SERVICES:</u></h5>
            ${data.services.split(",").join("<br>")}
            <h5><u>You Can Trust Us to Do the Job for You:</u></h5>
            ${data.pricing}<br>
            ${data.warranties}<br>
            ${data.technicians}<br>
            ${data.financing}
            <h5><u>TAGLINES:</u></h5>
            ${data.hasTaglines === "true" ? `${data.customTaglines.tagline1}<br>
            ${data.customTaglines.tagline2}<br>` : ``}
            ${data.premadeTaglines !== "null" ? `${data.premadeTaglines.split(",").join("<br>")}` : ``}
            <h5><u>RATINGS:</u></h5>
            ${data.stars.google !== "null" ? `Google: ${data.stars.google}<br>` : ``}
            ${data.stars.other1 !== "null" ? `${data.stars.other1}<br>` : ``}
            ${data.stars.other2 !== "null" ? `${data.stars.other2}<br>` : ``}
            ${data.stars.other3 !== "null" ? `${data.stars.other3}<br>` : ``}
            ${data.stars.other4 !== "null" ? `${data.stars.other4}` : ``}
            <h5><u>LOGOs to Use:</u></h5> 
            ${!data.logo ? `` : `${logoLinks}<br>`}
            ${!data.bbb ? `` : `${bbbLinks}<br>`}
            ${!data.awards ? `` : `${awardOrLogoLinks}<br>`}
            ${data.otherAwards !== "null" ? `<h6>Other Awards, Affiliations, or Organizations:</h6>  
            ${data.otherAwards.split(",").join("<br>")}` : ``}
            <h5><u>OTHER NOTES:</u></h5> 
            ${data.applicables}
           
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Home Owner mailings?</h6> Yes
                <h6>Radius Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.radiusOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>${data.radiusOffers.coupon3}</td>
                            <td>${data.radiusOffers.disclaimer3}</td>
                        </tr>`: ``}
                </table></p> 
              
            ` : ``}
            <h5><u>COUPONS(There are ${numCoupons()} coupons):</u></h5>        
             <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.radiusOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3}</td>
                    </tr>`: ``}
                </table></p> 
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

async function sendEmail(attachmentPath, company) {
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
        subject: `New HVAC Survey Submitted for ${company}`,
        text: 'Please see the attached document.',
        attachments: [{ path: attachmentPath }]
    };

    await transporter.sendMail(mailOptions);
    // #endregion
}


app.listen(PORT, () => {
    console.log(`This is running on port http://localhost:${PORT}`)
});