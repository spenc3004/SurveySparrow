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

const bbbImageMap = {
    "Horizontal black and white": "https://www.themailshark.com/prepress/surveysparrow/hs_bbb/Horizontal%20black%20and%20white.ai",
    "Vertical black and white": "https://www.themailshark.com/prepress/surveysparrow/hs_bbb/Vertical%20black%20and%20white.ai",
    "Horizontal blue and white": "https://www.themailshark.com/prepress/surveysparrow/hs_bbb/Horizontal%20blue%20and%20white.ai",
    "Vertical blue and white with blue A": "https://www.themailshark.com/prepress/surveysparrow/hs_bbb/Vertical%20blue%20and%20white%20with%20blue%20A.ai.ai",
    "Vertical blue and white with red A": "https://www.themailshark.com/prepress/surveysparrow/hs_bbb/Vertical%20blue%20and%20white%20with%20red%20A.ai",
    "Vertical blue and white no A": "https://www.themailshark.com/prepress/surveysparrow/hs_bbb/Vertical%20blue%20and%20white%20no%20A.ai"
}

const aaaImageMap = {
    "Red and blue": "https://www.themailshark.com/prepress/surveysparrow/auto/aaa/Red%20and%20blue.ai",
    "Black and white": "https://www.themailshark.com/prepress/surveysparrow/auto/aaa/Black%20and%20white.ai"
}

const napaImageMap = {
    "Black and yellow with AUTOCARE CENTER": "https://www.themailshark.com/prepress/surveysparrow/auto/napa/Black%20and%20yellow%20with%20AUTOCARE%20CENTER.ai",
    "Blue and yellow with AUTOCARE CENTER": "https://www.themailshark.com/prepress/surveysparrow/auto/napa/Blue%20and%20yellow%20with%20AUTOCARE%20CENTER.ai",
    "Blue and white with yellow border": "https://www.themailshark.com/prepress/surveysparrow/auto/napa/Blue%20and%20white%20with%20yellow%20border.ai",
    "Black and white NAPA": "https://www.themailshark.com/prepress/surveysparrow/auto/napa/Black%20and%20white%20NAPA.ai",
    "Blue and yellow NAPA": "https://www.themailshark.com/prepress/surveysparrow/auto/napa/Blue%20and%20yellow%20NAPA.ai"

}

const caOrTxImageMap = {
    "CA STAR Certified": "https://www.themailshark.com/prepress/surveysparrow/auto/states/CA%20Star%20Certified.ai",
    "TX Offical Vehicle Inspection": "https://www.themailshark.com/prepress/surveysparrow/auto/states/TX%20Official%20Vehical%20Inspection.jpg"
}

app.get("/", (req, res) => {
    res.send('Hello from server.js')
})


app.post("/hvac", async (req, res) => {
    // #region Receive HVAC JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "HVAC"
        const htmlContent = generateHvacHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'HVAC_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'HVAC_Brief.docx');

        fs.writeFileSync(htmlFilePath, htmlContent);
        await convertToDocx(htmlFilePath, docxFilePath);
        await sendEmail(docxFilePath, company, type);

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

function generateHvacHTML(data) {
    // #region insert data dynamically into html
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for logo photos
    const logoLinks = data.logo
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Logo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for award or logo photos
    const awardOrLogoLinks = data.awards
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Logo</i></u></a>`)
        .join("<br>");


    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
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
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService === "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``}
            <h5><u>SERVICES:</u></h5>
            ${data.services.split(",").join("<br>")}
            <h5><u>You Can Trust Us to Do the Job for You:</u></h5>
            <p>${data.pricing}</p>
            <p>${data.warranties}</p>
            <p>${data.technicians}</p>
            <p>${data.financing}</p>
            <h5><u>TAGLINES:</u></h5>
            ${data.hasTaglines === "true" ?
            `${data.customTaglines.tagline1 !== "null" ? `<p>${data.customTaglines.tagline1}</p>` : ``}
                ${data.customTaglines.tagline2 !== "null" ? `<p>${data.customTaglines.tagline2}</p>` : ``}`
            : ``}
            ${data.premadeTaglines !== "null" ? `${data.premadeTaglines.split(",").join("<br>")}` : ``}
            <h5><u>RATINGS:</u></h5>
            ${data.stars.google !== "null" ? `<p>Google: ${data.stars.google}</p>` : ``}
            ${data.stars.other1 !== "null" ? `<p>${data.stars.other1}</p>` : ``}
            ${data.stars.other2 !== "null" ? `<p>${data.stars.other2}</p>` : ``}
            ${data.stars.other3 !== "null" ? `<p>${data.stars.other3}</p>` : ``}
            <h5><u>LOGOs to Use:</u></h5> 
            ${!data.logo ? `` : `${logoLinks}<br>`}
            ${!data.awards ? `` : `${awardOrLogoLinks}<br>`}
            ${bbbImageMap[data.bbb] ? `<a href="${bbbImageMap[data.bbb]}" target="_blank"><u><i>View BBB Logo</i></u></a><br>` : ``}
            ${data.otherAwards !== "null" ? `<h6>Other Awards, Affiliations, or Organizations:</h6>  
            ${data.otherAwards.split(",").join("<br>")}` : ``}
            <h5><u>OTHER NOTES:</u></h5> 
            <p>${data.applicables}</p>
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

app.post("/auto", async (req, res) => {
    // #region AUTO Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "Auto"
        const htmlContent = generateAutoHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'Auto_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'Auto_Brief.docx');

        fs.writeFileSync(htmlFilePath, htmlContent);
        await convertToDocx(htmlFilePath, docxFilePath);
        await sendEmail(docxFilePath, company, type);

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

function generateAutoHTML(data) {
    // #region insert data dynamically into html
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for logo photos
    const logoLinks = data.logo
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Logo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for award or logo photos
    const awardOrLogoLinks = data.awards
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Award Logo</i></u></a>`)
        .join("<br>");

    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.coupons) {
            if (key.startsWith("coupon")) {
                if (data.coupons[key] !== "null") {
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
            <p>${data.companyName}</p>
            <p>${data.companyAddress}</p>
            ${data.companyAddress2 !== "null" ? `<p>${data.companyAddress2}</p>` : ``}
            <p>${data.companyCity}, ${data.companyState} ${data.companyZip}</p>
            ${data.onlineService === "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online! </p>
                OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <p>${phone}</p>
            <p>${data.website}</p>
            <p>Hours:</p>
            ${data.hours.Monday !== "null" ? `<p>Monday: ${data.hours.Monday}</p>` : ``}
            ${data.hours.Tuesday !== "null" ? `<p>Tuesday: ${data.hours.Tuesday}</p>` : ``}
            ${data.hours.Wednesday !== "null" ? `<p>Wednesday: ${data.hours.Wednesday}</p>` : ``}
            ${data.hours.Thursday !== "null" ? `<p>Thrusday: ${data.hours.Thursday}</p>` : ``}
            ${data.hours.Friday !== "null" ? `<p>Friday: ${data.hours.Friday}</p>` : ``}
            ${data.hours.Saturday !== "null" ? `<p>Saturday: ${data.hours.Saturday}</p>` : ``}
            ${data.hours.Sunday !== "null" ? `<p>Sunday: ${data.hours.Sunday}</p>` : ``}
            <h5><u>PHOTOS TO USE:</u></h5>
            <p>${data.vehicles}</p>
            ${photoLinks}
            <h5><u>You Can Trust Us to Do the Job for You:</u></h5>
            <p>${data.aseTechnicians}</p>
            <p>${data.warranties}</p>
            <p>${data.shuttleLoanerService}</p>
            <p>${data.financing}</p>
            <p>${data.amenities}</p>
            <p>${data.sameDayService}</p>
            ${data.approveFirst === "true" ? `<p>All Repairs Approved by the Customer Prior to Work Being Done</p>` : ``}
            <h5><u>TAGLINES:</u></h5>
            ${data.tagline1 !== "null" ? `<p>${data.tagline1}</p>` : ``}
            ${data.taglines !== "null" ? `${data.taglines.split(",").join("<br>")}` : ``}
            <h5><u>RATINGS & REVIEWS:</u></h5>
            ${data.stars.google !== "null" ? `<p>Google: ${data.stars.google}</p>` : ``}
            ${data.stars.other1 !== "null" ? `<p>${data.stars.other1}</p>` : ``}
            ${data.stars.other2 !== "null" ? `<p>${data.stars.other2}</p>` : ``}
            ${data.stars.other3 !== "null" ? `<p>${data.stars.other3}</p>` : ``}
            <h5><u>SHOP LOGO:</u></h5> 
            ${!data.logo ? `` : `${logoLinks}`}
            <h5><u>AWARDS LOGOS:</u></h5>
            ${!data.awards ? `` : `${awardOrLogoLinks}<br>`}
            ${bbbImageMap[data.bbb] ? `<a href="${bbbImageMap[data.bbb]}" target="_blank"><u><i>View BBB Logo</i></u></a><br>` : ``}
            ${aaaImageMap[data.aaa] ? `<a href="${aaaImageMap[data.aaa]}" target="_blank"><u><i>View AAA Logo</i></u></a><br>` : ``}
            ${napaImageMap[data.napa] ? `<a href="${napaImageMap[data.napa]}" target="_blank"><u><i>View NAPA Logo</i></u></a><br>` : ``}
            ${caOrTxImageMap[data.caOrTx] ? `<a href="${caOrTxImageMap[data.caOrTx]}" target="_blank"><u><i>View CA or TX Logo</i></u></a><br>` : ``}
            ${data.otherAwards !== "null" ? `<h6>Other Awards, Affiliations, or Organizations:</h6>  
            ${data.otherAwards.split(",").join("<br>")}` : ``}
            <h5><u>OTHER NOTES:</u></h5> 
            
           

            <h5><u>COUPONS(There are ${numCoupons()} coupons):</u></h5>        
             <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.coupons.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.coupons.coupon1}</td>
                        <td>${data.coupons.disclaimer1}</td>
                    </tr>`: ``}
                    ${data.coupons.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.coupons.coupon2}</td>
                        <td>${data.coupons.disclaimer2}</td>
                    </tr>`: ``}
                    ${data.coupons.coupon3 !== "null" ? `
                    <tr>
                        <td>${data.coupons.coupon3}</td>
                        <td>${data.coupons.disclaimer3}</td>
                    </tr>`: ``}
                    ${data.coupons.coupon4 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon4}</td>
                            <td>${data.coupons.disclaimer4}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon5 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon5}</td>
                            <td>${data.coupons.disclaimer5}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon6 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon6}</td>
                            <td>${data.coupons.disclaimer6}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon7 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon7}</td>
                            <td>${data.coupons.disclaimer7}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon8 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon8}</td>
                            <td>${data.coupons.disclaimer8}</td>
                        </tr>`: ``}
                </table></p> 
        </body>
    </body>
    
    </html>`
    // #endregion 
}




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

async function sendEmail(attachmentPath, company, type) {
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
        //bcc: "sharkymailson@gmail.com",
        subject: `New ${type} Survey Submitted for ${company}`,
        text: 'Please see the attached document.',
        attachments: [{ path: attachmentPath }]
    };

    await transporter.sendMail(mailOptions);
    // #endregion
}


app.listen(PORT, () => {
    console.log(`This is running on port http://localhost:${PORT}`)
});