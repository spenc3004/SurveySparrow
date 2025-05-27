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
    // dynamically create "a" tags for photos
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for other photos
    const otherPhotoLinks = data.otherPhotos
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;

    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.homeownerOffers) {
            if (key.startsWith("coupon")) {
                if (data.homeownerOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.carrierOffers) {
            if (key.startsWith("coupon")) {
                if (data.carrierOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.retentionOffers) {
            if (key.startsWith("coupon")) {
                if (data.retentionOffers[key] !== "null") {
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
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>

            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <br>
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``} <br>
            ${data.otherPhotos !== "" ? `${otherPhotoLinks}` : ``}
            <br>
            <h5><u>SERVICES:</u></h5>
            ${data.services.split(",").join("<br>")}
            <br>
            <h5><u>You Can Trust Us to Do the Job for You:</u></h5>
            <p>${data.pricing}</p>
            <p>${data.warranties}</p>
            <p>${data.technicians}</p>
            <p>${data.financing}</p>
            <br>
            <h5><u>TAGLINES:</u></h5>
            ${data.hasTaglines === "true" ?
            `${data.customTaglines.tagline1 !== "null" ? `<p>${data.customTaglines.tagline1}</p>` : ``}
                ${data.customTaglines.tagline2 !== "null" ? `<p>${data.customTaglines.tagline2}</p>` : ``}`
            : ``}
            ${data.premadeTaglines !== "null" ? `${data.premadeTaglines.split(",").join("<br>")}` : ``}
            <br>
            <h5><u>RATINGS:</u></h5>
            ${data.stars.google !== "null" ? `<p>Google: ${data.stars.google}</p>` : ``}
            ${data.stars.other1 !== "null" ? `<p>${data.stars.other1}</p>` : ``}
            ${data.stars.other2 !== "null" ? `<p>${data.stars.other2}</p>` : ``}
            ${data.stars.other3 !== "null" ? `<p>${data.stars.other3}</p>` : ``}
            <br>
            <h5><u>LOGOs to Use:</u></h5> 
            ${!data.logo ? `` : `${logoLinks}<br>`}
            ${!data.awards ? `` : `${awardOrLogoLinks}<br>`}
            ${bbbImageMap[data.bbb] ? `<a href="${bbbImageMap[data.bbb]}" target="_blank"><u><i>View BBB Logo</i></u></a><br>` : ``}
            ${data.otherAwards !== "null" ? `<h6>Other Awards, Affiliations, or Organizations:</h6>  
            ${data.otherAwards.split(",").join("<br>")}` : ``}
            <br>
            <h5><u>OTHER NOTES:</u></h5> 
            <p>${data.applicables}</p>
            <p>${data.additionalInfo}</p>
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Homeowner mailings?</h6> Yes
                <h6>Homeowners Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon1}</td>
                        <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                </table></p> 
              
            ` : ``}
            <br>
            <h5><u>COUPONS(There are ${numCoupons()} coupons):</u></h5>        
             <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.radiusOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1 !== "null" ? `${data.radiusOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2 !== "null" ? `${data.radiusOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3 !== "null" ? `${data.radiusOffers.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon1}</td>
                            <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.carrierOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon1}</td>
                            <td>${data.carrierOffers.disclaimer1 !== "null" ? `${data.carrierOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon2}</td>
                            <td>${data.carrierOffers.disclaimer2 !== "null" ? `${data.carrierOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon3}</td>
                            <td>${data.carrierOffers.disclaimer3 !== "null" ? `${data.carrierOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon4}</td>
                            <td>${data.carrierOffers.disclaimer4 !== "null" ? `${data.carrierOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon5 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon5}</td>
                            <td>${data.carrierOffers.disclaimer5 !== "null" ? `${data.carrierOffers.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon6 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon6}</td>
                            <td>${data.carrierOffers.disclaimer6 !== "null" ? `${data.carrierOffers.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon7 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon7}</td>
                            <td>${data.carrierOffers.disclaimer7 !== "null" ? `${data.carrierOffers.disclaimer7}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon8 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon8}</td>
                            <td>${data.carrierOffers.disclaimer8 !== "null" ? `${data.carrierOffers.disclaimer8}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon1}</td>
                            <td>${data.retentionOffers.disclaimer1 !== "null" ? `${data.retentionOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon2}</td>
                            <td>${data.retentionOffers.disclaimer2 !== "null" ? `${data.retentionOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon3}</td>
                            <td>${data.retentionOffers.disclaimer3 !== "null" ? `${data.retentionOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon4}</td>
                            <td>${data.retentionOffers.disclaimer4 !== "null" ? `${data.retentionOffers.disclaimer4}` : `None Entered By Client`}</td>
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;

    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    const autoTaglines = [
        "Quality Auto Service & Repair You Can Trust!",
        "Quality & Affordable Auto Service & Repair You Can Trust",
        "Quality & Dependable Auto Repair You Can Trust",
        "Quality, Honest & Dependable Auto Repair You Can Trust",
        "Car Troubles? Quality Auto Repair You Can Trust!",
        "Car Troubles? Quality Auto Service & Repair You Can Trust!",
        "Car Troubles? We Fix Everything!",
        "Car Troubles? We Repair With Care!",
        "Your Neighborhood Mechanic, No Job Too Big or Too Small!",
        "Today's Technology with Yesterday's Values",
        "Car Repair Doesn't Have to Be Inconvenient!",
        "No Matter What You Drive, We Fix Everything!",
        "Car Troubles? We Get It Right - The First Time!",
        "We Treat Your Car Like Our Own",
        "We Specialize in Exotic & Luxury Cars",
        "Exclusive Audi - Volkswagen - Porsche Repair... It's all we work on for a reason! (or list specific vehicles you service)",
        "Your Trusted European Automobile Experts",
        "Your Trusted Japanese Automobile Experts",
        "Your Trusted BMW and Mini Maintenance Repair Specialists",
        "Other"
    ]

    const taglineString = data.taglines

    const taglines = autoTaglines.filter(tagline => taglineString.includes(tagline))

    const formattedTaglines = taglines.join("<br>")


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
            <h4>We have a new Automotive client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>
            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${data.companyAddress === "null" ? `` : `${data.companyAddress}`}</p>
            ${data.companyAddress2 !== "null" ? `<p>${data.companyAddress2}</p>` : ``}
            <p>${data.companyCity === "null" ? `` : `${data.companyCity}, ${data.companyState === "null" ? `` : `${data.companyState},`}`} ${data.companyZip}</p>
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
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
            <br>
            <h5><u>PHOTOS TO USE:</u></h5>
            <p>${data.vehicles}</p>
            ${data.photos !== "" ? `${photoLinks}` : ``}
            <br>
            <h5><u>You Can Trust Us to Do the Job for You:</u></h5>
            <p>${data.aseTechnicians}</p>
            <p>${data.warranties}</p>
            <p>${data.shuttleLoanerService}</p>
            <p>${data.financing}</p>
            ${data.amenities !== "" ? `<p>Free ${data.amenities.split(",").join(", ")}</p>` : ``}
            <p>${data.sameDayService}</p>
            ${data.approveFirst === "true" ? `<p>All Repairs Approved by You</p>` : ``}
            <br>
            <h5><u>TAGLINES:</u></h5>
            ${data.tagline1 !== "null" ? `<p>${data.tagline1}</p>` : ``}
            ${formattedTaglines
            .split('<br>')
            .map(line => `<p>${line}</p>`)
            .join('')}
            <br>
            <h5><u>RATINGS & REVIEWS:</u></h5>
            ${data.stars.google !== "null" ? `<p>Google: ${data.stars.google}</p>` : ``}
            ${data.stars.other1 !== "null" ? `<p>${data.stars.other1}</p>` : ``}
            ${data.stars.other2 !== "null" ? `<p>${data.stars.other2}</p>` : ``}
            ${data.stars.other3 !== "null" ? `<p>${data.stars.other3}</p>` : ``}
            <br>
            <h5><u>SHOP LOGO:</u></h5> 
            ${!data.logo ? `` : `${logoLinks}`}
            <br>
            <h5><u>AWARDS LOGOS:</u></h5>
            ${!data.awards ? `` : `${awardOrLogoLinks}<br>`}
            ${bbbImageMap[data.bbb] ? `<a href="${bbbImageMap[data.bbb]}" target="_blank"><u><i>View BBB Logo</i></u></a><br>` : ``}
            ${aaaImageMap[data.aaa] ? `<a href="${aaaImageMap[data.aaa]}" target="_blank"><u><i>View AAA Logo</i></u></a><br>` : ``}
            ${napaImageMap[data.napa] ? `<a href="${napaImageMap[data.napa]}" target="_blank"><u><i>View NAPA Logo</i></u></a><br>` : ``}
            ${caOrTxImageMap[data.caOrTx] ? `<a href="${caOrTxImageMap[data.caOrTx]}" target="_blank"><u><i>View CA or TX Logo</i></u></a><br>` : ``}
            ${data.otherAwards !== "null" ? `<h6>Other Awards, Affiliations, or Organizations:</h6>  
            ${data.otherAwards.split(",").join("<br>")}` : ``}
            <br>
            <h5><u>OTHER NOTES:</u></h5> 
            <br>
            
           

            <h5><u>COUPONS(There are ${numCoupons()} coupons):</u></h5>        
             <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.coupons.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.coupons.coupon1}</td>
                        <td>${data.coupons.disclaimer1 !== "null" ? `${data.coupons.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.coupons.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.coupons.coupon2}</td>
                        <td>${data.coupons.disclaimer2 !== "null" ? `${data.coupons.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.coupons.coupon3 !== "null" ? `
                    <tr>
                        <td>${data.coupons.coupon3}</td>
                        <td>${data.coupons.disclaimer3 !== "null" ? `${data.coupons.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.coupons.coupon4 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon4}</td>
                            <td>${data.coupons.disclaimer4 !== "null" ? `${data.coupons.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon5 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon5}</td>
                            <td>${data.coupons.disclaimer5 !== "null" ? `${data.coupons.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon6 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon6}</td>
                            <td>${data.coupons.disclaimer6 !== "null" ? `${data.coupons.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon7 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon7}</td>
                            <td>${data.coupons.disclaimer7 !== "null" ? `${data.coupons.disclaimer7}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.coupons.coupon8 !== "null" ? `
                        <tr>
                            <td>${data.coupons.coupon8}</td>
                            <td>${data.coupons.disclaimer8 !== "null" ? `${data.coupons.disclaimer8}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                </table></p> 
        </body>
    </body>
    
    </html>`
    // #endregion 
}

app.post("/roofing", async (req, res) => {
    // #region roofing Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "Roofing"
        const htmlContent = generateRoofingHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'Roofing_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'Roofing_Brief.docx');

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

function generateRoofingHTML(data) {
    // #region insert data dynamically into html
    // dynamically create "a" tags for photos
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for other photos
    const otherPhotoLinks = data.otherPhotos
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;

    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    const roofingTaglines = [
        "Quality Roofing Experts You Can Trust",
        "Quality & Affordable Roofing Experts You Can Trust",
        "Quality, Trusted & Affordable Roofing Experts",
        "Quality Roofing Installation & Repair You Can Trust",
        "Your Trusted Neighborhood Roofing Specialists",
        "Problems With Your Roof? We Fix Everything!",
        "Problems With Your Roof? We Get It Right - The First Time!",
        "Your Neighborhood Roofing Experts",
        "Other"
    ]

    const taglineString = data.premadeTaglines

    const taglines = roofingTaglines.filter(tagline => taglineString.includes(tagline))

    const formattedTaglines = taglines.join("<br>")


    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.homeownerOffers) {
            if (key.startsWith("coupon")) {
                if (data.homeownerOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.carrierOffers) {
            if (key.startsWith("coupon")) {
                if (data.carrierOffers[key] !== "null") {
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
            <h4>We have a new Roofing client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>

            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``} <br>
            ${data.otherPhotos !== "" ? `${otherPhotoLinks}` : ``}
            <h5><u>SERVICES:</u></h5>
            ${data.services.split(",").join("<br>")}
            <h5><u>You Can Trust Us to Do the Job for You:</u></h5>
            <p>${data.pricing}</p>
            <p>${data.warranties}</p>
            <p>${data.financing}</p>
            <h5><u>TAGLINES:</u></h5>
            ${data.hasTaglines === "true" ?
            `${data.customTaglines.tagline1 !== "null" ? `<p>${data.customTaglines.tagline1}</p>` : ``}
                ${data.customTaglines.tagline2 !== "null" ? `<p>${data.customTaglines.tagline2}</p>` : ``}`
            : ``}
            ${formattedTaglines
            .split('<br>')
            .map(line => `<p>${line}</p>`)
            .join('')}
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
            <p>${data.additionalInfo}</p>
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Homeowner mailings?</h6> Yes
                <h6>Homeowners Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon1}</td>
                        <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
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
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1 !== "null" ? `${data.radiusOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2 !== "null" ? `${data.radiusOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3 !== "null" ? `${data.radiusOffers.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon1}</td>
                            <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.carrierOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon1}</td>
                            <td>${data.carrierOffers.disclaimer1 !== "null" ? `${data.carrierOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon2}</td>
                            <td>${data.carrierOffers.disclaimer2 !== "null" ? `${data.carrierOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon3}</td>
                            <td>${data.carrierOffers.disclaimer3 !== "null" ? `${data.carrierOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon4}</td>
                            <td>${data.carrierOffers.disclaimer4 !== "null" ? `${data.carrierOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon5 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon5}</td>
                            <td>${data.carrierOffers.disclaimer5 !== "null" ? `${data.carrierOffers.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon6 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon6}</td>
                            <td>${data.carrierOffers.disclaimer6 !== "null" ? `${data.carrierOffers.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                </table></p> 
        </body>
    </body>
    
    </html>`
    // #endregion 
}

app.post("/plumbing", async (req, res) => {
    // #region plumbing Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "Plumbing"
        const htmlContent = generatePlumbingHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'Plumbing_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'Plumbing_Brief.docx');

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

function generatePlumbingHTML(data) {
    // #region insert data dynamically into html
    // dynamically create "a" tags for photos
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for other photos
    const otherPhotoLinks = data.otherPhotos
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;


    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.homeownerOffers) {
            if (key.startsWith("coupon")) {
                if (data.homeownerOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.carrierOffers) {
            if (key.startsWith("coupon")) {
                if (data.carrierOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.retentionOffers) {
            if (key.startsWith("coupon")) {
                if (data.retentionOffers[key] !== "null") {
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
            <h4>We have a new Plumbing client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>

            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``} <br>
            ${data.otherPhotos !== "" ? `${otherPhotoLinks}` : ``}
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
            <p>${data.additionalInfo}</p>
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Homeowner mailings?</h6> Yes
                <h6>Homeowners Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon1}</td>
                        <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
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
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1 !== "null" ? `${data.radiusOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2 !== "null" ? `${data.radiusOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3 !== "null" ? `${data.radiusOffers.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon1}</td>
                            <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.carrierOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon1}</td>
                            <td>${data.carrierOffers.disclaimer1 !== "null" ? `${data.carrierOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon2}</td>
                            <td>${data.carrierOffers.disclaimer2 !== "null" ? `${data.carrierOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon3}</td>
                            <td>${data.carrierOffers.disclaimer3 !== "null" ? `${data.carrierOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon4}</td>
                            <td>${data.carrierOffers.disclaimer4 !== "null" ? `${data.carrierOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon5 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon5}</td>
                            <td>${data.carrierOffers.disclaimer5 !== "null" ? `${data.carrierOffers.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon6 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon6}</td>
                            <td>${data.carrierOffers.disclaimer6 !== "null" ? `${data.carrierOffers.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon7 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon7}</td>
                            <td>${data.carrierOffers.disclaimer7 !== "null" ? `${data.carrierOffers.disclaimer7}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon8 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon8}</td>
                            <td>${data.carrierOffers.disclaimer8 !== "null" ? `${data.carrierOffers.disclaimer8}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon1}</td>
                            <td>${data.retentionOffers.disclaimer1 !== "null" ? `${data.retentionOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon2}</td>
                            <td>${data.retentionOffers.disclaimer2 !== "null" ? `${data.retentionOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon3}</td>
                            <td>${data.retentionOffers.disclaimer3 !== "null" ? `${data.retentionOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon4}</td>
                            <td>${data.retentionOffers.disclaimer4 !== "null" ? `${data.retentionOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                </table></p> 
        </body>
    </body>
    
    </html>`
    // #endregion 
}

app.post("/electrical", async (req, res) => {
    // #region electrical Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "Electrical"
        const htmlContent = generateElectricalHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'Electrical_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'Electrical_Brief.docx');

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

function generateElectricalHTML(data) {
    // #region insert data dynamically into html
    // dynamically create "a" tags for photos
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for other photos
    const otherPhotoLinks = data.otherPhotos
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;

    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.homeownerOffers) {
            if (key.startsWith("coupon")) {
                if (data.homeownerOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.carrierOffers) {
            if (key.startsWith("coupon")) {
                if (data.carrierOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.retentionOffers) {
            if (key.startsWith("coupon")) {
                if (data.retentionOffers[key] !== "null") {
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
            <h4>We have a new Electrical client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>

            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``} <br>
            ${data.otherPhotos !== "" ? `${otherPhotoLinks}` : ``}
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
            <p>${data.additionalInfo}</p>
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Homeowner mailings?</h6> Yes
                <h6>Homeowners Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon1}</td>
                        <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
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
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1 !== "null" ? `${data.radiusOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2 !== "null" ? `${data.radiusOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3 !== "null" ? `${data.radiusOffers.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon1}</td>
                            <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.carrierOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon1}</td>
                            <td>${data.carrierOffers.disclaimer1 !== "null" ? `${data.carrierOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon2}</td>
                            <td>${data.carrierOffers.disclaimer2 !== "null" ? `${data.carrierOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon3}</td>
                            <td>${data.carrierOffers.disclaimer3 !== "null" ? `${data.carrierOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon4}</td>
                            <td>${data.carrierOffers.disclaimer4 !== "null" ? `${data.carrierOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon5 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon5}</td>
                            <td>${data.carrierOffers.disclaimer5 !== "null" ? `${data.carrierOffers.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon6 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon6}</td>
                            <td>${data.carrierOffers.disclaimer6 !== "null" ? `${data.carrierOffers.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon7 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon7}</td>
                            <td>${data.carrierOffers.disclaimer7 !== "null" ? `${data.carrierOffers.disclaimer7}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon8 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon8}</td>
                            <td>${data.carrierOffers.disclaimer8 !== "null" ? `${data.carrierOffers.disclaimer8}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon1}</td>
                            <td>${data.retentionOffers.disclaimer1 !== "null" ? `${data.retentionOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon2}</td>
                            <td>${data.retentionOffers.disclaimer2 !== "null" ? `${data.retentionOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon3}</td>
                            <td>${data.retentionOffers.disclaimer3 !== "null" ? `${data.retentionOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon4}</td>
                            <td>${data.retentionOffers.disclaimer4 !== "null" ? `${data.retentionOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                </table></p> 
        </body>
    </body>
    
    </html>`
    // #endregion 
}

app.post("/general", async (req, res) => {
    // #region general Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "General Business"
        const htmlContent = generateGeneralHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'General_Business_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'General_Business_Brief.docx');

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

function generateGeneralHTML(data) {
    // #region insert data dynamically into html
    // dynamically create "a" tags for photos
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for other photos
    const otherPhotoLinks = data.otherPhotos
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;

    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.homeownerOffers) {
            if (key.startsWith("coupon")) {
                if (data.homeownerOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.carrierOffers) {
            if (key.startsWith("coupon")) {
                if (data.carrierOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.retentionOffers) {
            if (key.startsWith("coupon")) {
                if (data.retentionOffers[key] !== "null") {
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
            <h4>We have a new General Business client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>

            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``} <br>
            ${data.otherPhotos !== "" ? `${otherPhotoLinks}` : ``}
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
            <p>${data.additionalInfo}</p>
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Homeowner mailings?</h6> Yes
                <h6>Homeowners Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon1}</td>
                        <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
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
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1 !== "null" ? `${data.radiusOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2 !== "null" ? `${data.radiusOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3 !== "null" ? `${data.radiusOffers.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon1}</td>
                            <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.carrierOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon1}</td>
                            <td>${data.carrierOffers.disclaimer1 !== "null" ? `${data.carrierOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon2}</td>
                            <td>${data.carrierOffers.disclaimer2 !== "null" ? `${data.carrierOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon3}</td>
                            <td>${data.carrierOffers.disclaimer3 !== "null" ? `${data.carrierOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon4}</td>
                            <td>${data.carrierOffers.disclaimer4 !== "null" ? `${data.carrierOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon5 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon5}</td>
                            <td>${data.carrierOffers.disclaimer5 !== "null" ? `${data.carrierOffers.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon6 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon6}</td>
                            <td>${data.carrierOffers.disclaimer6 !== "null" ? `${data.carrierOffers.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon7 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon7}</td>
                            <td>${data.carrierOffers.disclaimer7 !== "null" ? `${data.carrierOffers.disclaimer7}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon8 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon8}</td>
                            <td>${data.carrierOffers.disclaimer8 !== "null" ? `${data.carrierOffers.disclaimer8}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon1}</td>
                            <td>${data.retentionOffers.disclaimer1 !== "null" ? `${data.retentionOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon2}</td>
                            <td>${data.retentionOffers.disclaimer2 !== "null" ? `${data.retentionOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon3}</td>
                            <td>${data.retentionOffers.disclaimer3 !== "null" ? `${data.retentionOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon4}</td>
                            <td>${data.retentionOffers.disclaimer4 !== "null" ? `${data.retentionOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                </table></p> 
        </body>
    </body>
    
    </html>`
    // #endregion 
}

app.post("/dental", async (req, res) => {
    // #region dental Receive JSON from Survey Sparrow
    try {
        const data = req.body
        const company = data.companyName
        const type = "Dental"
        const htmlContent = generateDentalHTML(data)
        const htmlFilePath = path.join(OUTPUT_DIR, 'Dental_Brief.html');
        const docxFilePath = path.join(OUTPUT_DIR, 'Dental_Brief.docx');

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

function generateDentalHTML(data) {
    // #region insert data dynamically into html
    // dynamically create "a" tags for photos
    const photoLinks = data.photos
        .split(",")
        .map(photo => `<a href="${photo}" target="_blank"><u><i>View Photo</i></u></a>`)
        .join("<br>");
    // dynamically create "a" tags for other photos
    const otherPhotoLinks = data.otherPhotos
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
    // dynamically create "a" tags for design refeerence image if it exists
    const designRefImg = data.designReference !== "null" ? `<a href="${data.designReference}" target="_blank"><u><i>View Design Reference</i></u></a><br>` : ``;



    const phone = `${data.companyPhone.slice(3, 6)}-${data.companyPhone.slice(7, 10)}-${data.companyPhone.slice(11)}`

    // counts number of coupons in radiusOffers to then display in the HTML
    function numCoupons() {
        counter = 0
        for (const key in data.homeownerOffers) {
            if (key.startsWith("coupon")) {
                if (data.homeownerOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.radiusOffers) {
            if (key.startsWith("coupon")) {
                if (data.radiusOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.carrierOffers) {
            if (key.startsWith("coupon")) {
                if (data.carrierOffers[key] !== "null") {
                    counter += 1
                }
            }
        }
        for (const key in data.retentionOffers) {
            if (key.startsWith("coupon")) {
                if (data.retentionOffers[key] !== "null") {
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
            <h4>We have a new Dental client doing a "Insert Product and Size"</h4>
            <h5><u>DESIGN:</u></h5>
            ${data.design === "true" ? `<p>There is a specific design.</p>` : ``}
            ${data.designBasedOnWeb === "true" ? `<p>Base the design on their website.</p>` : ``}
            ${data.designInstructions !== "null" ? `${data.designInstructions}<br><br>` : ``}
            ${data.designReference !== "null" ? `${designRefImg}` : ``}<br>


            <h5><u>LOCATION INFORMATION:</u></h5>
            <p>${data.companyName}</p>
            <p>${phone}</p>
            <p>${data.website}</p>
            ${data.license !== "null" ? `<p>${data.license}</p>` : ``}
            ${data.onlineService !== "true" ? `<p>Call Today to Schedule Your Appointment!</p>` : `
                <p><strong>“Insert Call to Action” Based on Q4</strong></p>
                <p>OR Call Today or Conveniently Schedule Online!</p>
                <p>OR Call Today or Conveniently Schedule Online!  (Insert QR Code) Scan Here to Easily Schedule Your Appointment!</p>`}
            <h5><u>PHOTOS TO USE:</u></h5>
            ${data.photos !== "" ? `${photoLinks}` : ``} <br>
            ${data.otherPhotos !== "" ? `${otherPhotoLinks}` : ``}
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
            <p>${data.additionalInfo}</p>
            ${JSON.parse(data.homeOwner) ? `
                <h6>New Homeowner mailings?</h6> Yes
                <h6>Homeowners Offers: </h6>
                <table>
                    <tr>
                        <th>Coupon</th>
                        <th>Disclaimer</th>
                    </tr>
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon1}</td>
                        <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
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
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon1}</td>
                        <td>${data.radiusOffers.disclaimer1 !== "null" ? `${data.radiusOffers.disclaimer1}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon2}</td>
                        <td>${data.radiusOffers.disclaimer2 !== "null" ? `${data.radiusOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.radiusOffers.coupon3 !== "null" ? `
                    <tr>
                        <td>(RADIUS OFFER) ${data.radiusOffers.coupon3}</td>
                        <td>${data.radiusOffers.disclaimer3 !== "null" ? `${data.radiusOffers.disclaimer3}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.homeownerOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon1}</td>
                            <td>${data.homeownerOffers.disclaimer1 !== "null" ? `${data.homeownerOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.homeownerOffers.coupon2 !== "null" ? `
                    <tr>
                        <td>(NEW HOMEOWNER OFFER) ${data.homeownerOffers.coupon2}</td>
                        <td>${data.homeownerOffers.disclaimer2 !== "null" ? `${data.homeownerOffers.disclaimer2}` : `None Entered By Client`}</td>
                    </tr>`: ``}
                    ${data.carrierOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon1}</td>
                            <td>${data.carrierOffers.disclaimer1 !== "null" ? `${data.carrierOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon2}</td>
                            <td>${data.carrierOffers.disclaimer2 !== "null" ? `${data.carrierOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon3}</td>
                            <td>${data.carrierOffers.disclaimer3 !== "null" ? `${data.carrierOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon4}</td>
                            <td>${data.carrierOffers.disclaimer4 !== "null" ? `${data.carrierOffers.disclaimer4}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon5 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon5}</td>
                            <td>${data.carrierOffers.disclaimer5 !== "null" ? `${data.carrierOffers.disclaimer5}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon6 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon6}</td>
                            <td>${data.carrierOffers.disclaimer6 !== "null" ? `${data.carrierOffers.disclaimer6}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon7 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon7}</td>
                            <td>${data.carrierOffers.disclaimer7 !== "null" ? `${data.carrierOffers.disclaimer7}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.carrierOffers.coupon8 !== "null" ? `
                        <tr>
                            <td>(CARRIER ROUTE OFFER) ${data.carrierOffers.coupon8}</td>
                            <td>${data.carrierOffers.disclaimer8 !== "null" ? `${data.carrierOffers.disclaimer8}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon1 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon1}</td>
                            <td>${data.retentionOffers.disclaimer1 !== "null" ? `${data.retentionOffers.disclaimer1}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon2 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon2}</td>
                            <td>${data.retentionOffers.disclaimer2 !== "null" ? `${data.retentionOffers.disclaimer2}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon3 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon3}</td>
                            <td>${data.retentionOffers.disclaimer3 !== "null" ? `${data.retentionOffers.disclaimer3}` : `None Entered By Client`}</td>
                        </tr>`: ``}
                    ${data.retentionOffers.coupon4 !== "null" ? `
                        <tr>
                            <td>(RETENTION OFFER) ${data.retentionOffers.coupon4}</td>
                            <td>${data.retentionOffers.disclaimer4 !== "null" ? `${data.retentionOffers.disclaimer4}` : `None Entered By Client`}</td>
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
        bcc: "sharkymailson@gmail.com",
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