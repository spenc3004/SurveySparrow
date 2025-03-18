const express = require('express')
const app = express()
require('dotenv').config()

const PORT = process.env.PORT || 3001

//let data = require('./docs/response.json')

app.use(express.json())

app.get("/", (req, res) => {
    res.send('Something else!')
})


app.post("/ss", (req, res) => {
    const data = req.body
    console.log(data)
    res.status(200).send("OK")
})



app.listen(PORT, () => {
    console.log(`This is running on port http://localhost:${PORT}`)
})