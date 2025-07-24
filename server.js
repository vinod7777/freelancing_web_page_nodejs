const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3005;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));
app.use(bodyParser.json());

// Sign Up route
app.post('/signup', (req, res) => {
    const userData = req.body;
    userData.loginCount = 1;
    const timestamp = Date.now();
    const username = `${userData.name.toLowerCase().replace(/\s+/g, '')}_${timestamp}`;
    userData.username = username;

    const filePath = path.join(__dirname, 'data.xlsx');
    let workbook;

    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
    }

    const sheetName = 'Sheet1';
    let existingData = [];

    if (workbook.SheetNames.includes(sheetName)) {
        existingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }

    const dataExists = existingData.some(data => data.email === userData.email);
    if (dataExists) {
        return res.status(409).json({ message: 'User  already exists' });
    } else {
        existingData.push(userData);
        const updatedWorksheet = xlsx.utils.json_to_sheet(existingData);
        workbook.Sheets[sheetName] = updatedWorksheet;
        xlsx.writeFile(workbook, filePath);
        return res.render('welcome_page', { name: userData.name });
    }
});

// Sign In route
app.post('/signin', async (req, res) => {
    const userData = req.body;
    const filePath = path.join(__dirname, 'data.xlsx');

    try {
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ error: 'No users found. Please sign up first.' });
        }

        const workbook = xlsx.readFile(filePath);
        const sheetName = 'Sheet1';
        let existingData = [];

        if (workbook.SheetNames.includes(sheetName)) {
            existingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        }

        const userExists = existingData.find(data => data.email === userData.email && data.password === userData.password);

        if (userExists) {
            if (userExists.loginCount === 0) {
                userExists.loginCount += 1; 
                const updatedWorksheet = xlsx.utils.json_to_sheet(existingData);
                workbook.Sheets[sheetName] = updatedWorksheet;
                xlsx.writeFile(workbook, filePath);
                return res.render('welcome_page', { name: userExists.name });
            } else {
                userExists.loginCount += 1; 
                const updatedWorksheet = xlsx.utils.json_to_sheet(existingData);
                workbook.Sheets[sheetName] = updatedWorksheet;
                xlsx.writeFile(workbook, filePath);
                return res.render('profile', { name: userExists.name });
            }
        } else {
            return res.status(401).json({ alert: 'Invalid email or password. Please create an account or try again.' });
        }
    } catch (error) {
        console.error(error);
        return res.status(500).json({ alert: 'An error occurred while processing your request.' });
    }
});

// Subscribe route
app.post('/subscribe', (req, res) => {
    const email = req.body.email;

    const transporter = nodemailer.createTransport({
        host: 'smtp.gmail.com',
        port: 587,
        secure: false,
        auth: {
            user: 'giglancer97@gmail.com',
            pass: 'jtqp rxfm gdmu ovtw'
        }
    });

    const mailOptions = {
        from: 'giglancer97@gmail.com',
        to: email,
        subject: 'Thank You for Subscribing!',
        text: 'Thank you for subscribing to our newsletter! We appreciate your interest.'
    };
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return res.status(500).send('Error sending email: ' + error.toString());
        }
        const workbookPath = path.join(__dirname, 'subscribers.xlsx');
        let workbook;
        
        if (fs.existsSync(workbookPath)) {
            workbook = xlsx.readFile(workbookPath);
        } else {
            workbook = xlsx.utils.book_new();
        }
        const sheetName = 'Subscribers';
        let subscriberData = [];

        if (workbook.SheetNames.includes(sheetName)) {
            subscriberData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        }

        const subscriberExists = subscriberData.some(subscriber => subscriber.email === email);
        if (!subscriberExists) {
            subscriberData.push({ email: email });
            const updatedWorksheet = xlsx.utils.json_to_sheet(subscriberData);
            workbook.Sheets[sheetName] = updatedWorksheet;
            xlsx.writeFile(workbook, workbookPath);
        }

        res.status(200).send('Subscription successful! A confirmation email has been sent.');
    });
});

app.listen(PORT, () => {
    console.log(`Server is running on port http://localhost:${PORT}`);
});