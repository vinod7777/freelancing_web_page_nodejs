const express = require('express');
const bodyParser = require('body-parser');
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
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'register.html'));
});

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

app.listen(PORT, () => {
    console.log(`Sign-In Server is running on http://localhost:${PORT}`);
});