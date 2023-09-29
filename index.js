const express = require('express');
const app = express();
const cors = require("cors");
const { exec } = require('child_process');

app.use(cors());
app.use(express.json());

app.post('/manipulatePpt', (req, res) => {
    res.set('Access-Control-Allow-Origin', 'http://localhost:4200');
    const pptData = req.body;
    const fileName = pptData.fileName + '.pptx';

    const old_text = pptData.slideData.map(item => item.old_text).join('|');
    const new_text = pptData.slideData.map(item => item.new_text).join('|');
    const slide_indexes = pptData.slideData.map(item => item.slideIndex).join('|');

    const command = `python update_presentation.py "${fileName}" "${old_text}" "${new_text}" "${slide_indexes}"`;

    exec(command, (error, stdout, stderr) => {
        if (error)
        {
            console.error(`Error: ${error}`);
            res.status(500).send('Internal Server Error');
            return;
        }
        console.log(`stdout: ${stdout}`);
        console.error(`stderr: ${stderr}`);

        const modifiedFilePath = `${__dirname}/modified_presentation.pptx`;
        res.sendFile(modifiedFilePath);
    });
});

app.listen(3000, () => console.log('App listening on port 3000!'));
