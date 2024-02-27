const express = require('express');
const xlsx = require('xlsx')
const bodyParser = require('body-parser');
const path = require('path')
const app = express();


app.set('view engine', 'ejs');
app.set('views', 'views')

app.use(express.json())
app.use(bodyParser.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, 'public')));


app.get('/', (req, res, next) => {
    const workbook = xlsx.readFile('./AppTemplate.xlsx');  // Step 2
    let workbook_sheet = workbook.SheetNames;
    console.log(workbook_sheet)
    let workbook_response = xlsx.utils.sheet_to_json(        // Step 4
        workbook.Sheets[workbook_sheet[0]]
    );
    console.log(workbook_response);
    res.render('indexView', { data: workbook_response })
})

app.post('/', (req, res, next) => {
    console.log('PO')
    console.log(req.body) 
    console.log(req.body.date)

    const workbook = xlsx.readFile('./AppTemplate.xlsx');  // Step 2
    let workbook_sheet = workbook.SheetNames;
    xlsx.utils.sheet_add_aoa(workbook.Sheets[workbook_sheet[0]], [[req.body.date, req.body.text, req.body.startdate, req.body.enddate ]], { origin: -1 });
    xlsx.writeFile(workbook, 'Apptemplate.xlsx')
    res.redirect('/');
})

app.listen(3000, () => {
    console.log('adad')
})
