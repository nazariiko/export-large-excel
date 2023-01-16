const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const Stream = require('stream');

const PORT = process.env.PORT || 3000

const app = express();

app.use(function(req, res, next) {
  res.header("Access-Control-Allow-Origin", "http://localhost:3000");
  res.header('Access-Control-Allow-Methods', 'GET, PUT, POST, DELETE')
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  next();
});
app.use(cors())
app.use(
  bodyParser.urlencoded({
    extended: true,
    limit: '500mb',
    type: 'application/json'
  })
)
app.use(express.json());


app.post('/', (req, res) => {
  const data = req.body
  const temp = Object.keys(data)[0]
  const parsedData = JSON.parse(`[${temp}]`)

  const stream = new Stream.PassThrough();
  let workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: stream,
  });
  let worksheet = workbook.addWorksheet('Main sheet'); 
      
  const resultRows = []
  resultRows.push(Object.keys(parsedData[0]))
  parsedData.forEach(row => {
    resultRows.push(Object.values(row))
  })

  resultRows.forEach(row => {
    worksheet.addRow(row).commit()
  })

  worksheet.commit();
  workbook.commit();

  res.attachment('yourfile.xlsx');
  stream.pipe(res)
})

  
try {
  app.listen(PORT, () => console.log('listening on port', PORT))
} catch (error) {
  console.log('Something went wrong: ' + error)
}