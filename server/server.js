const cors = require('cors')
const express = require('express')
const excel = require('exceljs')
const {createShippingRatesExcel} = require('./app/mysqlAdapter')

const PORT = 3000;
const HOST = '0.0.0.0';

const app = express()

app.use(cors())

app.get('/', (req, res) => {
  res.setHeader(
    'Content-Type',
    'text/html'
  )
  res.send('<div style="text-align: center"><span>Visit <strong>"/download"</strong> to download the Shipping Rates</span></div>')
})

app.get('/download', async (req, res)=>{
    let workbook = new excel.Workbook()
    workbook = await createShippingRatesExcel()

    res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      )
      res.setHeader(
        'Content-Disposition',
        'attachment; filename=' + 'Current-Shipping-Rates.xlsx'
      )
      
    return workbook.xlsx.write(res).then(function () {
        res.status(200).end();
    })
})

app.listen(PORT, ()=>{`App server running on http://${HOST}:${PORT}`})
