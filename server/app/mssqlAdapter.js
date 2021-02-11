const excel = require('exceljs')
const sql = require('mssql/msnodesqlv8')
const groupBy = require('lodash.groupby')

const pool = new sql.ConnectionPool({
  database: 'PantryButler',
  server: '(localdb)\\mssqllocaldb',
  driver: 'msnodesqlv8',
  options: {
    trustedConnection: true
  }
})

const worksheets = {
    DOMESTIC_STANDARD: 'Domestic Standard Rates',
    DOMESTIC_EXPEDITED:'Domestic Expedited Rates',
    DOMESTIC_NEXT_DAY:'Domestic Next Day Rates',
    INTERNATIONAL_ECONOMY: 'International Economy Rates',
    INTERNATIONAL_EXPEDITED:'International Expedited Rates'
}

const poolConnect = pool.connect();

pool.on('error', err => console.log(err))

createShippingRatesExcel()

async function createShippingRatesExcel(){
    const queries = await runQueries()
    return processQueryResults(queries)
}

async function runQueries(){

    await poolConnect

    try
    {
        const queryResults = []
        //hard coded the quries for now, but ideally I'd make this a stored procedure and potential params would be client_id, shipping_speed, and locale
        const domesticStandardQuery = await pool.request().query('Select  * from rates where client_id = 1240  and shipping_speed = \'standard\' and locale = \'domestic\'')
        const domesticExpeditedQuery =  await pool.request().query('Select * from rates where client_id = 1240  and shipping_speed = \'expedited\' and locale = \'domestic\'')
        const domesticNextDayQuery = await pool.request().query('Select * from rates where client_id = 1240  and shipping_speed = \'nextDay\' and locale = \'domestic\'')
        const internationalEconomyQuery =  await pool.request().query('Select * from rates where client_id = 1240  and shipping_speed = \'intlEconomy\' and locale = \'international\'')
        const internationalExpeditedQuery = await pool.request().query('Select * from rates where client_id = 1240  and shipping_speed = \'intlExpedited\' and locale = \'international\'')
        
        queryResults.push(domesticStandardQuery.recordset)
        queryResults.push(domesticExpeditedQuery.recordset)
        queryResults.push(domesticNextDayQuery.recordset)
        queryResults.push(internationalEconomyQuery.recordset)
        queryResults.push(internationalExpeditedQuery.recordset)
        return queryResults;
    }
    catch(err)
    {
        console.log(err)
        return err
    }
}

function processQueryResults(results){
    const sheetData = []

    try 
    {
        
        for(const entry of results){
        
            //Group the rates by weight - essentially need to know what the price is at a certain weight in a certain zone
            const zones = groupBy(entry, (rate)=>{
                return JSON.parse(JSON.stringify(rate.start_weight))
            })
    
            const sortedZone = createZoneEntities(zones)
            const columns = createColumns(sortedZone[0])
            const sheetName = createSheetName(sortedZone[0])
            sheetData.push({sheetName, columns, sortedZone})
              
        }

       return saveWorkbook(sheetData)
    } 
    catch (err) 
    {
        console.log(err)
        return err
    }
   
}

function saveWorkbook(worksheets){

    const workbook = new excel.Workbook();

    try 
    {
        for(const sheet of worksheets){
            const worksheet = workbook.addWorksheet(sheet.sheetName)
            worksheet.columns = sheet.columns
            worksheet.addRows(sheet.sortedZone)
        }

        return workbook;
    } 
    catch (err) 
    {   
        console.log(err)
        return err
    }
  
}

function createZoneEntities(rates){
    const sortedZones = []
    for(const [key, value] of Object.entries(rates)){
        let object = {}
        for(var i=0; i <value.length; i++){
            //only update key/value pairs that change
            object = {...object, 
                start_weight: value[0].start_weight, 
                end_weight: value[0].end_weight, 
                [`zone_${value[i].zone}`]: value[i].rate,
                shipping_speed: value[0].shipping_speed,
                locale: value[0].locale
            }
        }
        sortedZones.push(object)
    }

    return sortedZones
}

function createColumns(zoneEntities){
    const cols = []
    for(const [key] of Object.entries(zoneEntities)){
        if(key !== 'locale' && key !== 'shipping_speed'){
            let col = {header: `${capitalize(key.split('_')[0])} ${isCharNumber(key.split('_')[1]) ? key.split('_')[1] : capitalize(key.split('_')[1])}`, key: key, width: 30}
            cols.push(col)
        }
    }
 
    return cols
}

function createSheetName(zone){
    //Hard coded these values but if the columns change, it'd have to be updated also
    if(zone.locale === 'international' && zone.shipping_speed.includes('Expedited')){
        return worksheets.INTERNATIONAL_EXPEDITED
    }

    if(zone.locale === 'international' && zone.shipping_speed.includes('Economy')){
        return worksheets.INTERNATIONAL_ECONOMY
    }

    if(zone.locale === 'domestic' && zone.shipping_speed.includes('expedited')){
        return worksheets.DOMESTIC_EXPEDITED
    }

    if(zone.locale === 'domestic' && zone.shipping_speed.includes('nextDay')){
        return worksheets.DOMESTIC_NEXT_DAY
    }

    if(zone.locale === 'domestic' && zone.shipping_speed.includes('standard')){
        return worksheets.DOMESTIC_STANDARD
    }
}

function isCharNumber(character) {
    return character >= '0' && character <= '9';
}

function capitalize(string){
    if (typeof string !== 'string') return ''
    return string.charAt(0).toUpperCase() + string.slice(1)
}


module.exports = {
    createShippingRatesExcel
}
