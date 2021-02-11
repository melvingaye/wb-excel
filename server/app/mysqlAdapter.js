const excel = require('exceljs')
var mysql = require('mysql2');
const groupBy = require('lodash.groupby')

const worksheets = {
    DOMESTIC_STANDARD: 'Domestic Standard Rates',
    DOMESTIC_EXPEDITED:'Domestic Expedited Rates',
    DOMESTIC_NEXT_DAY:'Domestic Next Day Rates',
    INTERNATIONAL_ECONOMY: 'International Economy Rates',
    INTERNATIONAL_EXPEDITED:'International Expedited Rates'
}

//pool error doesn't seem to fire without the connectTimeout
const pool = new mysql.createPool({host: 'db', user: 'mysql', password: 'password', database: 'shipping', connectTimeout: 2000})

const poolPromise = pool.promise()

pool.on('error', err => {
  console.log(`ERROR CONNECTING: ${err}`)
  //exit process to trigger docker-compose on-failure restart
  process.exit(1)
})

createShippingRatesExcel()

async function createShippingRatesExcel(){
    const queries = await runQueries()
    return processQueryResults(queries)
}

async function runQueries(){

    //await poolPromise

    try
    {
        const queryResults = []
        //hard coded the quries for now, but ideally I'd make this a stored procedure and potential params would be client_id, shipping_speed, and locale
        const [domesticStandardQuery] = await poolPromise.query('Select  * from rates where client_id = 1240  and shipping_speed = \'standard\' and locale = \'domestic\'')
        const [domesticExpeditedQuery] =  await poolPromise.query('Select * from rates where client_id = 1240  and shipping_speed = \'expedited\' and locale = \'domestic\'')


        const [domesticNextDayQuery] = await poolPromise.query('Select * from rates where client_id = 1240  and shipping_speed = \'nextDay\' and locale = \'domestic\'')
   

        const [internationalEconomyQuery] =  await poolPromise.query('Select * from rates where client_id = 1240  and shipping_speed = \'intlEconomy\' and locale = \'international\'')
    

        const [internationalExpeditedQuery] = await poolPromise.query('Select * from rates where client_id = 1240  and shipping_speed = \'intlExpedited\' and locale = \'international\'')
       

        queryResults.push(domesticStandardQuery)
        queryResults.push(domesticExpeditedQuery)
        queryResults.push(domesticNextDayQuery)
        queryResults.push(internationalEconomyQuery)
        queryResults.push(internationalExpeditedQuery)
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
