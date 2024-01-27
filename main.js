// Please,  make sure that modules are installed

//note: node.js must to be installed!!

// Use the following commands:


//npm install
//npm init -y
//npm install googleapis@105 @google-cloud/local-auth@2.1.0 --save



// imports needed modules

const fs = require('fs').promises
const path = require('path')
const process = require('process')
const {authenticate} = require('@google-cloud/local-auth')
const {google} = require('googleapis')

// Id to spreadsheet and range of document data

const spreadsheetId = '1ceRUKCLu-O5C1U7FxhPBHFfRzyAzwG5OhjQhaSIteXw'
const range = 'engenharia_de_software!A4:H27'


// permisssion scope to acess the sheet

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

// credentials and token path

const TOKEN_PATH = path.join(process.cwd(), 'token.json')
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json')


 // Reads previously authorized credentials

 /**
 * loads credentials if it already exists
 *
 * @return {Promise<OAuth2Client|null>} OAuth2Client object or null if don't 
 */

async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH)
    const credentials = JSON.parse(content)
    
    console.log('////////////////////////////////////////////////')
    console.log('loading credentials...')
    console.log('////////////////////////////////////////////////')

    return google.auth.fromJSON(credentials)


  } catch (err) {
    return null
  }
}

// Saves credentials to a file comptible with GoogleAUth.fromJSON.
/**
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */

async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH)
  const keys = JSON.parse(content)
  const key = keys.installed || keys.web
  const payload = JSON.stringify({
    type: 'authorized_user',
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  })
  await fs.writeFile(TOKEN_PATH, payload)
}


// load or request auth to call APIs

/**
 * @return {Promise<OAuth2Client>} Authorized OAuth2Client object
 */
 
async function authorize() {
  let client = await loadSavedCredentialsIfExist()
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client)
  }
  return client;
}

/**
 * @see https://docs.google.com/spreadsheets/d/1ceRUKCLu-O5C1U7FxhPBHFfRzyAzwG5OhjQhaSIteXw/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */

//
async function getStudents(auth) {

    console.log('////////////////////////////////////////////////')
    console.log('Getting data...') 
    console.log('////////////////////////////////////////////////')

    const sheets = google.sheets({version: 'v4', auth})
    const res = await sheets.spreadsheets.values.get({

    // Sheet ID and range of tables
    spreadsheetId: '1ceRUKCLu-O5C1U7FxhPBHFfRzyAzwG5OhjQhaSIteXw',
    range: 'engenharia_de_software!A4:H27'
  });

  // If no data is found, return an console.log
  const rows = res.data.values;
  if (!rows || rows.length === 0) {
    console.log('No data found.')
    return;
  }

  const updatedRows = []


  console.log('////////////////////////////////////////////////')
  console.log('applying logic...')
  console.log('////////////////////////////////////////////////')


  //  iterate each line
  rows.forEach(async (row) => {
    const studentName = row[1]
    const P1 = parseFloat(row[3])
    const P2 = parseFloat(row[4])
    const P3 = parseFloat(row[5])
    const absences = parseInt(row[2])

    // calculate grades average
    const average = (P1 + P2 + P3) / 3

    let status = ''

    
    //verifies if the student have to many absences and set the "Nota para aprovação final" to 0 
    if (absences > 0.25 * 60) {
      status = 'Reprovado por Falta'
      row[6] = status
      row[7] = 0
    } 
    
    //verifies if is a grade failure    
    else if (average < 50) {
      status = 'Reprovado por Nota'
      row[6] = status
      row[7] = 0;
    } else if (average < 70) {
      
      // calculate the final exam minimum required grade
      const naf = Math.ceil(100 - average)

      status = 'Exame Final'
      row[6] = status
      row[7] = naf
    } 
    
    // Else, is a grade sucess 

    else {
      status = 'Aprovado'
      row[6] = status
      row[7] = 0
    }

    //inject the status in the excel table
    updatedRows.push(row)

  })

  console.log('////////////////////////////////////////////////')
  console.log('updating excel table...')
  console.log('////////////////////////////////////////////////')
  // update the excel table

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'RAW',
    resource: { values: updatedRows },
  });

  console.log('////////////////////////////////////////////////')
  console.log('Results updated, take a look in the spreadsheet.')
  console.log('////////////////////////////////////////////////')

}

authorize().then(getStudents).catch(console.error);