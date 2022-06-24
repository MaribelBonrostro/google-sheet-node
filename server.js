const {GoogleSpreadsheet} = require('google-spreadsheet');
const {MongoClient} = require('mongodb');
const creds = require('./credentials.json');
require('dotenv/config');

async function main() {
    // Set up mongodb client connection
  var mongoDB = process.env.MONGODB_URI;
  const client = new MongoClient(mongoDB);
  const {DB_NAME, COLLECTION} = process.env;

  try {
  await client.connect();
  const data = await accessSpreadSheet(); // get data from google-sheet
  // Use connect method to connect to the server
  console.log('Connected successfully to server');
  const db = client.db(DB_NAME);
  const collection = db.collection(COLLECTION);
  console.log(DB_NAME)

    for(let i =0; i< data.length; i++) {
      const formData = {
        'order_date': data[i].OrderDate,
        'region': data[i].Region,
        'representative': data[i].Rep,
        'item': data[i].Item,
        'units': data[i].Units,
        'unit_cost': data[i].UnitCost,
        'total': data[i].Total,
      };
      const result = await collection.insertOne(formData)
      console.log(result)
    }
  }  finally {
    // Close the connection to the MongoDB cluster
    await client.close();
  }
}
main().catch(console.error);

async function accessSpreadSheet() {
  const doc = new GoogleSpreadsheet(process.env.SHEET_ID); // sheet id from URL
  await doc.useServiceAccountAuth(creds);

  await doc.loadInfo(); // loads document properties and worksheets
  const sheet = doc.sheetsByIndex[0]; // the first sheet
  return await sheet.getRows();
}
