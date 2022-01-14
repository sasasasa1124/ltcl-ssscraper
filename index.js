// import modules;
const utils = require('./utils');
const express = require("express");
const { google } = require("googleapis");
const { GoogleSpreadsheet } = require('google-spreadsheet');
const PORT = process.env.PORT || 5000;
const cors = require('cors');
var conf = require('config');
const path = require('path');
// require('dotenv').config();
const cred = require(process.env.GOOGLE_APPLICATION_CREDENTIALS);

// google API and spreadsheet initialization;
const auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
    scopes:"https://www.googleapis.com/auth/spreadsheets",
});

const spreadsheetsId = process.env.SPREADSHEET_ID;
const doc = new GoogleSpreadsheet(spreadsheetsId);

// slack reminder initialization;
const { IncomingWebhook } = require('@slack/webhook');
const key_webhook = new IncomingWebhook(process.env.SLACK_KEY_HOOK);
const self_study_webhook = new IncomingWebhook(process.env.SLACK_SELF_STUDY_HOOK);

// express initialization;
const app = express();
app.use(cors());
app.use(express.urlencoded({ extended: false }));
app.use(express.json());

app.get("/api/create", async (req,res) => {
    await doc.useServiceAccountAuth(cred);
    await doc.loadInfo();
    const today = utils.today;
    const isCreated = (typeof doc.sheetsByTitle[`${today['year']}年${today['month']}月`] !== 'undefined');
    const sheet = (isCreated ? doc.sheetsByTitle[`${today['year']}年${today['month']}月`] : await doc.addSheet({'title' : `${today['year']}年${today['month']}月`}));
    await sheet.loadCells('A1:A10');
    console.log(sheet.cellStats);
    const firstCell = sheet.getCell(0,0);
    firstCell.formula = `=IMPORTRANGE("1-Rdn9b--9e2lNH25KWILFOXkupQyBDbYbm_ivNbmMHc/edit#gid=1731820269","${today['year']}年${today['month']}月!A1:AK50")`
    await sheet.saveUpdatedCells();
    res.send(sheet.title);
});

app.get("/api/test", async (req, res) => {
    // initialization; 
    const today = utils.today;
    const mentors = conf.name_table;
    const dates = utils.dateFetch(today);

    // fetch raw data; loading data
    const row = await utils.fetchGoogleSheets(auth,spreadsheetsId,[`${today['year']}年${today['month']}月!A1:AG50`]);
    const matrix = utils.formatRawRow(row);

    res.send(matrix.map((o) => o));
});


// key API
app.get("/api/key", async (req,res) => {
    // initialization
    const today = utils.today;
    const tommorow = utils.tommorow;
    const ranges = (today['month'] == tommorow['month'] ? [`${today['year']}年${today['month']}月!A1:AG50`] : [`${today['year']}年${today['month']}月!A1:AG50`,`${tomorrow['year']}年${tomorrow['month']}月!A1:AG50`]);

    // fetch data from ss
    const row = await utils.fetchGoogleSheets(auth,spreadsheetsId,ranges);
    const matrix = utils.formatRawRow(row);

    // fetch key cells
    keyNames = conf.spreadsheet.keyNames;
    keyCells = utils.searchColumnNum(keyNames,matrix);

    // fetch the key holders for today/tommorow
    todayKeys = [];
    todayCell = utils.dateCell(today['string'],matrix)[today[['string']]];
    keyCells.map((element,index) => {
        todayKeys.push({
            "key": keyNames[index],
            "value" : utils.fetchColumnDate(element['columnName'],todayCell,matrix)
        });
    });
    tommorowKeys = [];
    tommorowCell = utils.dateCell(tommorow['string'],matrix)[tommorow[['string']]];
    keyCells.map((element,index) => {
        tommorowKeys.push({
            "key": keyNames[index],
            "value" : utils.fetchColumnDate(element['columnName'],tommorowCell,matrix)
        });
    });

    // compare today and tommorow
    let changes = [];
    keyNames.map((keyName) => {
        const todayKey = todayKeys.find(element => element.key == keyName);
        const tommorowKey = tommorowKeys.find(element => element.key == keyName);
        if (todayKey.value !== tommorowKey.value) {
            changes.push(`${utils.nameToMention(todayKey.value)} → ${utils.nameToMention(tommorowKey.value)}`);
        }
    });

    // create and send message to Slack
    const msg = utils.keymsgGenerator(today,todayKeys,changes);
    (async () => {
        await key_webhook.send({
            username: "鍵bot",
            icon_url:"https://1.bp.blogspot.com/-NKE9shT7SNk/W1a48Igi0pI/AAAAAAABNi4/mHCXBty4XsAdUEXKcvCXWWc4-KVbJTowQCLcBGAs/s800/monban_heitai_seiyou.png",
            text: msg,
        });
      })();
    console.log(msg);
    res.send(msg);
})

// self_study
app.get("/api/self_study", async (req,res) => {
    // initialization
    const today = utils.today;
    // fetch data from ss
    const row = await utils.fetchGoogleSheets(auth,spreadsheetsId,[`${today['year']}年${today['month']}月!A1:AG50`]);
    const matrix = utils.formatRawRow(row);

    // fetch self_study cells
    self_studyNames = ["受付兼オンライン質問部屋常駐\n（Aスペース）", "オンライン質問部屋常駐\n（Bスペースorオンライン）"]
    self_studyCells = utils.searchColumnNum(self_studyNames,matrix);

    // fetch the manager of self_study_room
    todayManagers = [];
    todayCell = utils.dateCell(today['string'],matrix)[today[['string']]];
    self_studyCells.map((element,index) => {
        todayManagers.push({
            "Room": self_studyNames[index],
            "value" : utils.fetchColumnDate(element['columnName'],todayCell,matrix)
        });
    });

    // create and send message to Slack  
    const msg = utils.self_studymsgGenerator(today,todayManagers);

    (async () => {
        await self_study_webhook.send({
            username: "自習室報告BOT",
            icon_url: "https://3.bp.blogspot.com/-u9SuZaqht48/VOsW_c0BG6I/AAAAAAAArzI/95U800_mtNE/s800/job_juku_koushi.png",
            text: msg,
        });
      })();
      res.send(msg);
    }
);

app.get('*',(req,res) => {
    res.send('there is no corresponding API...');
});

app.listen(PORT,(req, res) => console.log(`server listening on ${PORT}`));
