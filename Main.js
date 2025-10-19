import { utils, writeFile } from "./Data/xlsx.mjs";
import { dbID, initEL, KadLog } from "./KadUtils/KadUtils.js";

const Lbl_missingIDs = initEL({ id: "idLbl_missingIDs", resetValue: "" });
const Lbl_loadedSOU = initEL({ id: "idLbl_loadedSOU", resetValue: "..." });
const Lbl_fileName = initEL({ id: "idLbl_fileName", resetValue: "*.xlsx" });
initEL({ id: "idVin_mainAssemblyNr", fn: getMainNumber, resetValue: "MM-Nummern Anlage" });
initEL({ id: "idVin_mainAssemblyName", fn: getMainName, resetValue: "Anlagename" });
const Area_inputMenge = initEL({ id: "idArea_inputMenge", fn: readData, resetValue: "Mengenstückliste hier einfügen" });
const Area_inputStruktur = initEL({ id: "idArea_inputStruktur", fn: readData, resetValue: "Strukturstückliste hier einfügen" });

initEL({ id: "idBtn_infoUpload", fn: openInfoUpload });
initEL({ id: "idBtn_infoCloseUpload", fn: closeInfoUpload, resetValue: "Schliessen" });
initEL({ id: "idBtn_infoError", fn: openInfoError });
initEL({ id: "idBtn_infoCloseError", fn: closeInfoError, resetValue: "Schliessen" });

const Btn_download = initEL({ id: "idBtn_download", fn: startDownload, resetValue: "Download" });

window.onload = mainSetup;

function mainSetup() {
  enableDownload(true);
  Lbl_fileName.KadReset();
  populateTokenList("idUL_Upload", ulInfoUpload);
  populateTokenList("idUL_Error", ulInfoError);
}

const ulInfoUpload = [
  //
  'Strukturliste und Mengenstückliste über "Zwischenablage" kopieren und in die Textfelder einfügen.',
  "MM-Nummer der Hauptanlage im linken Feld eintragen",
  "Der Anlagenname im rechten Feld ist optional.",
  'Der Button "SOU-Liste als *.xlsx" ist blockiert wenn keine sechsstellige MM-Nummer der Anlage eingegeben wurde und keine Daten vorhanden sind.',
];
const ulInfoError = [
  //
  'MM-Numern von "Struktur" nicht in "Menge" enthalten!',
  'Liste von MM-Nummern, die in der "Struktur"-Tabelle enthalten sind aber nicht in der "Mengen"-Tabelle.',
  "Durch diesen Fehler wird keine Excel-Datei ausgegeben!",
];

function openInfoUpload() {
  dbID("idDia_Upload").showModal();
}
function closeInfoUpload() {
  dbID("idDia_Upload").close();
}
function openInfoError() {
  dbID("idDia_Error").showModal();
}
function closeInfoError() {
  dbID("idDia_Error").close();
}

function populateTokenList(parentID, list) {
  let ulParent = dbID(parentID);
  for (let token of list) {
    const li = document.createElement("li");
    li.append(token);
    ulParent.append(li);
  }
}

function getMainNumber(event) {
  fileData.outputMainNumber = event.target.value;
  if (fileData.outputMainNumber == "") fileData.outputMainNumber = null;
  updateMainName();
}

function getMainName(event) {
  fileData.outputMainName = event.target.value;
  if (fileData.outputMainName == "") fileData.outputMainName = null;
  updateMainName();
}

function updateMainName() {
  if (fileData.outputMainNumber === null || fileData.outputMainNumber.toString().length != 6) {
    enableDownload();
    Lbl_fileName.KadReset();
    return;
  }
  if (fileData.outputMainName == null) {
    fileData.outputName = `${mainNumberPadded(fileData.outputMainNumber)}.xlsx`;
  } else {
    fileData.outputName = `${mainNumberPadded(fileData.outputMainNumber)}_${fileData.outputMainName}.xlsx`;
  }
  enableDownload();
  Lbl_fileName.KadSetText(fileData.outputName);
}

function mainNumberPadded(num = null) {
  if (num == null) return "";
  return num.toString().padStart(6, "0");
}

const fileData = {
  rawStringAvailableObj: {
    Menge: null,
    Struktur: null,
  },
  get rawStringAvailable() {
    return this.rawStringAvailableObj.Menge && this.rawStringAvailableObj.Struktur;
  },
  rawStringData: {
    Menge: null,
    Struktur: null,
  },
  rawData: {
    Menge: null,
    Struktur: null,
  },
  listData: [],
  outputName: "",
  outputMainNumber: null,
};

function readData() {
  fileData.rawStringData.Menge = Area_inputMenge.KadGet({ noPlaceholder: true });
  fileData.rawStringAvailableObj.Menge = fileData.rawStringData.Menge != "";
  fileData.rawStringData.Struktur = Area_inputStruktur.KadGet({ noPlaceholder: true });
  fileData.rawStringAvailableObj.Struktur = fileData.rawStringData.Struktur != "";
  if (fileData.rawStringAvailable) {
    parseStringData("Menge");
    parseStringData("Struktur");
  }
  enableDownload();
}

const mmID = "ArtikelNr";
const name = "Bezeichnung";
const count = "Menge";
const sparePart = "Ersatzteil";
const wearPart = "Verschleissteil";
const partFamily = "ArtikelTeileFamilie";
const partDataFields = [mmID, name, count, sparePart, wearPart, partFamily];

const dataObject = {
  partData: {},
  listData: [],
  evArray: [],
  missingIDs: {},
};

function parseStringData(type) {
  fileData.rawData[type] = [];
  let rows = fileData.rawStringData[type].split("\n");
  let headerFields = rows.splice(0, 1)[0].split("\t");
  for (let i = headerFields.length - 1; i > 0; i--) {
    if (headerFields[i] === "") {
      headerFields.splice(i, 1);
    } else {
      break;
    }
  }

  for (let row of rows) {
    const data = row.split("\t");
    if (data[0] === "") continue;
    let obj = {};
    for (let i = 0; i < headerFields.length; i++) {
      if (data[i] == "") {
        obj[headerFields[i]] = "";
      } else if (isNaN(Number(data[i]))) {
        obj[headerFields[i]] = data[i];
      } else {
        obj[headerFields[i]] = Number(data[i]);
      }
    }
    fileData.rawData[type].push(obj);
  }
}

function enableDownload(enable = null) {
  if (enable === false) {
    Btn_download.KadEnable(false);
    return;
  }
  let state = fileData.rawStringAvailable && fileData.outputName != "" ? true : false;
  Btn_download.KadEnable(state);

  if (state) parseFile();
}

function startDownload() {
  const book = utils.book_new();
  dataObject.listData.unshift(["Zeichnung", name, "EV-Nummern"]);
  const listData = utils.aoa_to_sheet(dataObject.listData);
  utils.book_append_sheet(book, listData, "Baugruppen");
  writeFile(book, fileData.outputName);
}

function findParentAndAddAsChild(i, childID, startLevel) {
  let tempLevel = startLevel;
  let tempID = childID;
  for (let p = i - 1; p >= 0; p--) {
    const higherID = Number(fileData.rawData.Struktur[p][mmID]);
    const higherLevel = Number(fileData.rawData.Struktur[p].Ebene);
    if (tempLevel <= higherLevel) continue;
    const higherObj = dataObject.partData[higherID];
    if (!higherObj.children.includes(tempID)) {
      higherObj.children.push(tempID);
    }
    tempID = higherID;
    tempLevel--;
  }
}

// main calculating function!!!!
function parseFile() {
  Lbl_missingIDs.KadReset();
  dataObject.partData = {};
  dataObject.listData = [];
  dataObject.evArray = [];
  dataObject.missingIDs = {};

  //inject "Header"
  dataObject.partData[fileData.outputMainNumber] = {
    [mmID]: fileData.outputMainNumber,
    [name]: fileData.outputMainName,
    [sparePart]: false,
    [wearPart]: false,
    [partFamily]: "",
    children: [],
    level: 0,
  };
  fileData.rawData.Struktur.unshift({
    ArtikelArt: "F",
    ArtikelNr: fileData.outputMainNumber,
    Baustein: "D",
    Bezeichnung: fileData.outputMainName,

    Ebene: "0",
    Einheit: "Stk",
    Gesperrt: "false",
    Matchcode: "",
    Menge: "1,00",
    PosNr: "10",
  });
  for (let obj of fileData.rawData.Menge) {
    let id = Number(obj[mmID]); // get MM-Nummer

    dataObject.partData[id] = {};
    for (let field of partDataFields) {
      dataObject.partData[id][field] = obj[field];
    }
    dataObject.partData[id]["children"] = [];

    if (obj[sparePart] == "true" || obj[wearPart] == "true") {
      dataObject.evArray.push(id);
    }
  }
  Lbl_loadedSOU.KadSetText(dataObject.evArray.length);

  for (let i = 0; i < fileData.rawData.Struktur.length; i++) {
    const currObj = fileData.rawData.Struktur[i];
    const id = Number(currObj[mmID]);
    if (dataObject.partData[id] == undefined) {
      const newID = !dataObject.missingIDs.hasOwnProperty(id);
      if (newID) {
        dataObject.missingIDs[id] = [currObj];
      } else {
        dataObject.missingIDs[id].push(currObj);
      }
      Lbl_missingIDs.KadSetHTML(`${id.toString().padStart(6, 0)}<br>`);
      continue;
    }

    const level = Number(currObj.Ebene);
    dataObject.partData[id].level = level;

    // find all parents
    if (dataObject.evArray.includes(Number(currObj[mmID]))) {
      KadLog.log(currObj, mmID);
      findParentAndAddAsChild(i, id, level);
    }
  }

  dataObject.listData = [];
  for (let i = 0; i < fileData.rawData.Struktur.length; i++) {
    const currObj = fileData.rawData.Struktur[i];
    const id = Number(currObj[mmID]);
    if (dataObject.partData[id].children.length == 0) continue;
    if (dataObject.listData.some((arr) => arr[0] == id)) continue;

    dataObject.listData.push([id, dataObject.partData[id][name], ...dataObject.partData[id].children]);
  }
}
