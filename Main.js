import * as KadUtils from "./Data/KadUtils.js";
import { read, utils, writeFile } from "./Data/xlsx.mjs";

window.onload = mainSetup;

function mainSetup() {
  KadUtils.KadDOM.resetInput(idVin_mainAssemblyNr, "MM-Nummern Anlage");
  KadUtils.KadDOM.resetInput(idVin_mainAssemblyName, "Anlagename");
  KadUtils.daEL(idVin_mainAssemblyNr, "input", getMainNumber);
  KadUtils.daEL(idVin_mainAssemblyName, "input", getMainName);

  KadUtils.dbID(idLbl_loadedSOU).textContent = "nicht geladen";
  KadUtils.daEL(idVin_inputSOU, "change", (evt) => getFile(evt));
  KadUtils.KadDOM.enableBtn(idVin_inputSOU, false);
  KadUtils.KadDOM.enableBtn(idLbl_inputSOU, false);

  KadUtils.daEL(idBtn_infoUpload, "click", openInfoUpload);
  KadUtils.daEL(idBtn_infoCloseUpload, "click", closeInfoUpload);
  KadUtils.daEL(idBtn_infoError, "click", openInfoError);
  KadUtils.daEL(idBtn_infoCloseError, "click", closeInfoError);
  KadUtils.daEL(idBtn_infoSOU, "click", openInfoSOU);
  KadUtils.daEL(idBtn_infoCloseSOU, "click", closeInfoSOU);

  KadUtils.daEL(idBtn_download, "click", startDownload);
  KadUtils.KadDOM.enableBtn(idBtn_download, false);
  KadUtils.dbID(idLbl_fileName).textContent = `*.xlsx`;

  populateTokenList(idUL_Upload, ulInfoUpload);
  populateTokenList(idUL_Error, ulInfoError);
  populateTokenList(idUL_SOU, ulInfoSOU);
}

const ulInfoUpload = ['Tabellenblätter müssen mit "Struktur" und "Menge" beschriftet sein!', "MM-Nummer der Hauptanlage im linken Feld eintragen", "Der Anlagenname im rechten Feld ist optional.", 'Der Button "SOU-Liste als *.xlsx" ist blockiert wenn keine MM-Nummer der Anlage eingegeben wurde'];
const ulInfoError = ['MM-Numern von "Struktur" nicht in "Menge" enthalten!', 'Liste von MM-Nummern, die in der "Struktur"-Tabelle enthalten sind aber nicht in der "Mengen"-Tabelle.', "Durch diesen Fehler wird keine Excel-Datei ausgegeben!"];
const ulInfoSOU = ['Mengenstückliste und Strukturstückliste mit "Zwischenablage (Daten)" auf 2 Tabellenblätter in einer Datei speichern.', 'Ist gesperrt, wenn keine MM-Nummer im Feld "MM-Nummer Anlage" eingetragen ist.'];

function openInfoUpload() {
  KadUtils.dbID(idDia_Upload).showModal();
}
function closeInfoUpload() {
  KadUtils.dbID(idDia_Upload).close();
}
function openInfoError() {
  KadUtils.dbID(idDia_Error).showModal();
}
function closeInfoError() {
  KadUtils.dbID(idDia_Error).close();
}
function openInfoSOU() {
  KadUtils.dbID(idDia_SOU).showModal();
}
function closeInfoSOU() {
  KadUtils.dbID(idDia_SOU).close();
}

function populateTokenList(parentID, list) {
  let ulParent = KadUtils.dbID(parentID);
  for (let token of list) {
    const li = document.createElement("li");
    li.append(token);
    ulParent.append(li);
  }
}

let mainNumber = "";
function getMainNumber(event) {
  let results = event.target.value;
  mainNumber = results.match(/\d{6}/g);
  if (mainNumber == null) {
    return;
  }
  mainNumber = Number(mainNumber);
  KadUtils.KadDOM.enableBtn(idVin_inputSOU, true);
  KadUtils.KadDOM.enableBtn(idLbl_inputSOU, true);
  KadUtils.dbID(idLbl_fileName).textContent = `${mainNumberPadded(mainNumber)}_${mainName}.xlsx`;
}
function mainNumberPadded(num) {
  return num.toString().padStart(6, "0");
}

let mainName = "";
function getMainName(event) {
  mainName = event.target.value;
  KadUtils.dbID(idLbl_fileName).textContent = `${mainNumberPadded(mainNumber)}_${mainName}.xlsx`;
}

const fileData = {
  rawData: {
    Menge: null,
    Struktur: null,
  },
  outputName: "",
};

function getFile(file) {
  fileData.rawData = {};

  let selectedFile = file.target.files[0];
  let fileReader = new FileReader();

  fileReader.onload = (event) => {
    const data = event.target.result;
    let workbook = read(data, { type: "binary" });
    let error = 2;
    workbook.SheetNames.forEach((sheet) => {
      const title = KadUtils.KadString.firstLetterCap(sheet);
      if (title == "Struktur" || title == "Menge") {
        error--;
        const data = utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
        fileData.rawData[title] = data;
      }
    });
    if (error < 0) {
      alert(ulInfoUpload[0]);
      return;
    }
    fileIsParsed();
    parseFile();
  };
  fileReader.readAsBinaryString(selectedFile);
}

function fileIsParsed() {
  KadUtils.KadDOM.btnColor(idLbl_inputSOU, "positive");
  setTimeout(() => {
    KadUtils.KadDOM.btnColor(idLbl_inputSOU, null);
  }, 3000);
}

// -----------------------------

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
};

let missingIDs = {};

function parseFile() {
  dataObject.evArray = [];
  dataObject.partData = {};
  missingIDs = {};
  idLbl_missingIDs.innerHTML = "";
  //inject "MAIN"

  dataObject.partData[mainNumber] = {
    [mmID]: mainNumber,
    [name]: mainName,
    [count]: 1,
    [sparePart]: false,
    [wearPart]: false,
    [partFamily]: "",
    children: [],
    level: 0,
  };
  fileData.rawData.Struktur.unshift({
    ArtikelArt: "F",
    ArtikelNr: mainNumber,
    Baustein: "D",
    Bezeichnung: mainName,
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
  KadUtils.dbID(idLbl_loadedSOU).textContent = `${dataObject.evArray.length} E/V-Teile gefunden`;

  // console.log(dataObject.partData["153707"]);
  for (let i = 0; i < fileData.rawData.Struktur.length; i++) {
    const currObj = fileData.rawData.Struktur[i];
    const id = Number(currObj[mmID]);
    if (dataObject.partData[id] == undefined) {
      const newID = !missingIDs.hasOwnProperty(id);
      if (newID) {
        missingIDs[id] = [currObj];
      } else {
        missingIDs[id].push(currObj);
      }
      idLbl_missingIDs.innerHTML += `${id.toString().padStart(6, 0)}<br>`;
      continue;
    }

    const level = Number(currObj.Ebene);
    dataObject.partData[id].level = level;

    // find all parents
    if (dataObject.evArray.includes(Number(currObj[mmID]))) {
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
  KadUtils.KadDOM.enableBtn(idBtn_download, true);
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

function startDownload() {
  const book = utils.book_new();
  dataObject.listData.unshift(["Zeichnung", name, "EV-Nummern"]);
  const listData = utils.aoa_to_sheet(dataObject.listData);
  utils.book_append_sheet(book, listData, "Baugruppen");
  writeFile(book, `${mainNumberPadded(mainNumber)}_${mainName}_EV.xlsx`);
}
