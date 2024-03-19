import * as KadUtils from "./Data/KadUtils.js";
import { utils, read, writeFile } from "./Data/xlsx.mjs";

window.onload = mainSetup;

function mainSetup() {
	KadUtils.dbID(idLbl_loadedSOU).textContent = "nicht geladen";
	KadUtils.daEL(idVin_inputSOU, "change", (evt) => getFile(evt));
	KadUtils.daEL(idBtn_infoSOU, "click", openInfoSOU);
	KadUtils.daEL(idBtn_infoCloseSOU, "click", closeInfoSOU);

	KadUtils.daEL(idBtn_download, "click", startDownload);
	KadUtils.KadDOM.enableBtn(idBtn_download, false);
	KadUtils.dbID(idLbl_fileName).textContent = `*.xlsx`;

	populateTokenList(idUL_SOU, ulInfoSOU);
}

const ulInfoSOU = ["Hallo Marina, Infos kommen noch"];

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

const fileData = {
	rawData: {},
	outputName: "",
};

function getFile(file) {
	fileData.rawData = {};

	let selectedFile = file.target.files[0];
	let fileReader = new FileReader();

	fileReader.onload = (event) => {
		fileData.outputName = `${file.target.files[0].name.split(".")[0]}_EV`;
		KadUtils.dbID(idLbl_fileName).textContent = `${fileData.outputName}.xlsx`;

		const data = event.target.result;
		let workbook = read(data, { type: "binary" });
		workbook.SheetNames.forEach((sheet) => {
			const data = utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
			let title = Object.keys(data[0]).includes("Ebene") ? "Struktur" : "Menge";
			fileData.rawData[title] = data;
		});
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
const partDataFields = [mmID, name, count, sparePart, partFamily, wearPart];

const dataObject = {
	partData: {},
	listData: {},
	topDownList: {},
};

function parseFile() {
	dataObject.partData = {};
	for (let obj of fileData.rawData.Menge) {
		if (obj[sparePart] == "true" || obj[wearPart] == "true") {
			let id = Number(obj[mmID]); // get MM-Nummer
			dataObject.partData[id] = {};
			for (let field of partDataFields) {
				dataObject.partData[id][field] = obj[field];
			}
			dataObject.partData[id]["parents"] = {};
		}
	}

	KadUtils.dbID(idLbl_loadedSOU).textContent = `${KadUtils.objectLength(dataObject.partData)} Teile gefunden`;

	dataObject.topDownList = {};
	for (let i = 0; i < fileData.rawData.Struktur.length; i++) {
		const obj = fileData.rawData.Struktur[i];

		if (dataObject.partData.hasOwnProperty(Number(obj[mmID]))) {
			const id = Number(obj[mmID]);
			dataObject.partData[id].level = Number(obj.Ebene);

			// find all parents
			let currentLevel = Number(obj.Ebene);
			for (let p = i - 1; p > 0; p--) {
				const prevLevel = Number(fileData.rawData.Struktur[p].Ebene);
				if (prevLevel < currentLevel) {
					currentLevel = prevLevel;
					const parentID = Number(fileData.rawData.Struktur[p][mmID]);
					if (!dataObject.partData[id].parents.hasOwnProperty(parentID)) {
						dataObject.partData[id].parents[parentID] = currentLevel;
						if (!dataObject.topDownList[currentLevel]) dataObject.topDownList[currentLevel] = new Set();
						dataObject.topDownList[currentLevel].add(parentID);
					}
				}
			}
		}
	}
	// console.log(KadUtils.objectLength(dataObject.partData), dataObject.partData);

	for (let key of Object.keys(dataObject.topDownList)) {
		dataObject.topDownList[key] = Array.from(dataObject.topDownList[key]);
	}
	console.log(dataObject.topDownList);

	KadUtils.KadDOM.enableBtn(idBtn_download, true);
}

// -----------------------------
function startDownload() {
	const book = utils.book_new();
	const sheetDiff = utils.json_to_sheet(Object.values(dataObject.difference));

	utils.book_append_sheet(book, sheetDiff, "Struktur");

	writeFile(book, `${fileData.outputName}.xlsx`);
}
