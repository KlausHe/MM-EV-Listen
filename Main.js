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

const ulInfoSOU = ["Stückliste der Baugruppe aufrufen", "Reporte -> Mengenstückliste", "In Zwischenablage speichern", "Neue Excel-Datei öffnen", "Zwischenablage in Zelle A1 kopieren", "Mit beliebigem Namen speichern", "Der Dateiname wird für die Vergleichsdatei verwendet"];

const fileData = {
	rawData: {},
	outputName: "",
};

const dataObject = {
	partData: {},
	listData: {},
};

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
		console.log(fileData.rawData);
		fileIsParsed();
		parseFileExcel();
	};
	fileReader.readAsBinaryString(selectedFile);
}

const mmID = "ArtikelNr";
const name = "Bezeichnung";
const count = "Menge";
const sparePart = "Ersatzteil";
const wearPart = "Verschleissteil";
const partFamily = "ArtikelTeileFamilie";
const partDataFields = [mmID, name, count, sparePart, wearPart, "parents", "children"];

function parseFileExcel() {
	dataObject.partData = {};

	for (let obj of fileData.rawData.Menge) {
		if (obj[sparePart] == "true" || obj[wearPart] == "true") {
			let id = Number(obj[mmID]); // get MM-Nummer
			dataObject.partData[id] = {};
			for (let field of partDataFields) {
				dataObject.partData[id][field] = obj[field];
			}
		}
	}
	KadUtils.dbID(idLbl_loadedSOU).textContent = `${KadUtils.objectLength(dataObject.partData)} Teile gefunden`;

	console.log(dataObject.partData);
}

function fileIsParsed() {
	KadUtils.KadDOM.btnColor(idLbl_inputSOU, "positive");
	setTimeout(() => {
		KadUtils.KadDOM.btnColor(idLbl_inputSOU, null);
	}, 3000);
}

// -----------------------------
function startCompare() {
	dataObject.compared = {};
	dataObject.notInSOU = {};
	dataObject.notInCAD = {};
	dataObject.tokenlist = {};

	for (let token of [...Tokenlist[0], ...Tokenlist[1]]) {
		let found = false;
		for (let [souKey, souValue] of Object.entries(dataObject.SOU)) {
			if (souValue[name].toLowerCase().includes(token.toLowerCase())) {
				found = true;
				dataObject.tokenlist[souKey] = {
					[mmID]: souValue[mmID],
					SOU: souValue[count],
					[name]: souValue[name],
					[partFamily]: souValue[partFamily] ? souValue[partFamily] : "---",
				};
			}
		}
		if (!found) {
			dataObject.tokenlist[token] = {
				[mmID]: token,
				SOU: 0,
				[name]: "nicht vorhanden",
				[partFamily]: "",
			};
		}
	}

	for (let [souKey, souValue] of Object.entries(dataObject.SOU)) {
		const cadCount = dataObject.CAD[souKey] == null ? 0 : dataObject.CAD[souKey][count];
		dataObject.compared[souKey] = {
			[mmID]: souValue[mmID],
			SOU: souValue[count],
			CAD: cadCount,
			[name]: souValue[name],
			[partFamily]: souValue[partFamily] ? souValue[partFamily] : "---",
			[foundInSOU]: true,
			[foundInCAD]: cadCount == 0 ? false : true,
		};
		if (cadCount == 0) dataObject.notInCAD[souKey] = dataObject.compared[souKey];
	}

	for (let [cadKey, cadValue] of Object.entries(dataObject.CAD)) {
		if (dataObject.compared[cadKey] === undefined) {
			dataObject.compared[cadKey] = {
				[mmID]: cadValue[mmID],
				SOU: 0,
				CAD: cadValue[count],
				[name]: cadValue[name],
				[partFamily]: cadValue[partFamily],
				[foundInSOU]: false,
				[foundInCAD]: true,
			};
			dataObject.notInSOU[cadKey] = dataObject.compared[cadKey];
		}
	}

	KadUtils.KadDOM.enableBtn(idBtn_download, true);
}

function startDownload() {
	const book = utils.book_new();
	const sheetDiff = utils.json_to_sheet(Object.values(dataObject.difference));

	utils.book_append_sheet(book, sheetDiff, "Struktur");

	writeFile(book, `${fileData.outputName}.xlsx`);
}
