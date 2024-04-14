export function dbID(id) {
	if (id instanceof Object) return id;
	return document.getElementById(id);
}
export function dbIDStyle(id) {
	if (id instanceof Object) return id.style;
	return document.getElementById(id).style;
}
export function dbCL(id, loc = 0) {
	if (loc === null) return document.getElementsByClassName(id);
	return document.getElementsByClassName(id)[loc];
}
export function dbCLStyle(id, loc = 0) {
	if (loc === null) return [...document.getElementsByClassName(id)].map((s) => s.style);
	return document.getElementsByClassName(id)[loc].style;
}
export function daEL(id, type, fn) {
	dbID(id).addEventListener(type, fn);
}
export function objectLength(obj) {
	return Object.keys(obj).length;
}
export function hostDebug() {
	return ["local", "127.0.0.1"].some((s) => window.location.hostname.includes(s));
}
export function error(...errorText) {
	throw new Error(errorText.join("; "));
}

export function deepClone(data) {
	if (data === null || data === undefined) return data;
	return JSON.parse(JSON.stringify(data));
}
export function copyToClipboard(text, enabled = true) {
	if (!enabled) return;
	if (!navigator.clipboard) return;
	let val = text;
	if (!isNaN(val) && Number.isFinite(Number(val))) {
		val = val.toString().replace(/,/g, ""); //remove thousandscomma
		val = val.replace(".", ",");
	}
  navigator.clipboard.writeText(val);
}
export const KadCSS = {
	getRoot(object, numberOnly = false, RemToPX = false) {
		//  getCssRoot("navbarHeight", return only numberOnly=true)
		const obj = `--${object}`;
		const valOrig = getComputedStyle(document.body).getPropertyValue(obj);
		const unit = valOrig.match(/[a-zA-Z]{1,}/g);
		if (RemToPX == true && unit == "rem") {
			const size = getComputedStyle(document.body).getPropertyValue("--fontSize").replace(/px/g, "");
			const valConverted = valOrig.replace(/rem/g, "");
			return Number(size * valConverted);
		}
		if (numberOnly == false) return getComputedStyle(document.body).getPropertyValue(obj).trim();
		const valConverted = valOrig.replace(/s|px|rem/g, "");
		return Number(valConverted);
	},
	setRoot(object, value, dim = "") {
		document.styleSheets[0].cssRules[0].style.setProperty(`--${object}`, `${value}${dim}`);
	},
};
export const KadDOM = {
	scrollInView(id) {
		dbID(id).scrollIntoView({
			behavior: "smooth",
			block: "nearest",
			inline: "end", //start
		});
	},
	getImgPath(name) {
		return `Data/Images/SVG/${name}.svg`;
	},
	resetInput(id, ph = null, domOpts = null) {
		const obj = dbID(id);
		if (obj.type == "checkbox") {
			obj.checked = ph;
			return ph;
		}
		obj.value = "";
		if (ph != null) obj.placeholder = ph;
		if (domOpts != null) {
			for (let [key, val] of Object.entries(domOpts)) {
				obj[key] = val;
			}
		}
		return Number(obj.placeholder);
	},
	enableBtn(id, state) {
		const obj = typeof id == "string" ? dbID(id) : id;
		if (state) obj.removeAttribute("disabled");
		else obj.setAttribute("disabled", "true");
	},
	btnColor(id, opt = null) {
		const obj = dbID(id);
		if (opt === null) obj.removeAttribute("data-btnstatus");
		else if (opt === "positive") obj.dataset.btnstatus = "btnPositive";
		else if (opt === "negative") obj.dataset.btnstatus = "btnNegative";
		else if (opt === "colored") obj.dataset.btnstatus = "btnBasecolor";
	},
	vinChange(id, v) {
		let obj = null;
		let siblingList = Array.from(id.parentNode.children);
		for (let i = siblingList.indexOf(id) - 1; i >= 0; i--) {
			if (siblingList[i].type != "button") {
				obj = siblingList[i];
				break;
			}
		}
		if (obj == null) return;
		if (obj.disabled) return;
		const dir = Number(v);
		if (obj.type == "time") evaluateTime();
		if (obj.type == "number") evaluateNumber();
		obj.dispatchEvent(new Event("input"));
		obj.focus();
		function evaluateTime() {
			const h = Number(obj.value.slice(0, 2));
			const m = Number(obj.value.slice(3, 5));
			let time = m + h * 60;
			time += time % 5 == 0 ? dir * 5 : dir;
			const t = KadDate.minutesToObj(time);
			obj.value = `${t.h}:${t.m}`;
		}
		function evaluateNumber() {
			if (dir == 0) {
				const time = new Date().getTime();
				obj.setAttribute("data-ts", time);
				if (Number(obj.value) === 0 || Number(obj.value) === Number(obj.min)) {
					obj.value = "";
					return;
				}
				obj.value = obj.min || 0;
				return;
			}
			const time = new Date().getTime();
			let skip = false;
			if (obj.hasAttribute("data-ts")) {
				if (time - obj.dataset.ts < 1500) skip = true;
			}
			obj.setAttribute("data-ts", time);
			const actual = obj.value == "" && obj.placeholder != "" ? Number(obj.placeholder) : Number(obj.value);
			const num = skip && actual % 5 == 0 ? actual + dir * 5 : actual + dir;
			const min = obj.hasAttribute("min") && dir < 1 ? Number(obj.min) : null;
			const max = obj.hasAttribute("max") && dir > 0 ? Number(obj.max) : null;
			obj.value = KadValue.constrain(num, min, max);
		}
	},
	numberFromInput(id, failSafeVal = null, noPlaceholder = null) {
		const obj = dbID(id);
		if (!isNaN(obj.valueAsNumber)) return obj.valueAsNumber;
		if (failSafeVal != null) return failSafeVal;
		if (noPlaceholder != null) return null;
		return Number(obj.placeholder);
	},
	stringFromInput(id, failSafeVal = null, noPlaceholder = null) {
		const obj = dbID(id);
		const value = obj.value.trim();
		if (value != "") return obj.value;
		if (failSafeVal != null) return failSafeVal;
		if (noPlaceholder != null) return null;
		return obj.placeholder;
	},
	clearFirstChild(id) {
		const obj = typeof id == "string" ? dbID(id) : id;
		while (obj.firstChild) {
			obj.removeChild(obj.firstChild);
		}
		return obj;
	},
};
export const KadInteraction = {
	focus(obj, canv = null) {
		if (canv != null) canv.loop();
		dbID(obj).focus();
	},
	unfocus(obj, canv) {
		dbID(obj).blur();
		if (canv != null) {
			canv.noLoop();
			canv.redraw();
		}
	},
};
export const KadImage = {
	getFavicon(url, size = 15) {
		return `https://t2.gstatic.com/faviconV2?client=SOCIAL&type=FAVICON&fallback_opts=TYPE,SIZE,URL&size=${size}&url=${url}`; //alternative: `https://www.google.com/s2/favicons?domain=${url}&sz=${size}`;
	},
};
export const KadValue = {
	number(value = 1, { form = null, indicator = false, leadingDigits = 1, decimals = 1, currency = null, unit = null, notation = "standard" } = {}) {
		const formating = form == null ? "de-DE" : "," ? "de-DE" : "en-EN";
		let options = {
			useGrouping: indicator,
			notation: notation,
			minimumIntegerDigits: leadingDigits,
			maximumFractionDigits: decimals,
		};
		if (options.notation == "engineering" || options.notation == "scientific") {
			if (value < 1000 && value > -1000) {
				options.notation = "standard";
				options.maximumFractionDigits = 1;
			}
		}
		if (currency) {
			options.useGrouping = true;
			options.style = "currency";
			options.currency = currency;
			options.maximumFractionDigits = 3;
		}
		if (unit) {
			options.style = "unit";
			options.unit = unit;
			options.unitDisplay = "short";
			options.useGrouping = true;
		}
		return Intl.NumberFormat(formating, options).format(value);
	},
	constrain(val, min = null, max = null) {
		if (min == null && max == null) return val;
		if (min != null && max != null) return Math.max(Math.min(val, max), min);
		if (min == null && max != null) return Math.min(val, max);
		if (min != null && max == null) return Math.max(val, min);
	},
	mapping(i, start1, stop1, start2, stop2, bounds = false) {
		const val = ((i - start1) / (stop1 - start1)) * (stop2 - start2) + start2;
		if (!bounds) return val;
		const up = Math.max(start2, stop2);
		const down = Math.min(start2, stop2);
		return KadValue.constrain(val, down, up);
	},
	// untested!!!
	constrainArray(arr, val, low = null, high = null) {
		const arrayMin = Math.min(...arr);
		const arrayMax = Math.max(...arr);
		let a = low || arrayMin;
		let b = high || arrayMax;
		constrain(val, a, b);
	},
};
export const KadArray = {
	createArray(x, y = null, fillNum = null) {
		if (y == null) return new Array(x).fill(0).map((n, i) => i);
		let arrX = new Array(x);
		for (let i = 0; i < arrX.length; i++) {
			arrX[i] = fillNum == null ? (arrX[i] = new Array(y)) : new Array(y).fill(fillNum);
		}
		return arrX;
	},
	createIndexedArray(x, y = null) {
		if (y == null) {
			return new Array(x).fill(0).map((n, i) => i);
		}
		let arrXY = [];
		for (let i = 0; i < x; i++) {
			for (let j = 0; j < y; j++) {
				arrXY.push([i, j]);
			}
		}
		return arrXY;
	},
	indexTo2DxyPosition(index, col) {
		const i = Math.floor(index / col);
		const j = index % col;
		return { i, j };
	},
	arrayFromNumber(obj, num = null) {
		if (num == null && typeof obj == "number") return [...Array(obj).keys()];
		if (typeof obj == "number" && typeof num == "number") {
			const min = Math.min(obj, num);
			const max = Math.max(obj, num);
			let arr = [];
			for (let i = min; i <= max; i++) {
				arr.push(i);
			}
			return arr;
		}
	},
	getNearestValueInArray(arr, val, higher = false) {
		return arr.reduce((prev, curr) => {
			if (higher) return Math.abs(curr - val) > Math.abs(prev - val) ? curr : prev;
			else return Math.abs(curr - val) < Math.abs(prev - val) ? curr : prev;
		});
	},
	sortArrayByKey(arr, key, inverse = false, caseSensitive = false) {
		let array = Array.from(arr);
		return array.sort((a, b) => {
			if (typeof a[key] == "number" && typeof b[key] == "number") {
				return inverse ? b[key] - a[key] : a[key] - b[key];
			} else {
				const x = caseSensitive ? a[key].toLowerCase() : a[key];
				const y = caseSensitive ? b[key].toLowerCase() : b[key];
				const dir = inverse ? 1 : -1;
				return x < y ? dir : x > y ? dir : 0;
			}
		});
	},
};
export const KadRandom = {
	randomIndex(obj) {
		if (typeof obj == "string") return Math.floor(Math.random() * obj.length);
		if (Array.isArray(obj)) return Math.floor(Math.random() * obj.length);
		return Math.floor(Math.random() * Object.keys(obj).length);
	},
	randomObject(obj, top = null) {
		// takes a single Number, an Array or an Object
		if (typeof obj == "number") return KadRandom.randomObject(KadArray.arrayFromNumber(obj, top));
		if (typeof obj == "string") return obj[KadRandom.randomIndex(obj)];
		if (Array.isArray(obj) && obj.length <= 0) return null;
		if (Array.isArray(obj)) return obj[KadRandom.randomIndex(obj)];
		const objKeys = Object.keys(obj);
		return obj[objKeys[KadRandom.randomIndex(objKeys)]];
	},
	randomObjectCentered(obj, top = null, iterations = 1) {
		let sum = 0;
		for (let i = 0; i < iterations; i++) {
			sum += Random.randomObject(obj, top);
		}
		return Math.floor(sum / iterations);
	},
	randomSubset(array, numSubset) {
		let options = KadRandom.shuffleData(array);
		return options.slice(0, numSubset);
	},
	boolProbablity(threshold = 0.5) {
		return Math.random() < threshold ? 1 : 0;
	},
	shuffleData(arr) {
		let shuffled = deepClone(arr);
		if (!Array.isArray(arr)) return arr;
		for (let i = shuffled.length - 1; i > 0; i--) {
			const j = Math.floor(Math.random() * (i + 1));
			[shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
		}
		return shuffled;
	},
};
export const KadDate = {
	getDate(date = null, { format = "DD.MM.YYYY", leadingDigit = true, reversed = false } = {}) {
		const regexSplit = new RegExp(/([$-/:-?{-~!"^_`\ [\]])/);
		const conversions = {
			date: date === null ? new Date() : new Date(date),
			get YYYY() {
				return this.date.getFullYear();
			},
			get YY() {
				return parseInt(this.date.getFullYear().toString().slice(2, 4), 10);
			},
			get MM() {
				return this.date.getMonth() + 1;
			},
			get DD() {
				return this.date.getDate();
			},
			get WD() {
				const weekdaysShort = ["So.", "Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa."];
				return weekdaysShort[this.date.getDay()];
			},
			get HH() {
				return this.date.getHours();
			},
			get mm() {
				return this.date.getMinutes();
			},
			get ss() {
				return this.date.getSeconds();
			},
			get ms() {
				return this.date.getMilliseconds();
			},
		};
		let arr = format.split(regexSplit);
		const leadingDigits = leadingDigit ? 2 : 1;
		for (let i = 0; i < arr.length; i++) {
			if (Object.keys(conversions).includes(arr[i])) {
				let str = KadValue.number(conversions[arr[i]], { leadingDigits });
				arr[i] = !isNaN(str) ? str : conversions[arr[i]];
			}
		}
		if (reversed) arr = arr.reverse();
		return arr.join("");
	},
	minutesToObj(mins) {
		const h = Math.floor(mins / 60) < 10 ? `0${Math.floor(mins / 60)}` : Math.floor(mins / 60);
		const m = mins % 60 < 10 ? `0${mins % 60}` : mins % 60;
		return { h, m };
	},
	secondsToObj(second) {
		const secs = Math.floor(second % 60);
		const mins = Math.floor(second / 60);
		const hours = Math.floor(mins / 60);
		const h = hours < 10 ? `0${hours}` : hours;
		const m = mins < 10 ? `0${mins}` : mins;
		const s = secs < 10 ? `0${secs}` : secs;
		return { h, m, s };
	},
	hourAsNumber(time = null) {
		if (time === null) {
			return new Date().getHours();
		} else if (time < 10000000000) {
			return new Date(time * 1000).getHours();
		} else {
			return new Date(time).getHours();
		}
	},
	getWeekNumber(d) {
		const date = d ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : new Date();
		const curThu = new Date(date.getTime() + (3 - ((date.getDay() + 6) % 7)) * 86400000);
		const yearThu = curThu.getFullYear();
		const firstThu = new Date(new Date(yearThu, 0, 4).getTime() + (3 - ((new Date(yearThu, 0, 4).getDay() + 6) % 7)) * 86400000);
		return Math.floor(1 + 0.5 + (curThu.getTime() - firstThu.getTime()) / 604800000);
	},
};
export const KadString = {
	firstLetterCap(s) {
		if (s == "") return s;
		if (typeof s != "string") return s;
		return s[0].toUpperCase() + s.slice(1).toLowerCase();
	},
	firstLetterLow(s) {
		if (s == "") return s;
		if (typeof s != "string") return s;
		return s[0].toLowerCase() + s.slice(1);
	},
};
export const KadTable = {
	clear(id) {
		const obj = dbID(id);
		for (let i = obj.rows.length - 1; i >= 0; i--) {
			obj.deleteRow(i);
		}
	},
	insertRow(tabID) {
		const obj = dbID(tabID);
		return obj.insertRow(obj.rows.length);
	},
	addHeaderCell(row, opt) {
		opt.name = this.createName(opt);
		let cell = document.createElement("th");
		cell.id = `id${opt.type}${opt.name}`;
		const mainChild = this.createCell(opt.type, opt);
		cell.appendChild(mainChild);
		this.CellOptions(cell, opt);
		row.appendChild(cell);
		return cell;
	},
	addCell(row, opt = {}, prevCell = null) {
		opt.name = this.createName(opt);
		const mainChild = this.createCell(opt.type, opt);
		let cell = undefined;
		if (prevCell === null) {
			cell = row.insertCell(-1);
			cell.id = `id${opt.type}_cell${opt.name}`;
		}
		if (mainChild != null) {
			if (prevCell === null) {
				cell.appendChild(mainChild);
			} else {
				prevCell.appendChild(mainChild);
			}
		}
		this.CellOptions(cell || prevCell, opt);
		return cell || prevCell;
	},
	createCell(type, opt) {
		return KadTable.cells[type](opt);
	},
	cells: {
		Vin(opt) {
			const child = document.createElement("INPUT");
			if (opt.subGroup == "checkbox") {
				child.type = "checkbox";
				child.checked = opt.checked || false;
			} else {
				child.type = opt.subGroup;
				child.placeholder = opt.placeholder;
			}
			this.UIOptions(child, opt);
			return child;
		},
		Btn(opt) {
			const child = document.createElement("BUTTON");
			const elImg = document.createElement("img");
			child.appendChild(elImg);
			child.type = "button";
			switch (opt.subGroup) {
				case "button":
					elImg.src = KadDOM.getImgPath(opt.img);
					break;
				case "subgrid":
					child.type = "image";
					elImg.src = KadDOM.getImgPath(opt.img);
					elImg.setAttribute("uiFilter", "invBackground");
					break;
				case "gridtitle":
					child.type = "image";
					elImg.src = KadDOM.getImgPath(opt.img);
					elImg.setAttribute("uiFilter", "invGridtitle");
					break;
				case "navbar":
					elImg.src = KadDOM.getImgPath(opt.img);
					elImg.setAttribute("uiFilter", "invNavbar");
					break;
				case "url":
				case "urlInvert":
					child.type = "image";
					elImg.src = opt.img;
					break;
				default:
					child.type = "text";
					child.textContent = opt.text;
			}
			this.UIOptions(child, opt);
			return child;
		},
		Lbl(opt) {
			const child = document.createElement("label");
			child.type = "Lbl";
			child.innerHTML = opt.text;
			if (opt.hasOwnProperty("for")) {
				child.for = opt.for;
				opt.pointer = true;
			}
			this.UIOptions(child, opt);
			return child;
		},
		H1(opt) {
			const child = document.createElement("H1");
			child.innerHTML = opt.text;
			this.UIOptions(child, opt);
			return child;
		},
		Sel(opt) {
			const child = document.createElement("select");
			opt.type = "sel";
			let start = 0;
			if (opt.hasOwnProperty("optionTitle")) {
				child.options[start] = new Option(opt.optionTitle);
				start = 1;
			}
			for (let n = 0; n < opt.options.length; n++) {
				child.options[n + start] = new Option(opt.options[n]);
			}
			this.UIOptions(child, opt);
			return child;
		},
		Colbox(opt) {
			const child = document.createElement("div");
			opt.type = "Colbox";
			child.classList.add("coloredBox");
			child.style.background = KadColor.formatAsCSS(opt.color, "HSL");
			this.UIOptions(child, opt);
			return child;
		},
		Img(opt) {
			const child = document.createElement("IMG");
			opt.type = "Img";
			switch (opt.subGroup) {
				case "button":
					child.src = KadDOM.getImgPath(opt.img);
					break;
				case "subgrid":
					child.src = KadDOM.getImgPath(opt.img);
					child.setAttribute("uiFilter", "invBackground");
					break;
				case "gridtitle":
					child.src = KadDOM.getImgPath(opt.img);
					child.setAttribute("uiFilter", "invGridtitle");
					break;
				case "navbar":
					child.src = KadDOM.getImgPath(opt.img);
					child.setAttribute("uiFilter", "invNavbar");
					break;
				case "url":
				case "urlInvert":
					child.src = opt.img;
					break;
			}
			this.UIOptions(child, opt);
			return child;
		},
		Div(opt) {
			const child = document.createElement("div");
			this.UIOptions(child, opt);
			return child;
		},
		UIOptions(cell, opt) {
			opt.name = KadTable.createName(opt);
			cell.id = `id${opt.type}_child${opt.name}`;
			if (opt.hasOwnProperty("idNoChild")) cell.id = `id${opt.type}${opt.name}`;
			if (opt.hasOwnProperty("datasets")) {
				for (const [key, value] of Object.entries(opt.datasets)) {
					cell.setAttribute(`data-${key}`, value);
				}
			}
			if (opt.hasOwnProperty("createClass")) {
				for (const cl of opt.createClass) {
					if (cl != "" && cl != null) cell.classList.add(cl);
				}
			}
			if (opt.hasOwnProperty("title")) {
				cell.title = opt.title;
			}
			if (opt.hasOwnProperty("style")) {
				for (const [key, value] of Object.entries(opt.style)) {
					cell.style[key] = value;
					if (opt.styleChild) cell.childNodes[0].style[key] = value;
				}
			}
			if (opt.hasOwnProperty("ui")) {
				for (const [key, value] of Object.entries(opt.ui)) {
					cell.setAttribute(key, value);
				}
			}
			if (opt.pointer || opt.copy) cell.style.cursor = "copy";
			if (opt.copy) {
				cell.addEventListener("click", () => copyToClipboard(cell.textContent), false);
			}
			if (opt.alias) cell.style.cursor = "alias";
			if (opt.hasOwnProperty("for")) cell.setAttribute("for", opt.for);
			if (opt.hasOwnProperty("oninput")) cell.addEventListener(opt.oninput, "input", false);
			if (opt.hasOwnProperty("onclick")) {
				cell.addEventListener("click", opt.onclick, false);
				// cell.style.cursor = "pointer";
			}
			if (opt.onmouseover) {
				cell.onmouseover = opt.onmouseover;
				cell.addEventListener("mouseover", opt.onmouseover);
			}
			if (opt.onmouseleave) {
				cell.onmouseleave = opt.onmouseleave;
				cell.addEventListener("mouseleave", opt.onmouseleave);
			}
		},
	},
	CellOptions(cell, opt) {
		if (opt.hasOwnProperty("cellStyle")) {
			for (const [key, value] of Object.entries(opt.cellStyle)) {
				cell.style[key] = value;
			}
		}
		if (opt.hasOwnProperty("createCellClass")) {
			for (const cl of opt.createCellClass) {
				if (cl != "" && cl != null) cell.classList.add(cl);
			}
		}
		if (opt.hasOwnProperty("cellOnclick")) {
			cell.addEventListener("click", opt.cellOnclick, false);
		}
		cell.colSpan = opt.colSpan || 1;
		cell.rowSpan = opt.rowSpan || 1;
	},
	createName(opt) {
		return opt.hasOwnProperty("name") ? opt.name : `_${opt.names.join("_")}`;
	},
};
export const KadColor = {
	types: {
		HEX: { postfix: ["", "", ""], stateRange: ["FF", "FF", "FF"] },
		RGB: { postfix: ["", "", ""], stateRange: [255, 255, 255] },
		HSL: { postfix: ["", "%", "%"], stateRange: [0, 0, 100] },
		HSB: { postfix: ["", "%", "%"], stateRange: [0, 0, 100] },
		CMYK: { postfix: ["%", "%", "%", "%"], stateRange: [0, 0, 0, 100] },
	},
	colAsArray(colArr = [], cFrom = "", cTo = "") {
		let colFrom = cFrom.toUpperCase();
		const colTo = cTo.toUpperCase();
		if (!Object.keys(this.types).includes(colFrom)) return;
		if (!Object.keys(this.types).includes(colTo)) return;
		let c = colArr;
		if (colFrom != "RGB" && colTo != "RGB") {
			c = this[`${colFrom}toRGB`](colArr);
			colFrom = "RGB";
		}
		return this[`${colFrom}to${colTo}`](c);
	},
	colAsString(colArr = [], cFrom = "", cTo = "") {
		const c = this.colAsArray(colArr, cFrom, cTo);
		return this.formatAsString(c, cTo);
	},
	colAsCSS(colArr = [], cFrom = "", cTo = "") {
		const c = this.colAsArray(colArr, cFrom, cTo);
		return this.formatAsCSS(c, cTo);
	},

	stateAsBool(colArr = [], type = "HSL", invert = false) {
		let RGB = type == "RGB" ? colArr : this[`${type}toRGB`](colArr);
		let inv = type == "CMYK" ? !invert : invert;
		let uicolors = [RGB[0] / 255, RGB[1] / 255, RGB[2] / 255];
		let c = uicolors.map((col) => {
			if (col <= 0.03928) return col / 12.92;
			return Math.pow((col + 0.055) / 1.055, 2.4);
		});
		let L = 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
		return Number(inv ? !(L < 0.179) : L < 0.179);
	},
	stateAsArray(colArr = [], type = "HSL", invert = false) {
		const state = this.stateAsBool(colArr, type, invert);
		let c = [];
		const range = KadColor.types[type.toUpperCase()].stateRange;
		for (let i = 0; i < range.length; i++) {
			c.push(state ? range[i] : 0);
		}
		return c;
	},
	stateAsString(colArr = [], type = "HSL", invert = false) {
		let c = this.stateAsArray(colArr, type, invert);
		return this.formatAsString(c, type);
	},
	stateAsCSS(colArr = [], type = "HSL", invert = false) {
		let c = this.stateAsArray(colArr, type, invert);
		return this.formatAsCSS(c, type);
	},

	formatAsString(colArray, type = "HSL") {
		if (typeof colArray === "string") return `${colArray.toUpperCase()}`;
		const typePostfix = KadColor.types[type].postfix;
		let retString = "";
		for (let i = 0; i < colArray.length; i++) {
			retString += i > typePostfix.length ? ` / ${colArray[i]}` : ` ${colArray[i]}${typePostfix[i]}`;
		}
		return retString;
	},
	formatAsCSS(colArray, type = "HSL") {
		if (typeof colArray === "string") return `${colArray.toUpperCase()}`;
		const typePostfix = KadColor.types[type].postfix;
		let retString = `${type.toLowerCase()}(`;
		for (let i = 0; i < colArray.length; i++) {
			if (i > typePostfix.length) {
				retString += ` / ${colArray[i]}`;
			} else {
				retString += ` ${colArray[i]}${typePostfix[i]}`;
			}
		}
		retString += ")";
		return retString;
	},

	HEXtoRGB(HEX) {
		let rgb = [];
		const hex = HEX.charAt(0) === "#" ? HEX.substring(1, 7) : HEX;
		rgb[0] = parseInt(hex.substring(0, 2), 16);
		rgb[1] = parseInt(hex.substring(2, 4), 16);
		rgb[2] = parseInt(hex.substring(4, 6), 16);
		return rgb;
	},
	HSLtoRGB(HSL) {
		let h = HSL[0] / 60;
		let s = HSL[1] / 100;
		let l = HSL[2] / 100;
		let rgb = [0, 0, 0];
		const C = (1 - Math.abs(2 * l - 1)) * s;
		const X = C * (1 - Math.abs((h % 2) - 1));
		if (h >= 0 && h < 1) {
			rgb[0] = C;
			rgb[1] = X;
		} else if (h >= 1 && h < 2) {
			rgb[0] = X;
			rgb[1] = C;
		} else if (h >= 2 && h < 3) {
			rgb[1] = C;
			rgb[2] = X;
		} else if (h >= 3 && h < 4) {
			rgb[1] = X;
			rgb[2] = C;
		} else if (h >= 4 && h < 5) {
			rgb[0] = X;
			rgb[2] = C;
		} else {
			rgb[0] = C;
			rgb[2] = X;
		}
		const m = l - C / 2;

		rgb[0] += m;
		rgb[1] += m;
		rgb[2] += m;
		rgb[0] = Math.round(rgb[0] * 255);
		rgb[1] = Math.round(rgb[1] * 255);
		rgb[2] = Math.round(rgb[2] * 255);
		return Object.values(rgb);
	},
	HSBtoRGB(HSB) {
		let s = HSB[0] / 360;
		let v = HSB[1] / 100;
		let h = HSB[2] / 100;
		let r, g, b;
		let i = Math.floor(h * 6);
		let f = h * 6 - i;
		let p = v * (1 - s);
		let q = v * (1 - f * s);
		let t = v * (1 - (1 - f) * s);
		switch (i % 6) {
			case 0:
				(r = v), (g = t), (b = p);
				break;
			case 1:
				(r = q), (g = v), (b = p);
				break;
			case 2:
				(r = p), (g = v), (b = t);
				break;
			case 3:
				(r = p), (g = q), (b = v);
				break;
			case 4:
				(r = t), (g = p), (b = v);
				break;
			case 5:
				(r = v), (g = p), (b = q);
				break;
		}
		r = Math.round(r * 255);
		g = Math.round(g * 255);
		b = Math.round(b * 255);
		return [r, g, b];
	},
	CMYKtoRGB(CMYK) {
		let k = CMYK[3];
		let r = 255 - Math.min(1, CMYK[0] * (1 - k) + k) * 255;
		let g = 255 - Math.min(1, CMYK[1] * (1 - k) + k) * 255;
		let b = 255 - Math.min(1, CMYK[2] * (1 - k) + k) * 255;
		return [r, g, b];
	},
	RGBtoHEX(RGB) {
		let rgb = RGB.length === 1 ? [RGB[0], RGB[0], RGB[0]] : RGB;
		let hex = "#";
		rgb.forEach((c) => {
			let tempHex = Number(c).toString(16);
			hex += tempHex.length < 2 ? `0${tempHex}` : tempHex;
		});
		return hex;
	},
	RGBtoHSL(RGB) {
		let r = RGB[0] / 255;
		let g = RGB[1] / 255;
		let b = RGB[2] / 255;
		const ma = Math.max(r, g, b);
		const mi = Math.min(r, g, b);
		let h = 0;
		let s = 0;
		let l = (ma + mi) / 2;
		if (ma != mi) {
			const d = ma - mi;
			if (r == ma) {
				h = (g - b) / d;
			} else if (g == ma) {
				h = 2 + (b - r) / d;
			} else if (b == ma) {
				h = 4 + (r - g) / d;
			}
			h = Math.min(h * 60, 360);
			if (h < 0) h += 360;
			s = l > 0.5 ? d / (2 - ma - mi) : d / (ma + mi);
		}
		h = Math.round(h);
		s = Math.round(s * 100);
		l = Math.round(l * 100);
		return [h, s, l];
	},
	RGBtoHSB(RGB) {
		let r = RGB[0];
		let g = RGB[1];
		let b = RGB[2];
		let max = Math.max(r, g, b);
		let min = Math.min(r, g, b);
		let d = max - min;
		let h;
		let s = max === 0 ? 0 : d / max;
		let v = max / 255;

		switch (max) {
			case min:
				h = 0;
				break;
			case r:
				h = g - b + d * (g < b ? 6 : 0);
				h /= 6 * d;
				break;
			case g:
				h = b - r + d * 2;
				h /= 6 * d;
				break;
			case b:
				h = r - g + d * 4;
				h /= 6 * d;
				break;
		}
		h = Math.round(h * 360);
		s = Math.round(s * 100);
		v = Math.round(v * 100);
		return [h, s, v];
	},
	RGBtoCMYK(RGB) {
		let r = RGB[0] / 255;
		let g = RGB[1] / 255;
		let b = RGB[2] / 255;
		let max = Math.max(r, g, b);
		let c, m, y;
		let k = 1 - max;
		if (k == 1) {
			c = 0;
			m = 0;
			y = 0;
		} else {
			c = (1 - r - k) / (1 - k);
			m = (1 - g - k) / (1 - k);
			y = (1 - b - k) / (1 - k);
		}
		c = Math.round(c * 100);
		m = Math.round(m * 100);
		y = Math.round(y * 100);
		k = Math.round(k * 100);
		return [c, m, y, k];
	},
	RGBtoLuminance(RGB) {
		let r = RGB[0] / 255;
		let g = RGB[1] / 255;
		let b = RGB[2] / 255;
		return 0.2126 * r + 0.7152 * g + 0.0722 * b;
	},
};
export class KadDebug {
	#lbl;
	#text;
	#startTime;
	#lapTime;
	#showLabel;
	constructor({ label = "", showLabel = false } = {}) {
		this.#lbl = label;
		this.#showLabel = showLabel;
		this.#startTime = this.#time();
		this.#lapTime = [this.#startTime];
		if (label) {
			this.#newText();
			this.#print();
		}
	}
	#time() {
		return new Date();
	}
	#print() {
		console.log(this.#text);
	}
	#label() {
		return this.#lbl && this.#showLabel ? `${this.#lbl} ` : "";
	}
	#newText(prompt) {
		this.#text = "";
		this.#addText(this.#label());
		this.#addText(prompt);
	}
	#addText(t) {
		if (t == undefined) return;
		this.#text += `${t} `;
	}
	#elapsedtTime() {
		return `${this.#time() - this.#startTime}ms`;
	}
	#intervalTime(num) {
		return `${((this.#time() - this.#startTime) / num).toFixed(3)} / ${num}`;
	}
	#addLap() {
		this.#lapTime.push(this.#time());
	}

	lap(prompt) {
		this.#addLap();
		this.#newText(prompt || "Lap:");
		this.#addText(this.#lapTime.at(-1) - this.#lapTime.at(-2));
		this.#addText("ms");
		this.#print();
	}

	now(prompt) {
		this.#newText(prompt);
		this.#addText(this.#elapsedtTime());
		this.#print();
	}
	reset() {
		this.#startTime = this.#time();
		this.#lapTime = [this.#startTime];
	}

	average(num, prompt = null) {
		this.#newText(prompt);
		this.#addText(this.#intervalTime(num));
		this.#addText(`(${this.#elapsedtTime()})`);
		this.#print();
	}
	startProfile() {
		console.profile();
	}
	stopProfile() {
		console.profileEnd();
	}
}
