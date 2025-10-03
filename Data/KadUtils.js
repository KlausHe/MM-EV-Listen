export function dbID(id) {
  if (id instanceof Object) return id;
  return document.getElementById(id);
}
export function daEL(id, type, fn) {
  dbID(id).addEventListener(type, fn);
}
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
    if (["DIV", "LABEL"].includes(obj.nodeName)) {
      let s = state ? null : "negative";
      this.btnColor(obj, s);
    }
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
