class CellInfo {
  constructor(cell) {
    [[this.numberFormat]] = cell.numberFormat;
    [[this.formula]] = cell.formulas;
    [[this.values]] = cell.values;
    this.address = cell.address.replaceAll("'", "");
    this.arguments = cell.arguments;
    this.functions = getFunctions(this.formula);
  }

  get values() {
    return this._values;
  }

  set values(cellValues) {
    if (isDate(this.numberFormat)) {
      this._values = new Intl.DateTimeFormat("ko-KR").format(
        (cellValues - 25569) * 86400 * 1000,
      );
    } else {
      this._values = cellValues;
    }
  }

  get formula() {
    return this._formula;
  }

  set formula(cellFormula) {
    if (typeof this.formula === "string" && !this.formula.startsWith("=")) {
      this._formula = "";
    } else {
      this._formula = cellFormula;
    }
  }
}

export default CellInfo;

function getFunctions(formula) {
  const regex = /([A-Z]+)\(/g;
  const result = [
    ...new Set([...formula.matchAll(regex)].map((match) => match[1])),
  ];

  return result;
}

function isDate(numberFormat) {
  return (
    numberFormat?.includes("yy") ||
    numberFormat.includes("mm") ||
    numberFormat.includes("dd")
  );
}
