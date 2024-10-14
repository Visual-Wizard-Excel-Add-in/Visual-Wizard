class CellInfo {
  constructor(cell) {
    [[this.numberFormat]] = cell.numberFormat;
    [[this.formula]] = cell.formulas;
    [[this.values]] = cell.values;
    this.address = cell.address;
  }

  get values() {
    return this._values;
  }

  set values(cellValues) {
    if (this.numberFormat && this.numberFormat.includes("yy")) {
      const dateValue = (cellValues - 25569) * 86400 * 1000;
      this._values = new Intl.DateTimeFormat("ko-KR").format(dateValue);
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
