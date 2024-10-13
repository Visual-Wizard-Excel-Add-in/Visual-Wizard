class CellInfo {
  constructor(cell) {
    this.address = cell.address;
    this.scale = cell.address.includes(":") ? "range" : "cell";
    [[this.values]] = cell.values;
    [[this.formula]] = cell.formulas;
    [[this.numberFormat]] = cell.numberFormat;
  }

  isDate() {
    if (this.numberFormat && this.numberFormat.includes("yy")) {
      return true;
    }

    return false;
  }

  get date() {
    return Date((this.values - 25569) * 86400 * 1000).toLocaleDateString();
  }

  // isFormula() {
  //   if (typeof this.formula === "string" && !this.formula.startsWith("=")) {
  //     return false;
  //   }

  //   return true;
  // }
}

export default CellInfo;
