class CellInfo {
  constructor(address, values, formula, numberFormat) {
    this.address = address;
    this.scale = address.includes(":") ? "range" : "cell";
    this.values = this.isDate(values) ? this.date() : values;
    this.formula = this.isFormula() ? formula : "";
    this.numberFormat = numberFormat;
  }

  isDate() {
    if (
      this.numberFormat &&
      this.numberFormat.includes("yy") &&
      this.values !== ""
    ) {
      return true;
    }

    return false;
  }

  date() {
    return Date((this.values - 25569) * 86400 * 1000).toLocaleDateString();
  }

  isFormula() {
    if (typeof this.formula === "string" && !this.formula.startsWith("=")) {
      return false;
    }

    return true;
  }
}

export default CellInfo;
