const STYLE_OPTIONS_TO_LOAD = {
  address: true,
  format: {
    fill: {
      color: true,
    },
    font: {
      name: true,
      color: true,
      size: true,
      bold: true,
      italic: true,
      underline: true,
      strikethrough: true,
    },
    borders: {
      color: true,
      style: true,
      weight: true,
      tintAndShade: true,
    },
    horizontalAlignment: true,
    verticalAlignment: true,
    wrapText: true,
    indentLevel: true,
    readingOrder: true,
    textOrientation: true,
  },
  numberFormat: true,
  numberFormatLocal: true,

  protection: {
    locked: true,
    formulaHidden: true,
  },
};

const HIGHLIGHT_STYLES = {
  resultFill: {
    fill: { color: "#3d33ff" },
  },
  argsFill: {
    fill: { color: "#28f925" },
  },
  borders: {
    bottom: {
      color: "red",
      weight: Excel.BorderWeight.thick,
      style: Excel.BorderLineStyle.continuous,
    },
    top: {
      color: "red",
      weight: Excel.BorderWeight.thick,
      style: Excel.BorderLineStyle.continuous,
    },
    left: {
      color: "red",
      weight: Excel.BorderWeight.thick,
      style: Excel.BorderLineStyle.continuous,
    },
    right: {
      color: "red",
      weight: Excel.BorderWeight.thick,
      style: Excel.BorderLineStyle.continuous,
    },
  },
};

export { STYLE_OPTIONS_TO_LOAD, HIGHLIGHT_STYLES };
