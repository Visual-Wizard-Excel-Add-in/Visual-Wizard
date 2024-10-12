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

export default STYLE_OPTIONS_TO_LOAD;
