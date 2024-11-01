const createHandlerSlice = (set) => ({
  selectionChangeHandler: null,
  setSelectionChangeHandler: (handler) =>
    set({ selectionChangeHandler: handler }),

  worksheetChangedHandler: null,
  setWorksheetChangedHandler: (handler) =>
    set({ worksheetChangedHandler: handler }),

  tableChangedHandler: null,
  setTableChangedHandler: (handler) => set({ tableChangedHandler: handler }),

  chartAddedHandler: null,
  setChartAddedHandler: (handler) => set({ chartAddedHandler: handler }),

  tableAddedHandler: null,
  setTableAddedHandler: (handler) => set({ tableAddedHandler: handler }),

  formatChangedHandler: null,
  setFormatChangedHandler: (handler) => set({ formatChangedHandler: handler }),
});

export default createHandlerSlice;
