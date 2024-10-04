import { create } from "zustand";

const useHandlerStore = create((set) => ({
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
}));

export default useHandlerStore;
