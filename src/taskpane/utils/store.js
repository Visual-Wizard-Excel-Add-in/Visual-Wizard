import { create } from "zustand";

const useStore = create((set) => ({
  category: "Formula",
  setCategory: (selectedCategory) => set({ category: selectedCategory }),

  openTab: [],
  setOpenTab: (openedTabs) => set({ openTab: openedTabs }),

  isRecording: false,
  setIsRecording: (recordState) => set({ isRecording: recordState }),

  isCellHighlighting: false,
  setIsCellHighlighting: () =>
    set((state) => ({ isCellHighlighting: !state.isCellHighlighting })),

  selectedStylePreset: "",
  setSelectedStylePreset: (preset) => set({ selectedStylePreset: preset }),

  sheetId: "",
  setSheetId: (selectSheet) => set({ sheetName: selectSheet }),

  cellValue: "",
  setCellValue: (selectedCellValue) => set({ cellValue: selectedCellValue }),

  cellAddress: "",
  setCellAddress: (selectedCellAddress) =>
    set({ cellAddress: selectedCellAddress }),

  cellFormula: "",
  setCellFormula: (selectedCellFormula) =>
    set({ cellFormula: selectedCellFormula }),

  formulaSteps: [],
  setFormulaSteps: (currentFormulaSteps) =>
    set({ formulaSteps: currentFormulaSteps }),

  cellFunctions: [],
  setCellFunctions: (selectedCellFunctions) =>
    set({ cellFunctions: selectedCellFunctions }),

  cellArguments: [],
  setCellArguments: (selectedCellArguments) =>
    set({ cellArguments: selectedCellArguments }),
}));

export default useStore;
