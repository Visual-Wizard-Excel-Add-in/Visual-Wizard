import { create } from "zustand";

const usePublicStore = create((set) => ({
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

  messageList: [],
  setMessageList: (message) =>
    set((state) => ({
      messageList: [...state.messageList, { id: Date.now(), message }],
    })),
  removeMessage: (id) =>
    set((state) => ({
      messageList: state.messageList.filter((message) => message.id !== id),
    })),

  sheetId: "",
  setSheetId: (selectSheet) => set({ sheetId: selectSheet }),

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

  selectMacroPreset: "",
  setSelectMacroPreset: (selectedMacroPreset) =>
    set({ selectMacroPreset: selectedMacroPreset }),
}));

export default usePublicStore;
