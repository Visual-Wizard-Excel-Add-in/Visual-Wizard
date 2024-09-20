import { create } from "zustand";

interface StoreType {
  category: string;
  setCategory: (selectedCategory: string) => void;

  openTab: string[];
  setOpenTab: (openedTabs: string[]) => void;

  isRecording: boolean;
  setIsRecording: (recordState: boolean) => void;

  isCellHighlighting: boolean;
  setIsCellHighlighting: () => void;

  selectedStylePreset: string;
  setSelectedStylePreset: (preset: string) => void;

  messageList: { id: number; message: string }[];
  setMessageList: (message: string) => void;
  removeMessage: (id: number) => void;

  sheetId: string;
  setSheetId: (selectSheet: string) => void;

  cellValue: string;
  setCellValue: (selectedCellValue: string) => void;

  cellAddress: string;
  setCellAddress: (selectedCellAddress: string) => void;

  cellFormula: string;
  setCellFormula: (selectedCellFormula: string) => void;

  formulaSteps: string[];
  setFormulaSteps: (currentFormulaSteps: string[]) => void;

  cellFunctions: string[];
  setCellFunctions: (selectedCellFunctions: string[]) => void;

  cellArguments: string[];
  setCellArguments: (selectedCellArguments: string[]) => void;

  selectMacroPreset: string;
  setSelectMacroPreset: (selectedMacroPreset: string) => void;
}

const useStore = create<StoreType>((set) => ({
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

export default useStore;
