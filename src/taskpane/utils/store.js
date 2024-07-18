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
  cellValue: "",
  setCellValue: (selectedCellValue) => set({ cellValue: selectedCellValue }),
  cellAddress: "",
  setCellAddress: (selectedCellAddress) =>
    set({ cellAddress: selectedCellAddress }),
  cellFormulas: [],
  setCellFormulas: (selectedCellFormulas) =>
    set({ cellFormulas: selectedCellFormulas }),
  cellArguments: [],
  setCellArguments: (selectedCellArguments) =>
    set({ cellArguments: selectedCellArguments }),
}));

export default useStore;
