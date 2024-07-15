import { create } from "zustand";

const useStore = create((set) => ({
  category: "Formula",
  setCategory: (selectedCategory) => set({ category: selectedCategory }),
  openTab: [],
  setOpenTab: (openedTabs) => set({ openTab: openedTabs }),
  isRecording: false,
  setIsRecording: (recordState) => set({ isRecording: recordState }),
}));

export default useStore;
