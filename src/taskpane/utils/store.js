import { create } from "zustand";

const useStore = create((set) => ({
  category: "",
  setCategory: (selectedCategory) => set({ category: selectedCategory }),
}));

export default useStore;
