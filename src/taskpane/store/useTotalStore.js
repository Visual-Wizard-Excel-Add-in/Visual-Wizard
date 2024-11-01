import { create } from "zustand";
import createHandlerSlice from "./createHandlerSlice";
import createPublicSlice from "./createPublicSlice";

export const useTotalStore = create((...a) => ({
  ...createHandlerSlice(...a),
  ...createPublicSlice(...a),
}));

export default useTotalStore;
