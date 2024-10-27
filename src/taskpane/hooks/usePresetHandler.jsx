import { useReducer, useEffect } from "react";

function presetReducer(state, action) {
  switch (action.type) {
    case "SET_PRESETS":
      return { ...state, presets: action.payload };
    case "SET_SELECTED":
      return { ...state, selectedPreset: action.payload };
    case "ADD_PRESET":
      return { ...state, presets: [...state.presets, action.payload] };
    case "DELETE_PRESET":
      return {
        ...state,
        presets: state.presets.filter((preset) => preset !== action.payload),
      };
    default:
      return state;
  }
}

function usePresetHandler(presetName, unitName) {
  const [state, dispatch] = useReducer(presetReducer, {
    presets: [],
    selectedPreset: "",
  });

  const { selectedPreset: currentPreset } = state;

  useEffect(() => {
    fetchPresets();

    async function fetchPresets() {
      const savedPresets = await sortPreset();

      dispatch({ type: "SET_PRESETS", payload: savedPresets });

      if (savedPresets.length > 0 && !currentPreset) {
        dispatch({ type: "SET_SELECTED", payload: savedPresets[0] });
      }

      async function sortPreset() {
        return Object.keys(await load()).sort((a, b) => {
          const numA = parseInt(a.replace(/\D/g, ""), 10);
          const numB = parseInt(b.replace(/\D/g, ""), 10);

          return numA - numB;
        });
      }
    }

    async function load() {
      return JSON.parse(await OfficeRuntime.storage.getItem(presetName)) ?? {};
    }
  }, [currentPreset]);

  async function addPresetHandler() {
    let existOrderNums = [];

    if (state.presets.length > 0) {
      existOrderNums = state.presets.map((preset) =>
        parseInt(preset.replace(/\D/g, ""), 10),
      );
    }

    let newOrderNum = 1;

    while (existOrderNums.includes(newOrderNum)) {
      newOrderNum += 1;
    }

    const newUnitName = `${unitName}${newOrderNum}`;

    await saveNewPreset();

    dispatch({ type: "ADD_PRESET", payload: newUnitName });
    dispatch({ type: "SET_SELECTED", payload: newUnitName });

    async function saveNewPreset() {
      const loadedPresets =
        JSON.parse(await OfficeRuntime.storage.getItem(presetName)) ?? {};

      loadedPresets[newUnitName] = {};

      await OfficeRuntime.storage.setItem(
        presetName,
        JSON.stringify(loadedPresets),
      );
    }
  }

  async function deletePresetHandler() {
    await deletePreset();

    const selectIndex = state.presets.indexOf(currentPreset) - 1;

    dispatch({ type: "DELETE_PRESET", payload: currentPreset });

    if (state.presets.length > 0) {
      dispatch({ type: "SET_SELECTED", payload: state.presets[selectIndex] });
    }

    async function deletePreset() {
      const loadedPresets = JSON.parse(
        await OfficeRuntime.storage.getItem(presetName),
      );

      if (loadedPresets) {
        delete loadedPresets[currentPreset];

        await OfficeRuntime.storage.setItem(
          presetName,
          JSON.stringify(loadedPresets),
        );
      }
    }
  }

  return {
    presets: state.presets,
    selectedPreset: state.selectedPreset,
    addPresetHandler,
    deletePresetHandler,
    setSelectedPreset: (preset) =>
      dispatch({ type: "SET_SELECTED", payload: preset }),
  };
}

export default usePresetHandler;
