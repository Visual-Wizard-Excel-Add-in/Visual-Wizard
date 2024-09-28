import { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import useStore from "../../utils/store";
import CustomDropdown from "../common/CustomDropdown";
import {
  SaveIcon,
  DeleteIcon,
  RecordStart,
  RecordStop,
  PlusIcon,
} from "../../utils/icons";
import { addPreset, deletePreset, popUpMessage } from "../../utils/commonFuncs";
import { manageRecording, macroPlay } from "../../utils/macroFuncs";

function MacroRecord() {
  const [macroPresets, setMacroPresets] = useState([]);
  const [isRecording, setIsRecording, selectMacroPreset, setSelectMacroPreset] =
    useStore((state) => [
      state.isRecording,
      state.setIsRecording,
      state.selectMacroPreset,
      state.setSelectMacroPreset,
    ]);
  const styles = useStyles();

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();

      setMacroPresets(Object.keys(savedPresets));
    }

    fetchPresets();
  }, []);

  async function newPreset() {
    let lastPresetNum = 0;

    if (macroPresets.length > 0) {
      lastPresetNum = Number(
        macroPresets[macroPresets.length - 1].split("로")[1],
      );
    }

    if (macroPresets.includes("매크로1")) {
      await addPreset("allMacroPresets", `매크로${lastPresetNum + 1}`);
    } else {
      await addPreset("allMacroPresets", "매크로1");
    }

    const savedPresets = await loadPresets();
    const sortedPresets = Object.keys(savedPresets).sort((a, b) =>
      a.localeCompare(b),
    );

    setMacroPresets(sortedPresets);
  }

  function controlMacroRecording() {
    if (selectMacroPreset === "") {
      popUpMessage("loadFail", "프리셋을 선택해주세요!");

      return;
    }

    manageRecording(!isRecording, selectMacroPreset);
    setIsRecording(!isRecording);
  }

  async function loadPresets() {
    let presets = await OfficeRuntime.storage.getItem("allMacroPresets");

    if (!presets) {
      presets = {};
    } else {
      presets = JSON.parse(presets);
    }

    return presets;
  }

  async function handleDeletePreset() {
    if (!setSelectMacroPreset) {
      return;
    }

    await deletePreset("allMacroPresets", selectMacroPreset);

    const savedPresets = await loadPresets();

    setSelectMacroPreset("");
    setMacroPresets(Object.keys(savedPresets));
  }

  return (
    <>
      <div className="flex items-center justify-between space-x-5">
        <div className="flex items-center w-8/12 space-x-2">
          <button
            onClick={newPreset}
            className={styles.buttons}
            aria-label="plus"
            type="button"
          >
            <PlusIcon />
          </button>
          <CustomDropdown
            handleValue={(value) => setSelectMacroPreset(value)}
            options={macroPresets.map((preset) => ({
              name: preset,
              value: preset,
            }))}
            placeholder="매크로"
            selectedValue={selectMacroPreset}
          />
          <button
            onClick={handleDeletePreset}
            className={styles.buttons}
            aria-label="delete"
            type="button"
          >
            <DeleteIcon />
          </button>
          <button
            onClick={() => {}}
            className={styles.buttons}
            aria-label="save"
            type="button"
          >
            <SaveIcon />
          </button>
        </div>
        <Button
          as="button"
          className="self-center"
          onClick={() => macroPlay(selectMacroPreset)}
          size="small"
        >
          실행
        </Button>
      </div>
      <div className="flex items-center justify-between space-x-5">
        <span className="h-6">매크로 녹화</span>
        <button
          onClick={controlMacroRecording}
          className={styles.buttons}
          aria-label="record"
          type="button"
        >
          {isRecording ? <RecordStop /> : <RecordStart color="red" />}
        </button>
      </div>
    </>
  );
}

export default MacroRecord;
