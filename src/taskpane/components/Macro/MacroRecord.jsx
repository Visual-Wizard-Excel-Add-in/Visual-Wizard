import { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import useStore from "../../utils/store";
import CustomDropdown from "../common/CustomDropdown";
import {
  DeleteIcon,
  RecordStart,
  RecordStop,
  PlusIcon,
} from "../../utils/icons";
import { addPreset, deletePreset, popUpMessage } from "../../utils/commonFuncs";
import { manageRecording, macroPlay } from "../../utils/macroFuncs";

function MacroRecord() {
  const [macroPresets, setMacroPresets] = useState([]);
  const isRecording = useStore((state) => state.isRecording);
  const setIsRecording = useStore((state) => state.setIsRecording);
  const selectMacroPreset = useStore((state) => state.selectMacroPreset);
  const setSelectMacroPreset = useStore((state) => state.setSelectMacroPreset);
  const styles = useStyles();

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();
      const sortedPresets = Object.keys(savedPresets).sort((a, b) => {
        const numA = parseInt(a.replace(/\D/g, ""), 10);
        const numB = parseInt(b.replace(/\D/g, ""), 10);

        return numA - numB;
      });

      setMacroPresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectMacroPreset) {
        setSelectMacroPreset(sortedPresets[0]);
      }
    }

    fetchPresets();
  }, [selectMacroPreset]);

  async function loadPresets() {
    let presets = await OfficeRuntime.storage.getItem("allMacroPresets");

    if (!presets) {
      presets = {};
    } else {
      presets = JSON.parse(presets);
    }

    return presets;
  }

  async function newPreset() {
    let presetNumbers = [];

    if (macroPresets.length > 0) {
      presetNumbers = macroPresets.map((preset) =>
        parseInt(preset.replace(/\D/g, ""), 10),
      );
    }

    let lastPresetNum = 1;
    while (presetNumbers.includes(lastPresetNum)) {
      lastPresetNum += 1;
    }

    const newPresetName = `매크로${lastPresetNum}`;

    await addPreset("allMacroPresets", newPresetName);

    const savedPresets = await loadPresets();
    const sortedPresets = Object.keys(savedPresets).sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ""), 10);
      const numB = parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    setMacroPresets(sortedPresets);
    setSelectMacroPreset(newPresetName);
  }

  async function handleDeletePreset() {
    if (!selectMacroPreset) {
      return;
    }

    const selectIndex = macroPresets.indexOf(selectMacroPreset);

    await deletePreset("allMacroPresets", selectMacroPreset);

    const savedPresets = Object.keys(await loadPresets());
    const sortedPresets = savedPresets.sort((a, b) => {
      const numA = +parseInt(a.replace(/\D/g, ""), 10);
      const numB = +parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    setMacroPresets(Object.keys(sortedPresets));
    setSelectMacroPreset(sortedPresets[selectIndex]);
  }

  function controlMacroRecording() {
    if (selectMacroPreset === "") {
      popUpMessage("loadFail", "프리셋을 선택해주세요!");

      return;
    }

    manageRecording(!isRecording, selectMacroPreset);
    setIsRecording(!isRecording);
  }

  return (
    <div className="flex items-center justify-between space-x-5">
      <div className="flex items-center w-8/12 space-x-2">
        <button
          onClick={newPreset}
          className={styles.buttons}
          aria-label="add new preset"
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
          aria-label="delete preset"
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={controlMacroRecording}
          className={styles.buttons}
          aria-label="record button"
          type="button"
        >
          {isRecording ? <RecordStop /> : <RecordStart color="red" />}
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
  );
}

export default MacroRecord;
