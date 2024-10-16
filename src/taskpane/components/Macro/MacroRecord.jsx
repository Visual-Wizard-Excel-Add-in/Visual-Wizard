import { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import usePublicStore from "../../store/publicStore";
import CustomDropdown from "../common/CustomDropdown";
import {
  DeleteIcon,
  RecordStart,
  RecordStop,
  PlusIcon,
} from "../../utils/icons";
import { popUpMessage } from "../../utils/commonFuncs";
import { manageRecording, macroPlay } from "../../utils/macroFuncs";
import PresetHandler from "../../classes/PresetHandler";

function MacroRecord() {
  const [macroPresets, setMacroPresets] = useState([]);
  const isRecording = usePublicStore((state) => state.isRecording);
  const setIsRecording = usePublicStore((state) => state.setIsRecording);
  const selectMacroPreset = usePublicStore((state) => state.selectMacroPreset);
  const setSelectMacroPreset = usePublicStore(
    (state) => state.setSelectMacroPreset,
  );
  const styles = useStyles();
  const presets = new PresetHandler("allMacroPresets", "매크로");

  useEffect(() => {
    async function fetchPresets() {
      const sortedPresets = await presets.sort();

      setMacroPresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectMacroPreset) {
        setSelectMacroPreset(sortedPresets[0]);
      }
    }

    fetchPresets();
  }, [selectMacroPreset]);

  async function newPreset() {
    setSelectMacroPreset(await presets.add(macroPresets));
    setMacroPresets(await presets.sort());
  }

  async function handleDeletePreset() {
    const selectIndex = macroPresets.indexOf(selectMacroPreset);

    setMacroPresets(Object.keys(await presets.delete(selectMacroPreset)));
    setSelectMacroPreset((await presets.sort())[selectIndex]);
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
