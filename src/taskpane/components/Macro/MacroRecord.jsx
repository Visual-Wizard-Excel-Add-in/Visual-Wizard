import { Button } from "@fluentui/react-components";

import usePresetHandler from "../../hooks/usePresetHandler";
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

function MacroRecord() {
  const {
    presets,
    selectedPreset,
    addPresetHandler,
    deletePresetHandler,
    setSelectedPreset,
  } = usePresetHandler("allMacroPresets", "매크로");
  const isRecording = usePublicStore((state) => state.isRecording);
  const setIsRecording = usePublicStore((state) => state.setIsRecording);
  const setSelectMacroPreset = usePublicStore(
    (state) => state.setSelectMacroPreset,
  );
  const styles = useStyles();

  function controlMacroRecording() {
    if (selectedPreset === "") {
      popUpMessage("loadFail", "프리셋을 선택해주세요!");

      return;
    }

    manageRecording(!isRecording, selectedPreset);
    setIsRecording(!isRecording);
  }

  return (
    <div className="flex items-center justify-between space-x-5">
      <div className="flex items-center w-8/12 space-x-2">
        <button
          onClick={() => addPresetHandler()}
          className={styles.buttons}
          aria-label="add new preset"
          type="button"
        >
          <PlusIcon />
        </button>
        <CustomDropdown
          handleValue={(value) => {
            setSelectedPreset(value);
            setSelectMacroPreset(value);
          }}
          options={presets.map((preset) => ({
            name: preset,
            value: preset,
          }))}
          placeholder="매크로"
          selectedValue={selectedPreset}
        />
        <button
          onClick={() => deletePresetHandler()}
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
        onClick={() => macroPlay(selectedPreset)}
        size="small"
      >
        실행
      </Button>
    </div>
  );
}

export default MacroRecord;
