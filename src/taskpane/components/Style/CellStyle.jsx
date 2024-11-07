import { Button } from "@fluentui/react-components";

import usePresetHandler from "../../hooks/usePresetHandler";
import CustomDropdown from "../common/CustomDropdown";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import { useStyles } from "../../utils/style";
import { copyRangeStyle, pasteRangeStyle } from "../../utils/cellStyleFuncs";

function CellStyle() {
  const {
    presets,
    selectedPreset,
    addPresetHandler,
    deletePresetHandler,
    setSelectedPreset,
  } = usePresetHandler("cellStylePresets", "셀 서식");
  const styles = useStyles();

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
          handleValue={(value) => setSelectedPreset(value)}
          options={presets.map((preset) => ({
            name: preset,
            value: preset,
          }))}
          placeholder="프리셋"
          selectedValue={selectedPreset}
        />
        <button
          className={styles.buttons}
          onClick={() => deletePresetHandler()}
          aria-label="delete preset"
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={() => copyRangeStyle(selectedPreset)}
          className={styles.buttons}
          aria-label="save button"
          type="button"
        >
          <SaveIcon />
        </button>
      </div>
      <Button
        as="button"
        className="self-center w-7"
        onClick={() => pasteRangeStyle(selectedPreset)}
        size="small"
        aria-label="paste button"
      >
        적용
      </Button>
    </div>
  );
}

export default CellStyle;
