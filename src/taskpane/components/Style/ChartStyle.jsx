import { Button } from "@fluentui/react-components";

import usePresetHandler from "../../hooks/usePresetHandler";
import CustomDropdown from "../common/CustomDropdown";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import { useStyles } from "../../utils/style";
import { copyChartStyle, pasteChartStyle } from "../../utils/chartStyleFuncs";

function ChartStyle() {
  const {
    presets,
    selectedPreset,
    addPresetHandler,
    deletePresetHandler,
    setSelectedPreset,
  } = usePresetHandler("chartStylePresets", "차트 서식");
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
          onClick={() => deletePresetHandler()}
          className={styles.buttons}
          aria-label="delete preset"
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={() => copyChartStyle("chartStylePresets", selectedPreset)}
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
        onClick={() => pasteChartStyle("chartStylePresets", selectedPreset)}
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default ChartStyle;
