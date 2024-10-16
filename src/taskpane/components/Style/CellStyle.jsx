import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import { useStyles } from "../../utils/style";
import PresetHandler from "../../classes/PresetHandler";
import { copyRangeStyle, pasteRangeStyle } from "../../utils/cellStyleFuncs";

function CellStyle() {
  const [selectPreset, setSelectPreset] = useState("");
  const [cellStylePresets, setCellStylePresets] = useState([]);
  const styles = useStyles();
  const presets = new PresetHandler("cellStylePresets", "셀 서식");

  useEffect(() => {
    fetchPresets();

    async function fetchPresets() {
      setCellStylePresets(await presets.sort());

      if ((await presets.sort()).length > 0 && !selectPreset) {
        setSelectPreset((await presets.sort())[0]);
      }
    }
  }, [selectPreset]);

  async function newPresetHandler() {
    setSelectPreset(await presets.add(cellStylePresets));
    setCellStylePresets(await presets.sort());
  }

  async function deletePresetHandler() {
    const selectIndex = cellStylePresets.indexOf(selectPreset);
    const leftPresets = await presets.delete(selectPreset);

    setCellStylePresets(leftPresets);
    setSelectPreset(leftPresets[selectIndex]);
  }

  return (
    <div className="flex items-center justify-between space-x-5">
      <div className="flex items-center w-8/12 space-x-2">
        <button
          onClick={() => newPresetHandler()}
          className={styles.buttons}
          aria-label="add new preset"
          type="button"
        >
          <PlusIcon />
        </button>
        <CustomDropdown
          handleValue={(value) => setSelectPreset(value)}
          options={cellStylePresets.map((preset) => ({
            name: preset,
            value: preset,
          }))}
          placeholder="프리셋"
          selectedValue={selectPreset}
        />
        <button
          className={styles.buttons}
          onClick={deletePresetHandler}
          aria-label="delete preset"
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={() => copyRangeStyle(selectPreset)}
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
        onClick={() => pasteRangeStyle(selectPreset)}
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default CellStyle;
