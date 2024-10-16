import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import { useStyles } from "../../utils/style";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import {
  saveRangeStylePreset,
  loadRangeStylePreset,
} from "../../utils/cellStyleFuncs";
import PresetHandler from "../../classes/PresetHandler";

function CellStyle() {
  const [selectPreset, setSelectPreset] = useState("");
  const [cellStylePresets, setCellStylePresets] = useState([]);
  const styles = useStyles();
  const presets = new PresetHandler("cellStylePresets", "셀 서식");

  useEffect(() => {
    async function fetchPresets() {
      const sortedPresets = await presets.sorting();
    fetchPresets();

      setCellStylePresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectedStylePreset) {
        setSelectedStylePreset(sortedPresets[0]);
        setSelectPreset((await presets.sort())[0]);
      }
    }
  }, [selectPreset]);

  async function newPreset() {
    setCellStylePresets(await presets.sorting());
    setSelectPreset(await presets.add(cellStylePresets));
  }

  async function deletePreset() {
    const selectIndex = cellStylePresets.indexOf(selectPreset);
    const leftPresets = await presets.delete(selectPreset);

    setCellStylePresets(leftPresets);
    setSelectPreset(leftPresets[selectIndex]);
  }

  return (
    <div className="flex items-center justify-between space-x-5">
      <div className="flex items-center w-8/12 space-x-2">
        <button
          onClick={() => newPreset()}
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
          onClick={deletePreset}
          aria-label="delete preset"
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={() => saveRangeStylePreset(selectPreset)}
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
        onClick={() => loadRangeStylePreset(selectPreset)}
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default CellStyle;
