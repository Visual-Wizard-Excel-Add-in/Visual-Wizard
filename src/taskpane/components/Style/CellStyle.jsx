import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import usePublicStore from "../../store/publicStore";
import { useStyles } from "../../utils/style";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import {
  saveRangeStylePreset,
  loadRangeStylePreset,
} from "../../utils/cellStyleFuncs";
import PresetHandler from "../../classes/PresetHandler";

function CellStyle() {
  const [cellStylePresets, setCellStylePresets] = useState([]);
  const selectedStylePreset = usePublicStore(
    (state) => state.selectedStylePreset,
  );
  const setSelectedStylePreset = usePublicStore(
    (state) => state.setSelectedStylePreset,
  );
  const styles = useStyles();
  const presets = new PresetHandler("cellStylePresets", "셀 서식");

  useEffect(() => {
    async function fetchPresets() {
      const sortedPresets = await presets.sorting();

      setCellStylePresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectedStylePreset) {
        setSelectedStylePreset(sortedPresets[0]);
      }
    }

    fetchPresets();
  }, [selectedStylePreset]);

  async function newPreset() {
    setSelectedStylePreset(await presets.add(cellStylePresets));
    setCellStylePresets(await presets.sorting());
  }

  async function deletePreset() {
    const selectIndex = cellStylePresets.indexOf(selectedStylePreset);

    setCellStylePresets(await presets.delete(selectedStylePreset));
    setSelectedStylePreset((await presets.sorting())[selectIndex]);
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
          handleValue={(value) => setSelectedStylePreset(value)}
          options={cellStylePresets.map((preset) => ({
            name: preset,
            value: preset,
          }))}
          placeholder="프리셋"
          selectedValue={selectedStylePreset}
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
          onClick={() => saveRangeStylePreset(selectedStylePreset)}
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
        onClick={() => loadRangeStylePreset(selectedStylePreset)}
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default CellStyle;
