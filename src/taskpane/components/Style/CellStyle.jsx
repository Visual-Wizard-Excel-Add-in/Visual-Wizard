import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import useStore from "../../utils/store";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import {
  saveCellStylePreset,
  loadCellStylePreset,
} from "../../utils/cellStyleFunc";
import { addPreset, deletePreset } from "../../utils/cellCommonUtils";
import CustomDropdown from "../common/CustomDropdown";

function CellStyle() {
  const styles = useStyles();
  const { selectedStylePreset, setSelectedStylePreset } = useStore();
  const [cellStylePresets, setCellStylePresets] = useState([]);

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();

      setCellStylePresets(Object.keys(savedPresets));

      if (Object.keys(savedPresets).length > 0 && !selectedStylePreset) {
        setSelectedStylePreset(Object.keys(savedPresets)[0]);
      }
    }

    fetchPresets();
  }, []);

  async function loadPresets() {
    return Excel.run(async () => {
      let presets = Office.context.document.settings.get("cellStylePreset");

      if (!presets) {
        presets = {};
      } else {
        presets = JSON.parse(presets);
      }

      return presets;
    });
  }

  async function newPreset() {
    await addPreset("cellStylePreset", `셀 서식${cellStylePresets.length + 1}`);

    const savedPresets = await loadPresets();

    setCellStylePresets(Object.keys(savedPresets));
  }

  async function handleDeletePreset() {
    if (!selectedStylePreset) {
      return;
    }

    await deletePreset("cellStylePreset", selectedStylePreset);

    const savedPresets = await loadPresets();

    setSelectedStylePreset("");
    setCellStylePresets(Object.keys(savedPresets));
  }

  return (
    <>
      <div className="flex items-center justify-between space-x-5">
        <div className="flex items-center space-x-2">
          <button
            onClick={newPreset}
            className={styles.buttons}
            aria-label="plus"
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
            onClick={handleDeletePreset}
            aria-label="delete"
          >
            <DeleteIcon />
          </button>
          <button
            onClick={() => saveCellStylePreset(selectedStylePreset)}
            className={styles.buttons}
            aria-label="save"
          >
            <SaveIcon />
          </button>
        </div>
      </div>
      <Button
        as="button"
        className="self-center w-7"
        onClick={() => loadCellStylePreset(selectedStylePreset)}
        size="small"
      >
        적용
      </Button>
    </>
  );
}

export default CellStyle;
