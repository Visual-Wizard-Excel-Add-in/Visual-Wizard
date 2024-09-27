import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import useStore from "../../utils/store";
import { useStyles } from "../../utils/style";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import {
  saveCellStylePreset,
  loadCellStylePreset,
} from "../../utils/cellStyleFuncs";
import { addPreset, deletePreset } from "../../utils/commonFuncs";

function CellStyle() {
  const [cellStylePresets, setCellStylePresets] = useState([]);
  const [selectedStylePreset, setSelectedStylePreset] = useStore((state) => [
    state.selectedStylePreset,
    state.setSelectedStylePreset,
  ]);
  const styles = useStyles();

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();

      setCellStylePresets(Object.keys(savedPresets));

      if (Object.keys(savedPresets).length > 0 && !selectedStylePreset) {
        setSelectedStylePreset(Object.keys(savedPresets)[0]);
      }
    }

    fetchPresets();
  }, [selectedStylePreset]);

  async function loadPresets() {
    let presets = await OfficeRuntime.storage.getItem("cellStylePresets");

    if (!presets) {
      presets = {};
    } else {
      presets = JSON.parse(presets);
    }

    return presets;
  }

  async function newPreset() {
    let lastPresetNum = 0;

    if (cellStylePresets.length > 0) {
      lastPresetNum = Number(
        cellStylePresets[cellStylePresets.length - 1].split("식")[1],
      );
    }

    let newPresetName = "셀 서식1";

    if (cellStylePresets.includes("셀 서식1")) {
      newPresetName = `셀 서식${lastPresetNum + 1}`;
    }

    await addPreset("cellStylePresets", newPresetName);

    const savedPresets = await loadPresets();
    const sortedPresets = Object.keys(savedPresets).sort((a, b) =>
      a.localeCompare(b),
    );

    setCellStylePresets(sortedPresets);

    setSelectedStylePreset(newPresetName);
  }

  async function handleDeletePreset() {
    if (!selectedStylePreset) {
      return;
    }

    await deletePreset("cellStylePresets", selectedStylePreset);

    const savedPresets = await loadPresets();

    setSelectedStylePreset("");
    setCellStylePresets(Object.keys(savedPresets));
  }

  return (
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
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={() => saveCellStylePreset(selectedStylePreset)}
          className={styles.buttons}
          aria-label="save"
          type="button"
        >
          <SaveIcon />
        </button>
      </div>
      <Button
        as="button"
        className="self-center w-7"
        onClick={() => loadCellStylePreset(selectedStylePreset)}
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default CellStyle;
