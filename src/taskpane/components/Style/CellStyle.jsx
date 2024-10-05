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
import { addPreset, deletePreset } from "../../utils/commonFuncs";

function CellStyle() {
  const [cellStylePresets, setCellStylePresets] = useState([]);
  const selectedStylePreset = usePublicStore(
    (state) => state.selectedStylePreset,
  );
  const setSelectedStylePreset = usePublicStore(
    (state) => state.setSelectedStylePreset,
  );
  const styles = useStyles();

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();
      const sortedPresets = Object.keys(savedPresets).sort((a, b) => {
        const numA = parseInt(a.replace(/\D/g, ""), 10);
        const numB = parseInt(b.replace(/\D/g, ""), 10);

        return numA - numB;
      });

      setCellStylePresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectedStylePreset) {
        setSelectedStylePreset(sortedPresets[0]);
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
    let presetNumbers = [];

    if (cellStylePresets.length > 0) {
      presetNumbers = cellStylePresets.map((preset) =>
        parseInt(preset.replace(/\D/g, ""), 10),
      );
    }

    let lastPresetNum = 1;
    while (presetNumbers.includes(lastPresetNum)) {
      lastPresetNum += 1;
    }

    const newPresetName = `셀 서식${lastPresetNum}`;

    await addPreset("cellStylePresets", newPresetName);

    const savedPresets = await loadPresets();
    const sortedPresets = Object.keys(savedPresets).sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ""), 10);
      const numB = parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    setCellStylePresets(sortedPresets);
    setSelectedStylePreset(newPresetName);
  }

  async function handleDeletePreset() {
    if (!selectedStylePreset) {
      return;
    }

    const selectIndex = cellStylePresets.indexOf(selectedStylePreset);

    await deletePreset("cellStylePresets", selectedStylePreset);

    const savedPresets = Object.keys(await loadPresets());
    const sortedPresets = savedPresets.sort((a, b) => {
      const numA = +parseInt(a.replace(/\D/g, ""), 10);
      const numB = +parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    setCellStylePresets(sortedPresets);
    setSelectedStylePreset(sortedPresets[selectIndex]);
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
          onClick={handleDeletePreset}
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
