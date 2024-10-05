import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import CustomDropdown from "../common/CustomDropdown";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import { addPreset, deletePreset } from "../../utils/commonFuncs";
import {
  saveChartStylePreset,
  loadChartStylePreset,
} from "../../utils/chartStyleFuncs";

function ChartStyle() {
  const [selectedChartPreset, setSelectedChartPreset] = useState("");
  const [chartStylePresets, setChartStylePresets] = useState([]);
  const styles = useStyles();

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();
      const sortedPresets = Object.keys(savedPresets).sort((a, b) => {
        const numA = parseInt(a.replace(/\D/g, ""), 10);
        const numB = parseInt(b.replace(/\D/g, ""), 10);

        return numA - numB;
      });

      setChartStylePresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectedChartPreset) {
        setSelectedChartPreset(sortedPresets[0]);
      }
    }

    fetchPresets();
  }, [selectedChartPreset]);

  async function loadPresets() {
    let presets = await OfficeRuntime.storage.getItem("chartStylePresets");

    if (!presets) {
      presets = {};
    } else {
      presets = JSON.parse(presets);
    }

    return presets;
  }

  async function newPreset() {
    let presetNumbers = [];

    if (chartStylePresets.length > 0) {
      presetNumbers = chartStylePresets.map((preset) =>
        parseInt(preset.replace(/\D/g, ""), 10),
      );
    }

    let lastPresetNum = 1;
    while (presetNumbers.includes(lastPresetNum)) {
      lastPresetNum += 1;
    }

    const newPresetName = `차트 서식${lastPresetNum}`;

    await addPreset("chartStylePresets", newPresetName);

    const savedPresets = await loadPresets();
    const sortedPresets = Object.keys(savedPresets).sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ""), 10);
      const numB = parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    setChartStylePresets(sortedPresets);
    setSelectedChartPreset(newPresetName);
  }

  async function handleDeletePreset() {
    if (!selectedChartPreset) {
      return;
    }

    const selectIndex = chartStylePresets.indexOf(selectedChartPreset);

    await deletePreset("chartStylePresets", selectedChartPreset);

    const savedPresets = Object.keys(await loadPresets());
    const sortedPresets = savedPresets.sort((a, b) => {
      const numA = +parseInt(a.replace(/\D/g, ""), 10);
      const numB = +parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    setChartStylePresets(Object.keys(sortedPresets));
    setSelectedChartPreset(sortedPresets[selectIndex]);
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
          handleValue={(value) => setSelectedChartPreset(value)}
          options={chartStylePresets.map((preset) => ({
            name: preset,
            value: preset,
          }))}
          placeholder="프리셋"
          selectedValue={selectedChartPreset}
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
          onClick={() =>
            saveChartStylePreset("chartStylePresets", selectedChartPreset)
          }
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
        onClick={() =>
          loadChartStylePreset("chartStylePresets", selectedChartPreset)
        }
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default ChartStyle;
