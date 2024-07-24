import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import CustomDropdown from "../common/CustomDropdown";
import { addPreset, deletePreset } from "../../utils/cellCommonUtils";
import {
  saveChartStylePreset,
  loadChartStylePreset,
} from "../../utils/cellStyleFunc";

function ChartStyle() {
  const styles = useStyles();
  const [selectedChartPreset, setSelectedChartPreset] = useState("");
  const [chartStylePresets, setChartStylePresets] = useState([]);

  useEffect(() => {
    async function fetchPresets() {
      const savedPresets = await loadPresets();

      setChartStylePresets(Object.keys(savedPresets));
    }

    fetchPresets();
  }, []);

  async function loadPresets() {
    return Excel.run(async () => {
      let presets = Office.context.document.settings.get("chartStylePresets");

      if (!presets) {
        presets = {};
      } else {
        presets = JSON.parse(presets);
      }

      return presets;
    });
  }

  async function newPreset() {
    await addPreset(
      "chartStylePresets",
      `차트 서식${chartStylePresets.length + 1}`,
    );

    const savedPresets = await loadPresets();

    setChartStylePresets(Object.keys(savedPresets));
  }

  async function handleDeletePreset() {
    if (!selectedChartPreset) {
      return;
    }

    await deletePreset("chartStylePresets", selectedChartPreset);

    const savedPresets = await loadPresets();

    setSelectedChartPreset("");
    setChartStylePresets(Object.keys(savedPresets));
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
            aria-label="delete"
          >
            <DeleteIcon />
          </button>
          <button
            onClick={() => saveChartStylePreset(selectedChartPreset)}
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
        onClick={() => loadChartStylePreset(selectedChartPreset)}
        size="small"
      >
        적용
      </Button>
    </>
  );
}

export default ChartStyle;
