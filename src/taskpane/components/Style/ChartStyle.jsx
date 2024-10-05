import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import PresetHandler from "../../classes/PresetHandler";
import { useStyles } from "../../utils/style";
import CustomDropdown from "../common/CustomDropdown";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import {
  saveChartStylePreset,
  loadChartStylePreset,
} from "../../utils/chartStyleFuncs";

function ChartStyle() {
  const [selectedChartPreset, setSelectedChartPreset] = useState("");
  const [chartStylePresets, setChartStylePresets] = useState([]);
  const styles = useStyles();
  const presets = new PresetHandler("chartStylePresets", "차트 서식");

  useEffect(() => {
    async function fetchPresets() {
      const sortedPresets = await presets.sorting();

      setChartStylePresets(sortedPresets);

      if (sortedPresets.length > 0 && !selectedChartPreset) {
        setSelectedChartPreset(sortedPresets[0]);
      }
    }

    fetchPresets();
  }, [selectedChartPreset]);

  async function newPreset() {
    setSelectedChartPreset(await presets.add(chartStylePresets));
    setChartStylePresets(await presets.sorting());
  }

  async function deletePreset() {
    const selectIndex = chartStylePresets.indexOf(selectedChartPreset);

    setChartStylePresets(await presets.delete(selectedChartPreset));
    setSelectedChartPreset((await presets.sorting())[selectIndex]);
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
          onClick={deletePreset}
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
