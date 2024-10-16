import { useEffect, useState } from "react";
import { Button } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import { SaveIcon, DeleteIcon, PlusIcon } from "../../utils/icons";
import { useStyles } from "../../utils/style";
import PresetHandler from "../../classes/PresetHandler";
import {
  copyChartStylePreset,
  pasteChartStylePreset,
} from "../../utils/chartStyleFuncs";

function ChartStyle() {
  const [selectedChartPreset, setSelectedChartPreset] = useState("");
  const [chartStylePresets, setChartStylePresets] = useState([]);
  const styles = useStyles();
  const presets = new PresetHandler("chartStylePresets", "차트 서식");

  useEffect(() => {
    fetchPresets();

    async function fetchPresets() {
      setChartStylePresets(await presets.sort());

      if ((await presets.sort()).length > 0 && !selectedChartPreset) {
        setSelectedChartPreset((await presets.sort())[0]);
      }
    }
  }, [selectedChartPreset]);

  async function newPresetHandler() {
    const newPreset = await presets.add(chartStylePresets);
    const newPresetList = await presets.sort();

    setSelectedChartPreset(newPreset);
    setChartStylePresets(newPresetList);
  }

  async function deletePresetHandler() {
    const forwardPreset = (await presets.sort())[
      chartStylePresets.indexOf(selectedChartPreset)
    ];
    const newPresetList = await presets.delete(selectedChartPreset);

    setChartStylePresets(newPresetList);
    setSelectedChartPreset(forwardPreset);
  }

  return (
    <div className="flex items-center justify-between space-x-5">
      <div className="flex items-center w-8/12 space-x-2">
        <button
          onClick={newPresetHandler}
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
          onClick={deletePresetHandler}
          className={styles.buttons}
          aria-label="delete preset"
          type="button"
        >
          <DeleteIcon />
        </button>
        <button
          onClick={() =>
            copyChartStylePreset("chartStylePresets", selectedChartPreset)
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
          pasteChartStylePreset("chartStylePresets", selectedChartPreset)
        }
        size="small"
      >
        적용
      </Button>
    </div>
  );
}

export default ChartStyle;
