import { addPreset, deletePreset } from "../utils/commonFuncs";

class PresetHandler {
  constructor(listName, presetName) {
    this.listName = listName;
    this.presetName = presetName;
    this.presets = null;
  }

  async loadStorage() {
    const preset = await OfficeRuntime.storage.getItem(this.listName);

    return preset;
  }

  async loadPresets() {
    this.presets = await this.loadStorage(this.listName);

    if (!this.presets) {
      this.presets = {};
    } else {
      this.presets = JSON.parse(this.presets);
    }

    return this.presets;
  }

  async add(presetList) {
    let presetNumbers = [];

    if (presetList.length > 0) {
      presetNumbers = presetList.map((preset) =>
        parseInt(preset.replace(/\D/g, ""), 10),
      );
    }

    let lastPresetNum = 1;

    while (presetNumbers.includes(lastPresetNum)) {
      lastPresetNum += 1;
    }

    const newPresetName = `${this.presetName}${lastPresetNum}`;

    await addPreset(this.listName, newPresetName);

    return newPresetName;
  }

  async sorting() {
    const sortedPresets = Object.keys(await this.loadPresets()).sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ""), 10);
      const numB = parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    return sortedPresets;
  }

  async delete(selectPreset) {
    await deletePreset(this.listName, selectPreset);

    const sortedPresets = await this.sorting();

    return sortedPresets;
  }
}

export default PresetHandler;
