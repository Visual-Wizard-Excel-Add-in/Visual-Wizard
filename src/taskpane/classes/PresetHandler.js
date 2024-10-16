class PresetHandler {
  constructor(listName, presetName) {
    this.listName = listName;
    this.presetName = presetName;
    this.presets = this.load();
  }

  async load() {
    const existingData = await loadStorage(this.listName);

    if (!this.presets) {
      this.presets = {};
    } else {
      this.presets = JSON.parse(existingData);
    }

    return this.presets;
  }

  async sort() {
    const result = Object.keys(await this.load()).sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ""), 10);
      const numB = parseInt(b.replace(/\D/g, ""), 10);

      return numA - numB;
    });

    return result;
  }

  async add(presetList) {
    let existOrderNums = [];

    if (presetList.length > 0) {
      existOrderNums = presetList.map((preset) =>
        parseInt(preset.replace(/\D/g, ""), 10),
      );
    }

    let newOrderNum = 1;

    while (existOrderNums.includes(newOrderNum)) {
      newOrderNum += 1;
    }

    const newPresetName = `${this.presetName}${newOrderNum}`;

    await addPreset(this.listName, newPresetName);

    return newPresetName;
  }

  async delete(selectPreset) {
    await deletePreset(this.listName, selectPreset);

    const sortedPresets = await this.sort();

    return sortedPresets;
  }
}

export default PresetHandler;

async function addPreset(presetCategory, presetName) {
  const savePreset = (await loadStorage(presetCategory))
    ? JSON.parse(await loadStorage(presetCategory))
    : {};

  savePreset[presetName] = {};

  await OfficeRuntime.storage.setItem(
    presetCategory,
    JSON.stringify(savePreset),
  );
}

async function deletePreset(presetCategory, presetName) {
  const currentPresets = JSON.parse(
    await OfficeRuntime.storage.getItem(presetCategory),
  );

  if (currentPresets) {
    delete currentPresets[presetName];

    await OfficeRuntime.storage.setItem(
      presetCategory,
      JSON.stringify(currentPresets),
    );
  }
}

async function loadStorage(presetCategory) {
  const preset = await OfficeRuntime.storage.getItem(presetCategory);

  return preset;
}
