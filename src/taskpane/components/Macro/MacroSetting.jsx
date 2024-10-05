import { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";
import ActionDetails from "./ActionDetails";
import MacroNoticeBar from "./MacroNoticeBar";

import usePublicStore from "../../store/publicStore";

function MacroSetting() {
  const [storedMacro, setStoredMacro] = useState([]);
  const [modifiedActions, setModifiedActions] = useState({});
  const [isShowNoticeBar, setIsShowNoticeBar] = useState(true);
  const selectMacroPreset = usePublicStore((state) => state.selectMacroPreset);

  useEffect(() => {
    async function fetchMacroPresets() {
      const data = await OfficeRuntime.storage.getItem("allMacroPresets");

      if (data) {
        const parsedData = JSON.parse(data);

        setStoredMacro(parsedData[selectMacroPreset]?.actions || []);
      } else {
        setStoredMacro([]);
      }
    }

    fetchMacroPresets();
  }, [selectMacroPreset]);

  const applyChanges = async () => {
    const updatedActions = storedMacro.map((action, index) => {
      const modifiedAction = modifiedActions[index] || {};

      return {
        ...action,
        ...modifiedAction,
        details: {
          ...action.details,
          ...modifiedAction.details,
        },
      };
    });

    setStoredMacro(updatedActions);

    let allMacroPresets =
      await OfficeRuntime.storage.getItem("allMacroPresets");
    allMacroPresets = allMacroPresets ? JSON.parse(allMacroPresets) : {};

    if (!allMacroPresets[selectMacroPreset]) {
      allMacroPresets[selectMacroPreset] = { actions: [] };
    }

    allMacroPresets[selectMacroPreset].actions = updatedActions;

    await OfficeRuntime.storage.setItem(
      "allMacroPresets",
      JSON.stringify(allMacroPresets),
    );
  };

  return (
    <>
      <div className="flex justify-center">
        {isShowNoticeBar && (
          <MacroNoticeBar setIsShowNoticeBar={setIsShowNoticeBar} />
        )}
      </div>
      <div className="flex justify-between">
        선택한 프리셋: {selectMacroPreset}
        <Button as="button" onClick={applyChanges} size="small">
          변경사항 적용
        </Button>
      </div>
      {storedMacro.map((action, index) => (
        <ActionDetails
          key={`${action.type}-${index + 1}`}
          action={action}
          index={index}
          setModifiedActions={setModifiedActions}
        />
      ))}
    </>
  );
}

export default MacroSetting;
