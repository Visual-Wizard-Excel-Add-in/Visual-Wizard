import { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";
import ActionDetails from "./ActionDetails";
import MacroNoticeBar from "./MacroNoticeBar";

import useTotalStore from "../../store/useTotalStore";
import { popUpMessage } from "../../utils/commonFuncs";

function MacroSetting() {
  const [storedMacro, setStoredMacro] = useState([]);
  const [modifiedActions, setModifiedActions] = useState({});
  const [isShowNoticeBar, setIsShowNoticeBar] = useState(true);
  const selectMacroPreset = useTotalStore((state) => state.selectMacroPreset);

  useEffect(() => {
    async function fetchMacroPresets() {
      const loadedData = JSON.parse(
        await OfficeRuntime.storage.getItem("allMacroPresets"),
      );

      if (loadedData) {
        setStoredMacro(loadedData[selectMacroPreset]?.actions || []);
      } else {
        setStoredMacro([]);
      }
    }

    fetchMacroPresets();
  }, [selectMacroPreset]);

  const applyChanges = async () => {
    const updatedActions = storedMacro.map((action, index) =>
      modifyChanges(index, action),
    );
    const savedPresets =
      JSON.parse(await OfficeRuntime.storage.getItem("allMacroPresets")) || {};

    savedPresets[selectMacroPreset].actions = updatedActions;

    setStoredMacro(updatedActions);

    await OfficeRuntime.storage.setItem(
      "allMacroPresets",
      JSON.stringify(savedPresets),
    );

    popUpMessage("saveSuccess", "변경사항이 적용됐습니다");
  };

  function modifyChanges(index, action) {
    const modifiedAction = modifiedActions[index] || {};

    return {
      ...action,
      ...modifiedAction,
      details: {
        ...action.details,
        ...modifiedAction.details,
      },
    };
  }

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
