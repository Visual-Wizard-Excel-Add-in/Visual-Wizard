import { useState, useEffect, useCallback } from "react";
import { Button, Input, Divider } from "@fluentui/react-components";

import useStore from "../../utils/store";
import CustomDropdown from "../common/CustomDropdown";
import chartTypeList from "../../utils/chartTypeList";
import {
  getChartTypeInKorean,
  getChartTypeInEnglish,
} from "../../utils/cellCommonUtils";

function MacroSetting() {
  const [storedMacro, setStoredMacro] = useState([]);
  const [modifiedActions, setModifiedActions] = useState({});
  const [selectChartType, setSelectChartType] = useState("");
  const selectMacroPreset = useStore((state) => state.selectMacroPreset);

  useEffect(() => {
    async function fetchMacroPresets() {
      const data = await OfficeRuntime.storage.getItem("allMacroPresets");

      if (data) {
        const parsedData = JSON.parse(data);

        setStoredMacro(parsedData[selectMacroPreset].actions || []);
      } else {
        setStoredMacro([]);
      }
    }

    fetchMacroPresets();
  }, [selectMacroPreset]);

  function getMergedRange(action) {
    let mergedRange = "";

    if (action.dataRange[0].includes(":")) {
      mergedRange = `${action.dataRange[0].split(":")[0]}:${action.dataRange[action.dataRange.length - 1].split(":")[1]}`;
    } else {
      mergedRange = `${action.dataRange[0]}:${action.dataRange[action.dataRange.length - 1]}`;
    }

    return mergedRange;
  }

  function renderActionType(action, index) {
    let actionContent = null;

    switch (action.type) {
      case "WorksheetChanged":
        actionContent = (
          <div
            key={`sheetChanged-${action.address}-${action.details.value}`}
            className="mb-3"
          >
            <p className="mb-2 text-base font-bold bg-green-500 bg-opacity-20">
              {index + 1}. 셀 내용 변경
            </p>
            <p>
              셀 주소:&nbsp;
              <Input
                onChange={(e) => handleChange(index, "address", e.target.value)}
                placeholder={`${action.address ? action.address : ""}`}
              />
            </p>
            <p>
              입력값:&nbsp;&nbsp;
              <Input
                onChange={(e) =>
                  handleChange(index, "details.value", e.target.value)
                }
                placeholder={`현재값: ${action.details.value ? action.details.value : ""}`}
              />
            </p>
          </div>
        );
        break;

      case "WorksheetFormatChanged":
        actionContent = (
          <span
            key={`sheetFormatChanged-${action.address}-${index}`}
            className="text-base font-bold bg-green-500 bg-opacity-20"
          >
            {index + 1}. 셀 서식 변경
          </span>
        );
        break;

      case "TableChanged":
        actionContent = (
          <span
            key={`tableChanged-${action.tableId}-${index}`}
            className="text-base font-bold bg-green-500 bg-opacity-20"
          >
            {index + 1}. 테이블 변경
          </span>
        );
        break;

      case "ChartAdded":
        actionContent = (
          <div key={`chartAdded-${action.chartId}-${index}`}>
            <p className="mb-2 text-base font-bold bg-green-500 bg-opacity-20">
              {index + 1}. 차트 추가
            </p>
            <div>
              차트 타입:&nbsp;{" "}
              <CustomDropdown
                handleValue={(value) => {
                  setSelectChartType(value);
                  handleChange(
                    index,
                    "chartType",
                    getChartTypeInEnglish(value),
                  );
                }}
                options={chartTypeList.map((chartType) => ({
                  name: chartType.name,
                  value: chartType.value,
                  label: chartType.label,
                }))}
                placeholder={getChartTypeInKorean(action.chartType)}
                selectedValue={selectChartType}
              />
            </div>
            <div className="my-2">
              데이터 범위:&nbsp;
              <Input
                onChange={(e) =>
                  handleChange(index, "dataRange", e.target.value)
                }
                placeholder={`현재값: ${getMergedRange(action)}`}
              />
            </div>
          </div>
        );
        break;

      case "TableAdded":
        actionContent = (
          <div key={`TableAdded-${action.tableId}-${index}`}>
            <div className="mb-2 text-base font-bold bg-green-500 bg-opacity-20">
              {index + 1}. 표 추가
            </div>
            <div className="my-2">
              데이터 범위:&nbsp;
              <Input
                onChange={(e) => handleChange(index, "address", e.target.value)}
                placeholder={action.address}
              />
            </div>
          </div>
        );
        break;

      default:
        return "지원하지 않는 형식의 기록입니다.";
    }

    return (
      <div>
        {actionContent}
        <Divider className="my-2" appearance="strong" />
      </div>
    );
  }

  function handleChange(index, fieldPath, value) {
    setModifiedActions((prev) => {
      const [mainField, subField] = fieldPath.split(".");
      const updatedAction = {
        ...prev[index],
        [mainField]: { ...prev[index]?.[mainField] },
      };

      if (subField) {
        updatedAction[mainField][subField] = value;
      } else {
        updatedAction[mainField] = value;
      }

      return {
        ...prev,
        [index]: updatedAction,
      };
    });
  }

  const applyChanges = useCallback(async () => {
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
  }, [storedMacro, modifiedActions]);

  return (
    <>
      <div className="flex justify-between">
        선택한 프리셋: {selectMacroPreset}
        <Button as="button" onClick={applyChanges} size="small">
          변경사항 적용
        </Button>
      </div>
      {storedMacro.map((action, index) => renderActionType(action, index))}
    </>
  );
}

export default MacroSetting;
