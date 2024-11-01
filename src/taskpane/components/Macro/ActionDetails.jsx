import { useState } from "react";
import { Input, Divider } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import {
  CHART_TYPE_LIST,
  translateChartTypeKOR,
  translateChartTypeENG,
} from "../../utils/chartTypeUtils";

function ActionDetails({ action, index, setModifiedActions }) {
  return (
    <div key={`${action.type}-${index}`}>
      <ActionContents
        action={action}
        index={index}
        setModifiedActions={setModifiedActions}
      />
      <Divider className="my-2" appearance="strong" />
    </div>
  );
}

export default ActionDetails;

function ActionContents({ action, index, setModifiedActions }) {
  const [selectChartType, setSelectChartType] = useState("");

  function getMergedRange(act) {
    let mergedRange = "";

    if (act.dataRange[0].includes(":")) {
      mergedRange = `${act.dataRange[0].split(":")[0]}:${act.dataRange[act.dataRange.length - 1].split(":")[1]}`;
    } else {
      mergedRange = `${act.dataRange[0]}:${act.dataRange[act.dataRange.length - 1]}`;
    }

    return mergedRange;
  }

  function handleChange(order, fieldPath, value) {
    setModifiedActions((prev) => {
      const [mainField, subField] = fieldPath.split(".");
      const updatedAction = {
        ...prev[order],
        [mainField]: { ...prev[order]?.[mainField] },
      };

      if (subField) {
        updatedAction[mainField][subField] = value;
      } else {
        updatedAction[mainField] = value;
      }

      return {
        ...prev,
        [order]: updatedAction,
      };
    });
  }

  switch (action.type) {
    case "WorksheetChanged":
      return (
        <div className="mb-3">
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

    case "WorksheetFormatChanged":
      return (
        <span className="text-base font-bold bg-green-500 bg-opacity-20">
          {index + 1}. 셀 서식 변경
        </span>
      );

    case "TableChanged":
      return (
        <span className="text-base font-bold bg-green-500 bg-opacity-20">
          {index + 1}. 테이블 변경
        </span>
      );

    case "ChartAdded":
      return (
        <div>
          <p className="mb-2 text-base font-bold bg-green-500 bg-opacity-20">
            {index + 1}. 차트 추가
          </p>
          <div>
            차트 타입:&nbsp;{" "}
            <CustomDropdown
              handleValue={(value) => {
                setSelectChartType(value);
                handleChange(index, "chartType", translateChartTypeENG(value));
              }}
              options={CHART_TYPE_LIST.map((chartType) => ({
                name: chartType.name,
                value: chartType.value,
                label: chartType.label,
              }))}
              placeholder={translateChartTypeKOR(action.chartType)}
              selectedValue={selectChartType}
            />
          </div>
          <div className="my-2">
            데이터 범위:&nbsp;
            <Input
              onChange={(e) => handleChange(index, "dataRange", e.target.value)}
              placeholder={`현재값: ${getMergedRange(action)}`}
            />
          </div>
        </div>
      );

    case "TableAdded":
      return (
        <div>
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

    default:
      return <p>지원하지 않는 형식의 기록입니다.</p>;
  }
}
