import { useState } from "react";
import { Button, Divider } from "@fluentui/react-components";

import CustomDropdown from "../common/CustomDropdown";
import executeFunction from "../../utils/extractFileFunc";
import ShareNoticeBar from "./ShareNoticeBar";

function Extraction() {
  const [dataLocation, setDataLocation] = useState("선택 영역");
  const [isLoading, setIsLoading] = useState(false);
  const [isShowNoticeBar, setIsShowNoticeBar] = useState(true);

  const dataLocationOptions = [
    { name: "선택 영역", value: "selectRange" },
    { name: "현재 시트", value: "currentSheet" },
  ];

  const handleSaveData = async () => {
    setIsLoading(true);
    await executeFunction(dataLocation);
    setIsLoading(false);
  };

  return (
    <>
      <div className="flex justify-between space-x-5">
        {isShowNoticeBar && (
          <ShareNoticeBar setIsShowNoticeBar={setIsShowNoticeBar} />
        )}
        <p>저장할 자료 위치</p>
        <CustomDropdown
          options={dataLocationOptions}
          handleValue={(value) => setDataLocation(value)}
          selectedValue={dataLocation}
          placeholder="데이터 선택"
        />
      </div>
      <Divider className="my-2" appearance="strong" />
      <div className="flex justify-center">
        <Button onClick={handleSaveData}>저장</Button>
        {isLoading && (
          <div>
            <span>저장 중...</span>
          </div>
        )}
      </div>
    </>
  );
}

export default Extraction;
