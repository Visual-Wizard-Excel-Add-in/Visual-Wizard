import { Switch } from "@fluentui/react-components";

function ValidateTest() {
  const lastCell = "D8";

  return (
    <div>
      <Switch label="에러 셀 검사" onChange={() => {}} />
      <p>
        사용중인 마지막 셀 영역:&nbsp;
        <span className="font-bold">{lastCell}</span>
      </p>
    </div>
  );
}

export default ValidateTest;
