import { Field, Input, Textarea, Button } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import CustomDropdown from "../common/CustomDropdown";

function Mail() {
  const styles = useStyles();
  const dataLocation = [
    {
      name: "선택 영역",
    },
    {
      name: "현재 워크시트",
    },
    {
      name: "현재 통합문서",
    },
  ];
  const toInfo = ["받는사람", "참조", "제목"];

  return (
    <>
      <div className="flex items-center justify-between space-x-5">
        <p>보낼 자료 위치</p>
        <CustomDropdown options={dataLocation} placeholder="데이터" />
      </div>
      <hr className={styles.border} />
      {toInfo.map((info) => (
        <div className="flex items-center justify-between space-x-5" key={info}>
          <p>{info}</p>
          <Input />
        </div>
      ))}
      <Field label="본문">
        <Textarea className="h-32" />
      </Field>
      <div className="flex justify-center">
        <Button>전송</Button>
      </div>
    </>
  );
}

export default Mail;
