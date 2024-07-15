import { Button, Input } from "@fluentui/react-components";

import { useStyles } from "../../utils/style";
import { SaveIcon, EditIcon } from "../../utils/icons";
import MACRO_PRESETS from "../../utils/MacroPreset";
import CustomDropdown from "../common/CustomDropdown";

function MacroSetting() {
  const styles = useStyles();

  return (
    <>
      <div className="flex items-center justify-between space-x-5">
        <span>매크로</span>
        <div className="flex items-center space-x-2">
          <CustomDropdown options={MACRO_PRESETS} placeholder="매크로" />
          <button className={styles.buttons} aria-label="save">
            <EditIcon />
          </button>
          <button className={styles.buttons} aria-label="edit">
            <SaveIcon />
          </button>
        </div>
      </div>
      <div className="flex items-center justify-between space-x-5">
        <span>버튼 생성</span>
        <div className="space-x-2">
          <Button
            size="small"
            icon={
              <img
                className="border-2 border-slate-300"
                src="../../../../assets/macroButton.png"
                alt="macro button"
              />
            }
          />
          <Button
            size="small"
            icon={
              <img
                className="border-2 border-slate-300"
                src="../../../../assets/roterShartButton.png"
                alt="macro button"
              />
            }
          />
        </div>
      </div>
      <div className="flex items-center justify-between space-x-5">
        <span>버튼 매크로 할당</span>
        <Button size="small">적용</Button>
      </div>
      <div className="flex items-center justify-between space-x-5">
        <span>바로가기 키 지정</span>
        <div className="flex items-center">
          <span className={styles.blurText}>option+cmd+</span>
          <Input className={styles.macroKey} placeholder="key" />
          <button className={`${styles.buttons} ml-2`} aria-label="edit">
            <SaveIcon />
          </button>
        </div>
      </div>
    </>
  );
}

export default MacroSetting;
