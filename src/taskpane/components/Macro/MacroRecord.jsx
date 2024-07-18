import { useStyles } from "../../utils/style";
import useStore from "../../utils/store";
import { SaveIcon, EditIcon, RecordStart, RecordStop } from "../../utils/icons";
import MACRO_PRESETS from "../../Presets/MacroPreset";
import CustomDropdown from "../common/CustomDropdown";

function MacroRecord() {
  const styles = useStyles();
  const { isRecording, setIsRecording } = useStore();

  function controlRecoding() {
    setIsRecording(!isRecording);
  }

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
        <span className="h-6">매크로 녹화</span>
        <button
          onClick={controlRecoding}
          className={styles.buttons}
          aria-label="edit"
        >
          {isRecording ? <RecordStop /> : <RecordStart color="red" />}
        </button>
      </div>
    </>
  );
}

export default MacroRecord;
