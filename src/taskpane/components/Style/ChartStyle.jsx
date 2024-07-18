import { useStyles } from "../../utils/style";
import { SaveIcon, EditIcon } from "../../utils/icons";
import CHART_STYLE_PRESETS from "../../Presets/CellStylePresets";
import CustomDropdown from "../common/CustomDropdown";

function ChartStyle() {
  const styles = useStyles();

  return (
    <div className="flex items-center justify-between space-x-5">
      <span>서식 프리셋</span>
      <div className="flex items-center space-x-2">
        <CustomDropdown options={CHART_STYLE_PRESETS} placeholder="프리셋" />
        <button className={styles.buttons} aria-label="save">
          <EditIcon />
        </button>
        <button className={styles.buttons} aria-label="edit">
          <SaveIcon />
        </button>
      </div>
    </div>
  );
}

export default ChartStyle;
