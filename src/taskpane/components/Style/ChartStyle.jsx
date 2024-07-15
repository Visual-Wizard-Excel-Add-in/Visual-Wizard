import { Dropdown, Option, useId } from "@fluentui/react-components";
import { useStyles } from "../../utils/style";
import CHART_STYLE_PRESETS from "../../utils/CellStylePresets";
import { SaveIcon, EditIcon } from "../../utils/icons";

function ChartStyle() {
  const selectId = useId();
  const styles = useStyles();

  return (
    <div className="flex items-center justify-between space-x-5">
      <span>서식 프리셋</span>
      <div className="flex items-center space-x-2">
        <Dropdown id={selectId} className="w-24 min-w-0" placeholder="프리셋">
          {CHART_STYLE_PRESETS.map((preset) => (
            <Option key={preset.num} className="!w-24">
              {preset.num}
            </Option>
          ))}
        </Dropdown>
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
