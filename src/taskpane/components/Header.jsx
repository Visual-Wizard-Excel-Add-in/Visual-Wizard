import { Tab, TabList } from "@fluentui/react-components";

import useStore from "../utils/store";
import { useStyles } from "../utils/style";

function Header() {
  const styles = useStyles();
  const { category, setCategory } = useStore();

  function selectCategory(event, data) {
    setCategory(data.value);
  }

  return (
    <div className={styles.list}>
      <TabList
        defaultSelectedValue="Fomula"
        appearance="subtle"
        onTabSelect={selectCategory}
      >
        <Tab value="Fomula" className="h-6">
          수식
        </Tab>
        <Tab value="Style" className="h-6">
          서식
        </Tab>
        <Tab value="Macro" className="h-6">
          매크로
        </Tab>
        <Tab value="Validate" className="h-6">
          유효성
        </Tab>
        <Tab value="Share" className="h-6">
          공유하기
        </Tab>
      </TabList>
    </div>
  );
}

export default Header;
