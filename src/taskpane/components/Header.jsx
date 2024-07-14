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
        <Tab value="Fomula">수식</Tab>
        <Tab value="Style">서식</Tab>
        <Tab value="Macro">매크로</Tab>
        <Tab value="Validate">유효성</Tab>
        <Tab value="Share">공유하기</Tab>
      </TabList>
    </div>
  );
}

export default Header;
