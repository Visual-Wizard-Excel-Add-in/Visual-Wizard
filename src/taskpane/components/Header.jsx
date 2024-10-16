import { Tab, TabList } from "@fluentui/react-components";

import usePublicStore from "../store/publicStore";
import { useStyles } from "../utils/style";

function Header() {
  const styles = useStyles();
  const setCategory = usePublicStore((state) => state.setCategory);
  const setOpenTab = usePublicStore((state) => state.setOpenTab);

  const selectCategory = (event, data) => {
    setCategory(data.value);
    setOpenTab([]);
  };

  const TABS = [
    { name: "수식", value: "Formula" },
    { name: "서식", value: "Style" },
    { name: "매크로", value: "Macro" },
    { name: "유효성", value: "Validity" },
    { name: "공유하기", value: "Share" },
  ];

  return (
    <div className={`sticky top-0 z-10 ${styles.list}`}>
      <TabList
        defaultSelectedValue="Formula"
        appearance="subtle"
        onTabSelect={selectCategory}
      >
        {TABS.map((tab) => (
          <Tab value={tab.value} key={tab.value} className="h-6">
            {tab.name}
          </Tab>
        ))}
      </TabList>
    </div>
  );
}

export default Header;
