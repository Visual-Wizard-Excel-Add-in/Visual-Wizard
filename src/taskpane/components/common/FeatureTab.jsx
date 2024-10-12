import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Card,
} from "@fluentui/react-components";

import { useStyles } from "../../utils/style";

function FeatureTab({
  order,
  featureName,
  openTab,
  handleToggle,
  featureContents: FeatureContents,
  currentFormula,
}) {
  const styles = useStyles();

  function accordionBackgroundColor() {
    if (openTab && openTab.includes(order)) {
      return styles.openedAccordion;
    }

    return styles.accordion;
  }

  return (
    <Accordion openItems={openTab} onToggle={handleToggle} multiple collapsible>
      <AccordionItem value={order}>
        <AccordionHeader className={`${accordionBackgroundColor()} h-7`}>
          {featureName}
        </AccordionHeader>
        <AccordionPanel>
          <Card appearance="subtle" className={styles.card}>
            <FeatureContents currentFormula={currentFormula} />
          </Card>
        </AccordionPanel>
      </AccordionItem>
    </Accordion>
  );
}

export default FeatureTab;
