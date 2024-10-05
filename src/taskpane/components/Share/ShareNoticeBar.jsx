import {
  Button,
  Link,
  MessageBar,
  MessageBarActions,
  MessageBarTitle,
  MessageBarBody,
  MessageBarGroup,
} from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";

import { useStyles } from "../../utils/style";

function ShareNoticeBar({ setIsShowNoticeBar }) {
  const styles = useStyles();

  return (
    <MessageBarGroup animate="both" className={`${styles.messageBarGroup}`}>
      <MessageBar intent="warning">
        <MessageBarBody>
          <MessageBarTitle>주의:</MessageBarTitle>
          추출하기 이용을 위해선
          <br />
          먼저&nbsp;
          <Link
            className={styles.fontBolder}
            appearance="inline"
            href="https://fair-gram-629.notion.site/b201f6a19fec4e1c8f5c1f83aaf0a8ab?pvs=4"
          >
            이곳
          </Link>
          을 방문해주세요!
        </MessageBarBody>
        <MessageBarActions
          containerAction={
            <Button
              onClick={() => setIsShowNoticeBar(false)}
              aria-label="close message"
              appearance="transparent"
              icon={<DismissRegular />}
            />
          }
        />
      </MessageBar>
    </MessageBarGroup>
  );
}

export default ShareNoticeBar;
