import {
  Button,
  MessageBar,
  MessageBarActions,
  MessageBarTitle,
  MessageBarBody,
  MessageBarGroup,
} from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";

import { useStyles } from "../../utils/style";

function NoticeBar({ setIsShowNoticeBar }) {
  const styles = useStyles();

  return (
    <MessageBarGroup animate="both" className={`${styles.messageBarGroup}`}>
      <MessageBar intent="warning">
        <MessageBarBody>
          <MessageBarTitle>안내:</MessageBarTitle>
          매크로 기록 가능 종류.
          <br />셀 입력, 셀 서식 변경, 차트 추가, 표 추가
        </MessageBarBody>
        <MessageBarActions
          containerAction={
            <Button
              onClick={() => setIsShowNoticeBar(false)}
              aria-label="dismiss"
              appearance="transparent"
              icon={<DismissRegular />}
            />
          }
        />
      </MessageBar>
    </MessageBarGroup>
  );
}

export default NoticeBar;
