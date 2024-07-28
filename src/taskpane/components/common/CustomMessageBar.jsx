import { useEffect, useId } from "react";
import { DismissRegular } from "@fluentui/react-icons";
import {
  MessageBar,
  MessageBarActions,
  MessageBarTitle,
  MessageBarBody,
  MessageBarGroup,
  Button,
} from "@fluentui/react-components";

import useStore from "../../utils/store";
import { useStyles } from "../../utils/style";

function CustomMessageBar() {
  const { messageList, removeMessage } = useStore();
  const styles = useStyles();
  const messageId = useId();

  useEffect(() => {
    let removeTimer;

    messageList.forEach((message) => {
      removeTimer = setTimeout(() => removeMessage(message.id), 2500);
    });

    return () => clearTimeout(removeTimer);
  }, [messageList]);

  return (
    <MessageBarGroup animate="both" className={styles.messageBarGroup}>
      {messageList.map(({ id, message }) => (
        <MessageBar key={messageId} intent={message.type}>
          <MessageBarBody>
            <MessageBarTitle>{message.title}</MessageBarTitle>
            {message.body}
          </MessageBarBody>
          <MessageBarActions
            containerAction={
              <Button
                onClick={() => removeMessage(id)}
                aria-label="dismiss"
                appearance="transparent"
                icon={<DismissRegular />}
              />
            }
          />
        </MessageBar>
      ))}
    </MessageBarGroup>
  );
}

export default CustomMessageBar;
