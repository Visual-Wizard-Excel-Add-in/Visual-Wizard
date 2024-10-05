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

import usePublicStore from "../../store/publicStore";
import { useStyles } from "../../utils/style";

function CustomMessageBar() {
  const messageList = usePublicStore((state) => state.messageList);
  const removeMessage = usePublicStore((state) => state.removeMessage);
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
    <div className="flex justify-center mt-1 whitespace-pre-wrap">
      <MessageBarGroup animate="both" className={styles.messageBarGroup}>
        {messageList.map(({ id, message }) => (
          <MessageBar key={messageId} intent={message.type}>
            <MessageBarBody className="whitespace-pre-wrap">
              <MessageBarTitle>{message.title}</MessageBarTitle>
              {message.body}
            </MessageBarBody>
            <MessageBarActions
              containerAction={
                <Button
                  onClick={() => removeMessage(id)}
                  aria-label="close message"
                  appearance="transparent"
                  icon={<DismissRegular />}
                />
              }
            />
          </MessageBar>
        ))}
      </MessageBarGroup>
    </div>
  );
}

export default CustomMessageBar;
