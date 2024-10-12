import { useEffect } from "react";
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

  useEffect(() => {
    const timerList = [];

    function setRemoveMessageTimer() {
      messageList.forEach((message) => {
        const timer = setTimeout(() => removeMessage(message.id), 2500);

        timerList.push(timer);
      });
    }

    setRemoveMessageTimer();

    return () => {
      timerList.forEach((timer) => clearTimeout(timer));
    };
  }, [messageList]);

  return (
    <div className="flex justify-center mt-1 whitespace-pre-wrap">
      <MessageBarGroup animate="both" className={styles.messageBarGroup}>
        {messageList.map(({ id, message }) => (
          <MessageBar key={id} intent={message.type}>
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
