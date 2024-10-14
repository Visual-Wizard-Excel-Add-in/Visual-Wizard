import {
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Button,
} from "@fluentui/react-components";

function CustomPopover({ position, triggerContents, PopoverContents }) {
  return (
    <Popover withArrow="true" positioning={position}>
      <PopoverTrigger disableButtonEnhancement>
        <Button>{triggerContents}</Button>
      </PopoverTrigger>

      <PopoverSurface tabIndex={-1} className="whitespace-pre-wrap max-w-48">
        {PopoverContents}
      </PopoverSurface>
    </Popover>
  );
}

export default CustomPopover;
