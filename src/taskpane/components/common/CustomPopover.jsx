import {
  Popover,
  PopoverSurface,
  PopoverTrigger,
} from "@fluentui/react-components";

function CustomPopover({ position, PopoverContents, triggerContents }) {
  return (
    <Popover withArrow="true" positioning={position}>
      <PopoverTrigger disableButtonEnhancement>
        {triggerContents}
      </PopoverTrigger>

      <PopoverSurface tabIndex={-1} className="whitespace-pre-wrap max-w-48">
        {PopoverContents}
      </PopoverSurface>
    </Popover>
  );
}

export default CustomPopover;
