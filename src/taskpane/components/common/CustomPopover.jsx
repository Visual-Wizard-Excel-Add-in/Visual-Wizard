import {
  Popover,
  PopoverSurface,
  PopoverTrigger,
} from "@fluentui/react-components";

function CustomPopover({ position, PopoverContents, triggerContents }) {
  return (
    <Popover positioning={position}>
      <PopoverTrigger disableButtonEnhancement>
        {triggerContents}
      </PopoverTrigger>

      <PopoverSurface tabIndex={-1} className="whitespace-pre-wrap">
        {PopoverContents}
      </PopoverSurface>
    </Popover>
  );
}

export default CustomPopover;
