import { Dropdown, Option, useId } from "@fluentui/react-components";

function CustomDropdown({ options, placeholder }) {
  const selectId = useId();

  return (
    <Dropdown id={selectId} className="w-24 min-w-0" placeholder={placeholder}>
      {options.map((option, index) => (
        <Option key={{ option } + { index }} className="!w-24">
          {option.name}
        </Option>
      ))}
    </Dropdown>
  );
}

export default CustomDropdown;
