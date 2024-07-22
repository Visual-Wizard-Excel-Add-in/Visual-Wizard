import { Dropdown, Option, useId } from "@fluentui/react-components";

function CustomDropdown({ handleValue, options, placeholder, selectedValue }) {
  const selectId = useId();

  function handleChange(event, value) {
    event.stopPropagation();

    handleValue(value.optionText);
  }

  return (
    <Dropdown
      id={selectId}
      placeholder={placeholder}
      onOptionSelect={handleChange}
      value={selectedValue}
    >
      {options.map((option) => (
        <Option key={option.value} value={option.value} className="!w-32">
          {option.name}
        </Option>
      ))}
    </Dropdown>
  );
}

export default CustomDropdown;
