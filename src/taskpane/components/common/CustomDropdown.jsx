import { Dropdown, Option } from "@fluentui/react-components";
import { useStyles } from "../../utils/style";

function CustomDropdown({
  handleValue,
  options,
  placeholder,
  selectedValue = null,
}) {
  const styles = useStyles();

  const handleChange = (event, value) => {
    event.stopPropagation();
    handleValue(value.optionText);
  };

  return (
    <Dropdown
      positioning="below"
      placeholder={placeholder}
      button={
        <span className={styles.optionBox}>{selectedValue || placeholder}</span>
      }
      onOptionSelect={handleChange}
      value={selectedValue}
    >
      {options.map((option) => (
        <Option
          key={`without group-${option.value}`}
          value={option.value}
          className="!w-fit"
        >
          {option.name}
        </Option>
      ))}
    </Dropdown>
  );
}

export default CustomDropdown;
