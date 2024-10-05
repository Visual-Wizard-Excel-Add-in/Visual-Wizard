import { Dropdown, Option, OptionGroup } from "@fluentui/react-components";
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

  const groupedOptions = options.reduce((acc, option) => {
    const label = option.label || "";

    if (!acc[label]) {
      acc[label] = [];
    }

    acc[label].push(option);

    return acc;
  }, {});

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
      {Object.entries(groupedOptions).map(([label, categorizedOptions]) =>
        label ? (
          <OptionGroup key={`group-${label}`} label={label}>
            {categorizedOptions.map((option) => (
              <Option key={`option-${option.value}`} value={option.value}>
                {option.name}
              </Option>
            ))}
          </OptionGroup>
        ) : (
          categorizedOptions.map((option) => (
            <Option
              key={`without group-${option.value}`}
              value={option.value}
              className="!w-fit"
            >
              {option.name}
            </Option>
          ))
        ),
      )}
    </Dropdown>
  );
}

export default CustomDropdown;
