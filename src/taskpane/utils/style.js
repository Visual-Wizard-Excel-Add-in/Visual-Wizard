import {
  tokens,
  makeStyles,
  createLightTheme,
  createDarkTheme,
} from "@fluentui/react-components";
import {
  iconFilledClassName,
  iconRegularClassName,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground6,
  },
  list: {
    display: "flex",
    alignItems: "flex-start",
    flexDirection: "column",
    justifyContent: "flex-start",
    backgroundColor: tokens.colorNeutralBackground6,
  },
  accordion: {
    backgroundColor: tokens.colorNeutralBackground3,
    border: "1px solid #AAAAAA",
  },
  openedAccordion: {
    backgroundColor: tokens.colorNeutralBackground3Pressed,
    border: "1px solid #AAAAAA",
  },
  card: {
    borderRadius: "0%",
    boxShadow: "0",
  },
  margin0: {
    margin: "0%",
  },
  blurText: {
    color: tokens.colorNeutralForeground4,
  },
  border: {
    margin: "0.3rem 0",
    borderBlockWidth: "1px",
    borderBlockColor: tokens.colorNeutralBackground3Pressed,
  },
  buttons: {
    color: tokens.colorNeutralForeground4,
    ":hover": {
      borderRadius: tokens.borderRadiusMedium,
      backgroundColor: tokens.colorNeutralBackground5Hover,
    },
    ":active": {
      [`& .${iconFilledClassName}`]: {
        display: "block",
      },
      [`& .${iconRegularClassName}`]: {
        display: "none",
      },
    },
  },
  macroKey: {
    margin: "0px",
    width: "3rem",
  },
  fontBolder: {
    fontWeight: "bolder",
  },
  messageBarGroup: {
    padding: tokens.spacingHorizontalMNudge,
    display: "flex",
    flexDirection: "column",
    marginTop: "10px",
    gap: "10px",
    overflow: "auto",
  },
});

const excelTheme = {
  10: "#020402",
  20: "#101C14",
  30: "#162E1F",
  40: "#193C27",
  50: "#1C4A2F",
  60: "#1E5937",
  70: "#20683F",
  80: "#28764A",
  90: "#40835A",
  100: "#56906B",
  110: "#6A9E7C",
  120: "#7FAB8D",
  130: "#93B89F",
  140: "#A7C6B1",
  150: "#BCD3C3",
  160: "#D1E1D5",
};

const lightTheme = {
  ...createLightTheme(excelTheme),
};

const darkTheme = {
  ...createDarkTheme(excelTheme),
};
const { 110: color110, 120: color120 } = excelTheme;
darkTheme.colorBrandForeground1 = color110;
darkTheme.colorBrandForeground2 = color120;

export { lightTheme, darkTheme, useStyles };
