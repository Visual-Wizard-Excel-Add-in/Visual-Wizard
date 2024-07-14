import {
  tokens,
  makeStyles,
  createLightTheme,
  createDarkTheme,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground6,
  },
  list: {
    alignItems: "flex-start",
    display: "flex",
    flexDirection: "column",
    justifyContent: "flex-start",
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
