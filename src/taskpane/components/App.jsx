import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    margin: "0",
    padding: "0",
  },
});

function App() {
  const styles = useStyles();

  return <div className={styles.root}></div>;
}

export default App;
