import useStore from "../utils/store";
import Header from "./Header";
import Formula from "./Fomula/Formula";
import Style from "./Style/Style";
import Macro from "./Macro/Macro";
import Validate from "./Validate/Validate";
import Share from "./Share/Share";
import { useStyles } from "../utils/style";

function App() {
  const styles = useStyles();
  const { category } = useStore();

  const categories = {
    Formula: <Formula />,
    Style: <Style />,
    Macro: <Macro />,
    Validate: <Validate />,
    Share: <Share />,
  };
  const CurrentCategory = categories[category] || null;

  return (
    <div className={styles.root}>
      <Header />
      {CurrentCategory || <div />}
    </div>
  );
}

export default App;
