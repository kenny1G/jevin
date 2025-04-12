import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import FinancialModelGenerator from "./FinancialModelGenerator";
import FilingUploader from "./FilingUploader";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Generate 3-Statement Financial Models",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Upload JSON Financial Data",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Visualize Financial Statements",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Upload your financial data JSON file to create a 3-statement model!" items={listItems} />
      <FinancialModelGenerator />
      <FilingUploader />
    </div>
  );
};

export default App;
