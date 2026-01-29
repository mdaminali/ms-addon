import * as React from "react";
import { Image, tokens, makeStyles } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingBottom: "10px",
    paddingTop: "10px",
    backgroundColor: tokens.colorNeutralBackground3,
    position: "fixed",
    width: "100%",
    top: 0,
    left: 0,
    zIndex: 1000,
    boxShadow: tokens.shadow2,
  },
  message: {
    fontSize: "24px",
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title, logo, message } = props;
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image width="180" src={logo} alt={title} />
      {/* <h1 className={styles.message}>{message}</h1> */}
    </section>
  );
};

export default Header;
