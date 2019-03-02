import * as React from "react";
import styles from "./header.module.scss";

export interface IHeaderProps {
  title: string;
}

export default class Header extends React.Component<IHeaderProps, {}> {
  public render(): React.ReactElement<IHeaderProps> {
    const { title } = this.props;
    const displayHeader = (typeof title === "string" && title.length > 0);
    return (
      <div className={displayHeader ? styles.header : ""}>
        {displayHeader && <div className={styles.row}>
          <span className={styles.headerText}>{this.props.title}</span>
        </div>}
      </div>
    );
  }
}