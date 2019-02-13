import * as React from "react";
import styles from "./header.module.scss";

export interface IHeaderProps {
    title: string;
}

export default class Header extends React.Component<IHeaderProps, {}> {
    public render(): React.ReactElement<IHeaderProps> {
      return (
        <div className={styles.header}>
          <div className={styles.row}>
            <span className={styles.headerText}>{this.props.title}</span>
          </div>
        </div>
      );
    }
  }