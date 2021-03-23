import * as React from "react";
import styles from "./CamlQueryPnpjs.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";
import { ICamlQueryPnpjsProps } from "./CamlQueryPnpjsProps";

export default class CamlQueryPnpjs extends React.Component<
  ICamlQueryPnpjsProps,
  {}
> {
  public render(): React.ReactElement<ICamlQueryPnpjsProps> {
    return (
      <div className={styles.camlQueryPnpjs}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
              <ul className={styles.list}>
                {this.props.items.map((li) => (
                  <li key={li.Id} className={styles.item}>
                    Id: {li.Id}, Title: {li.Title}
                  </li>
                ))}
              </ul>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
