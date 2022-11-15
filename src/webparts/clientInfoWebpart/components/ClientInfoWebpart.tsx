import * as React from "react";
import styles from "./ClientInfoWebpart.module.scss";
import { IClientInfoWebpartProps } from "./IClientInfoWebpartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import App from "../AppComponents/App";

export default class ClientInfoWebpart extends React.Component<
  IClientInfoWebpartProps,
  {}
> {
  public render(): React.ReactElement<IClientInfoWebpartProps> {
    return <App spContext={this.props.spContext}></App>;
  }
}
