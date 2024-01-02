import * as React from "react";
import styles from "./CreateReleaseAutomation.module.scss";
import { ICreateReleaseAutomationProps } from "./ICreateReleaseAutomationProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import MainComponent from "./MainComponent";
export default class CreateReleaseAutomation extends React.Component<
  ICreateReleaseAutomationProps,
  {}
> {
  constructor(prop: ICreateReleaseAutomationProps) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context,
    });
  }
  public render(): React.ReactElement<ICreateReleaseAutomationProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return <MainComponent context={this.props.context} />;
  }
}
