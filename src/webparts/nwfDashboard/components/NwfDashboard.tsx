import * as React from 'react';
import styles from './NwfDashboard.module.scss';
import { INwfDashboardProps } from './INwfDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import NWDashBoardAdmin from './DashboardNWF';
import { sp } from "@pnp/sp/presets/all";
export default class NwfDashboard extends React.Component<INwfDashboardProps, {}> {
  constructor(prop: INwfDashboardProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context
    });
  }
  public render(): React.ReactElement<INwfDashboardProps> {
    return (
      <NWDashBoardAdmin description={this.props.description} context={this.props.context} spcontext={sp.web}/>
    );
  }
}
