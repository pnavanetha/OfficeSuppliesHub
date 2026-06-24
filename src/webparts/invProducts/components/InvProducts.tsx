import * as React from 'react';
import type { IInvProductsProps } from './IInvProductsProps';
import { HashRouter } from 'react-router-dom';
import RoutesItems from './Navigation/Routes';
// import { SPHttpClient } from '@microsoft/sp-http';
import NavigationBar from "./Navigation/NavigationBar";
import { ToastContainer } from 'react-toastify';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import "./CSS/App.css";

interface IState {
  role: 'Admin' | 'Staff' | null;
}

export default class InvProducts extends React.Component<IInvProductsProps, IState> {

  constructor(props: IInvProductsProps) {
    super(props);
    this.state = { role: null };
  }

  public async componentDidMount(): Promise<void> {
    const isAdmin = await this.isUserInGroup('Inventory Admin');

    if (isAdmin) {
      this.setState({ role: 'Admin' });
    } else {
      this.setState({ role: 'Staff' });
    }
  }

  private async isUserInGroup(groupName: string): Promise<boolean> {
    
    console.log("Group Name:", groupName);
    console.log("User email:", this.props.context.pageContext.user.email);

    // const response = await this.props.context.spHttpClient.get(
    //   `${this.props.context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups`,
    //   SPHttpClient.configurations.v1
    // );
    const sp = spfi().using(SPFx(this.props.context));



    // const data = await response.json();
    // return data.value.some((g: any) => g.Title === groupName);

    const groups = await sp.web.currentUser.groups();
    return groups.some((g: any) => g.Title === groupName);
    
  }

  public render(): React.ReactElement<IInvProductsProps> {
    if (!this.state.role) {
       return <div>Loading...</div>;
    }

    return (
      <HashRouter>
        <ToastContainer />
        <div className="app-container">
          <NavigationBar role={this.state.role} />

          <div className="main-content">
            <RoutesItems
              context={this.props.context}
              role={this.state.role}
            />
          </div>
        </div>
      </HashRouter>
    );
  }
}