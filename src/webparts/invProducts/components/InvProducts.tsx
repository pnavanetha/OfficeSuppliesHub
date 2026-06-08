import * as React from 'react';
import type { IInvProductsProps } from './IInvProductsProps';

import { HashRouter } from 'react-router-dom';

import RoutesItems from './Navigation/Routes';


export default class InvProducts extends React.Component<IInvProductsProps> {

  public render(): React.ReactElement<IInvProductsProps> {
    return (
      <HashRouter>
        <RoutesItems context={this.props.context} />
      </HashRouter>
    );
  }
}