import * as React from 'react';
import type { IPedidosProps } from './IPedidosProps';
import { DetailsListPedidos } from './pedidos/list/DetailsListPedidos';

export default class Pedidos extends React.Component<IPedidosProps> {
  public render(): React.ReactElement<IPedidosProps> {
    return (
      <DetailsListPedidos {...this.props}/>
    );
  }
}
