import { IPedidos } from '../Models/pedidos/IPedidos';
export interface ICommandBarPedidosState {
  items: IPedidos[];
  isVisible: boolean;
  _goBack: VoidFunction;
  _reload: VoidFunction;
}
