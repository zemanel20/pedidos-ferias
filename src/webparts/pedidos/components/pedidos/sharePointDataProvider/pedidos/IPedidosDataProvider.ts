import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IPedidos } from '../../Models/pedidos/IPedidos';
export interface IPedidosDataProvider {
  webPartContext: IWebPartContext;  
  getItems(): Promise<IPedidos[]>;
  createItem(itemCreated: IPedidos): Promise<IPedidos>;
  deleteItem(id: number): Promise<void>;
}
