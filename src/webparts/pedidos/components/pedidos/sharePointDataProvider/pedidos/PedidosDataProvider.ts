import { IWebPartContext } from '@microsoft/sp-webpart-base';

import { IPedidosDataProvider } from './IPedidosDataProvider';
import { IPedidos } from '../../Models/pedidos/IPedidos';
import { SPFI } from "@pnp/sp";
import { getSP } from '../../utils/pnpjsConfig';

const LIST_PEDIDOS = "Pedidos_de_Ferias";

export class PedidosDataProvider implements IPedidosDataProvider {
  private _sp: SPFI;
  private _webPartContext: IWebPartContext;
  constructor() {
    this._sp = getSP();
  }
  public set webPartContext(value: IWebPartContext) {
    this._webPartContext = value;
    this._sp = getSP(value as any);
  }
  public get webPartContext(): IWebPartContext {
    return this._webPartContext;
  }
  public async getItems(): Promise<IPedidos[]> {
    try {
      const result_pedidos = await this._sp.web.lists
        .getByTitle(LIST_PEDIDOS)
        .items
        .select(
          "Id", "Colaborador/Id", "Colaborador/Title", "Colaborador/EMail", "Data_Inicio", "Data_Fim", "Estado")
        .expand("Colaborador")();

      const pedidosFerias: IPedidos[] = [];

      result_pedidos.forEach(pedido => {
        if (typeof pedido !== 'undefined' && pedido) {
          pedidosFerias.push({
            Id: pedido.Id,
            Colaborador: pedido.Colaborador ? {
              Id: pedido.Colaborador.Id,
              Title: pedido.Colaborador.Title,
              EMail: pedido.Colaborador.EMail,
            } : undefined,
            Data_Inicio: pedido.Data_Inicio ? pedido.Data_Inicio : null,
            Data_Fim: pedido.Data_Fim ? pedido.Data_Fim : null,
            Estado: pedido.Estado ? pedido.Estado : null,
          });
        }
      });

      return pedidosFerias;
    } catch (error) {
      console.error("Erro ao buscar itens:", error);
      throw error;
    }
  }

  public async createItem(itemCreated: IPedidos): Promise<IPedidos> {
    try {
      const result = await this._sp.web.lists.getByTitle(LIST_PEDIDOS).items.add({
        ColaboradorId: itemCreated.Colaborador?.Id || null,
        Data_Inicio: itemCreated.Data_Inicio || null,
        Data_Fim: itemCreated.Data_Fim || null,
        Estado: itemCreated.Estado || null,
      });

      itemCreated.Id = result?.data?.Id || result?.Id || null;
      return itemCreated;
    } catch (error) {
      console.error("Erro ao criar item:", error);
      throw error;
    }
  }
   public updateEstado(PedidoUpdated: IPedidos): Promise<IPedidos> {
    const id = PedidoUpdated.Id;
    const updateEstadoPedido = {
      Estado: PedidoUpdated.Estado
    };
    return this._sp.web.lists
      .getByTitle(LIST_PEDIDOS)
      .items.getById(id)
      .update(updateEstadoPedido)
      .then(() => {
        return PedidoUpdated;
      });
  }
  public async deleteItem(id: number): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle(LIST_PEDIDOS).items.getById(id).delete();
      console.log(`Pedido com ID ${id} apagado com sucesso`);
    } catch (error) {
      console.error("Erro ao apagar item:", error);
      throw error;
    }
  }
}