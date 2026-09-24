import { IPedidos } from "../Models/pedidos/IPedidos";
export interface IFormPedidosCreateState {
  isBusy: boolean;
  isCreate?: boolean;
  hideDialog: boolean;
  mensagemErrosCreatePedidos: string[];
  createPedidos?: IPedidos;
  _showErrors: boolean;
  siteUrl: string;

  _goBack: VoidFunction;
  _reload: VoidFunction;
  _CreatePedidos?(event: React.MouseEvent<HTMLButtonElement, MouseEvent>): Promise<void>;
}
