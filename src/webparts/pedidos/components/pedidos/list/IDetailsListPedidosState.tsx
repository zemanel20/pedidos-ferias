import { IFilterPedidosColab } from '../utils/IFilterPedidosColab';
import { IPedidos } from '../Models/pedidos/IPedidos';
import { PedidosDataProvider } from '../sharePointDataProvider/pedidos/PedidosDataProvider';
import { IColumn } from '@fluentui/react';
export interface IDetailsListPedidosState {
    
    items: IPedidos[];
    isDataLoaded?: boolean;
    columns: IColumn[];
    _PedidosDataProvider: PedidosDataProvider;
    selectionDetails: string;
    _goBack: VoidFunction;
    _FilterPedidosColab: IFilterPedidosColab;
    announcedMessage?: string;
    selectedPedido: IPedidos;
    userPhotos: { [email: string]: string }; 
}
