import * as React from 'react';
import { CommandBarPedidos } from '../utils/CommandBarPedidos';
import { IDetailsListPedidosState } from './IDetailsListPedidosState';
import { IFilterPedidosColab } from '../utils/IFilterPedidosColab';
import { FilterPedidosColab } from '../utils/FilterPedidosColab';

import { PedidosDataProvider } from '../sharePointDataProvider/pedidos/PedidosDataProvider';

import { IPedidos } from '../Models/pedidos/IPedidos';
import { IPedidosProps } from '../../IPedidosProps';
import {
  DefaultButton, DetailsList, Fabric, IColumn, 
  IStackProps,  MarqueeSelection, ProgressIndicator, Stack,
  Selection, SelectionMode, DetailsListLayoutMode,
  Persona, PersonaSize, mergeStyles
} from '@fluentui/react';

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

//Aproximar a imagem importada via Graph
const personaHoverWrapperClass = mergeStyles({
  selectors: {
    '.ms-Persona-imageArea': {
      transition: 'transform 0.2s ease-in-out',
    },
    '&:hover .ms-Persona-imageArea': {
      transform: 'scale(1.8)',
      zIndex: 10,
    },
  },
});

export class DetailsListPedidos extends React.Component<IPedidosProps, IDetailsListPedidosState> {
  private _FilterPedidosColab: IFilterPedidosColab;
  private _selection: Selection;
  private _allItems: IPedidos[];
  private percentComplete: number;
  private titleProgressBar: string = "A obter o Pedido de Férias.";
  private descriptionProgressBar: string = "a processar...";
  private _PedidosDataProvider: PedidosDataProvider;

  constructor(props) {

    super(props);

    this._PedidosDataProvider = new PedidosDataProvider();
    this._PedidosDataProvider.webPartContext = this.props.context as any;
    this._LoadPedidos();
    this._FilterPedidosColab = new FilterPedidosColab();
    //Colunas da listagem
    const columns: IColumn[] = [
      {
        key: 'Id',
        name: 'Id',
        isResizable: true,
        fieldName: 'Id',
        minWidth: 50,
        maxWidth: 50,
        data: 'string',
        onColumnClick: this._onColumnClick,
        isPadded: true
      },
      {
        key: 'Colaborador',
        name: 'Colaborador',
        fieldName: 'Colaborador',
        minWidth: 80,
        maxWidth: 200,
        isResizable: true,
        isPadded: true,
        data: 'object',
        onColumnClick: this._onColumnClick,
        onRender: (item: IPedidos) => {
          if (!item.Colaborador) return '-';

          const photoUrl = this._getUserPhotoUrl(item.Colaborador.EMail);

          return (
            <div className={personaHoverWrapperClass}>
              <Persona
                text={item.Colaborador.Title}
                secondaryText={item.Colaborador.EMail}
                size={PersonaSize.size32}
                hidePersonaDetails={false}
                imageUrl={photoUrl}
                imageInitials={item.Colaborador.Title?.charAt(0)?.toUpperCase() || '?'}
              />
            </div>
          );
        }
      },
      {
        key: 'Data_Inicio',
        name: 'Data de Início',
        fieldName: 'Data_Inicio',
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        isPadded: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        isSorted: true,
        isSortedDescending: false,
        onRender: (item: IPedidos) =>
          item.Data_Inicio ? this._onFormatDate(new Date(item.Data_Inicio)) : '-'
      },
      {
        key: 'Data_Fim',
        name: 'Data de Fim',
        fieldName: 'Data_Fim',
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        isPadded: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item: IPedidos) =>
          item.Data_Fim ? this._onFormatDate(new Date(item.Data_Fim)) : '-'
      },
      {
        key: 'Estado',
        name: 'Estado',
        isResizable: true,
        fieldName: 'Estado',
        minWidth: 80,
        maxWidth: 100,
        data: 'string',
        onColumnClick: this._onColumnClick,
        isPadded: true,
        onRender: (item: IPedidos) => {
          const estado = item.Estado || '';
          let backgroundColor = '#f3f2f1'; // Cor padrão
          let color = '#323130';

          switch (estado) {
            case 'Aprovado':
              backgroundColor = '#107C10'; // Verde
              color = 'white';
              break;
            case 'Rejeitado':
              backgroundColor = '#D13438'; // Vermelho
              color = 'white';
              break;
            case 'Pendente':
              backgroundColor = '#FFB900'; // Amarelo
              color = 'black';
              break;
          }

          return (
            <div
              style={{
                backgroundColor,
                color,
                padding: '4px 8px',
                borderRadius: '4px',
                textAlign: 'center',
                fontWeight: '600',
                fontSize: '12px'
              }}
            >
              {estado}
            </div>
          );
        }
      }
    ];
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
        });
      }
    });
    this.state = {
      items: [],
      isDataLoaded: true,
      columns: columns,
      _PedidosDataProvider: this._PedidosDataProvider,
      selectionDetails: this._getSelectionDetails(),
      _goBack: this._hidePanel,
      _FilterPedidosColab: this._FilterPedidosColab,
      selectedPedido: null,
      userPhotos: {},
    };
  }
  public render() {
    const { columns, items} = this.state;
    return (
      <Fabric>
        <div style={{ marginBottom: '20px' }}>
          <Stack 
            horizontal 
            verticalAlign="center" 
            tokens={{ childrenGap: 30 }}
            styles={{
              root: {
                width: '100%',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'flex-start'
              }
            }}
          >
            <Stack.Item>
              <CommandBarPedidos
                {...this.props}
                _reload={this._LoadPedidos}
              />
            </Stack.Item>
            <Stack.Item styles={{ root: { minWidth: '300px' } }}>
              <PeoplePicker
                titleText="Filtrar por colaborador"
                personSelectionLimit={1}
                groupName={""}
                showtooltip={true}
                required={false}
                onChange={this._onPeoplePickerChange}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={500}
                context={{
                  ...this.props.context,
                  absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
                  siteAbsoluteUrl: this.props.context.pageContext.site.absoluteUrl,
                  webAbsoluteUrl: this.props.context.pageContext.web.absoluteUrl,
                  msGraphClientFactory: this.props.context.msGraphClientFactory,
                  spHttpClient: this.props.context.spHttpClient
                } as any}
                ensureUser={true}
                suggestionsLimit={15}
              />
            </Stack.Item>
          </Stack>
        </div>
        {this._selection.getSelectedCount() === 1 && (
          <div style={{ marginBottom: '10px' }}>
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              {(this._selection.getSelection()[0] as IPedidos)?.Estado === 'Pendente' && (
                <>
                  <DefaultButton
                    text="Aprovar"
                    iconProps={{ iconName: 'CheckMark' }}
                    onClick={() => this._onApprove(this._selection.getSelection()[0] as IPedidos)}
                    styles={{
                      root: {
                        backgroundColor: '#107C10',
                        color: 'white',
                        border: 'none'
                      },
                      rootHovered: {
                        backgroundColor: '#0E6B0E',
                        color: 'white'
                      }
                    }}
                  />
                  <DefaultButton
                    text="Rejeitar"
                    iconProps={{ iconName: 'Cancel' }}
                    onClick={() => this._onReject(this._selection.getSelection()[0] as IPedidos)}
                    styles={{
                      root: {
                        backgroundColor: '#D13438',
                        color: 'white',
                        border: 'none'
                      },
                      rootHovered: {
                        backgroundColor: '#B92B2F',
                        color: 'white'
                      }
                    }}
                  />
                </>
              )}
              <DefaultButton
                text="Apagar"
                iconProps={{ iconName: 'Delete' }}
                onClick={() => this._onDelete(this._selection.getSelection()[0] as IPedidos)}
                styles={{
                  root: {
                    backgroundColor: '#A80000',
                    color: 'white',
                    border: 'none'
                  },
                  rootHovered: {
                    backgroundColor: '#960000',
                    color: 'white'
                  }
                }}
              />
            </Stack>
          </div>
        )}
        <MarqueeSelection selection={this._selection}>
          {this.state.isDataLoaded == false ? <ProgressIndicator label={this.titleProgressBar} description={this.descriptionProgressBar} percentComplete={this.percentComplete} /> : null}
          <DetailsList
            items={items}
            columns={columns}
            selectionMode={SelectionMode.single}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={(item) => { this._onItemInvoked(item, this); }}
            enterModalSelectionOnTouch={true}
          />
        </MarqueeSelection>
      </Fabric>
    );
  }
  //Carregar lista
  private _LoadPedidos = async (): Promise<void> => {
    try {
      const itemsList: IPedidos[] = await this._PedidosDataProvider.getItems();

      const sortedItems = _copyAndSort(itemsList, 'Data_Inicio', false);

      this.setState({ items: sortedItems});
      this._allItems = sortedItems;
    } catch (error) {
      console.error("Erro ao carregar pedidos:", error);
    }
  }
  public componentDidUpdate(previousProps: any, previousState: IDetailsListPedidosState) {
  }
  private _onFormatDate = (date: Date): string => {
    return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
  }
  public async componentDidMount(): Promise<void> {
    this._LoadPedidos();
    this.descriptionProgressBar = "concluido.";
    this.percentComplete = 1;
    this.setState({ isDataLoaded: true });
  }
  private _onItemInvoked(item: any, value: any): void {
    const itemPedidos = item as IPedidos;
    value.setState({ selectedPedido: itemPedidos });
  }
  //Esconder form
  private _hidePanel = () => {
    this.setState({ isDataLoaded: false });
    this._LoadPedidos();
    this.descriptionProgressBar = "concluido.";
    this.percentComplete = 1;
    this.setState({ isDataLoaded: true });
  }
  private _getSelectionDetails(): string {
    this.setState({ selectedPedido: this._selection.getSelection()[0] as IPedidos });
    return;
  }

  //OnChange People Picker para filtragem
  private _onPeoplePickerChange = (items: any[]): void => {
    try {
      if (items.length > 0) {
        const selectedUser = items[0];
        const userEmail = selectedUser?.secondaryText?.toLowerCase() || selectedUser?.mail?.toLowerCase();
        this._FilterPedidosColab.Colaborador = userEmail;
        this._applyFilter(this._FilterPedidosColab);
      } else {
        this._FilterPedidosColab.Colaborador = '';
        this._applyFilter(this._FilterPedidosColab);
      }
    } catch (error) {
      console.error('Erro ao processar seleção de usuário:', error);
    }
  };

  //Aplica filtro
  private _applyFilter = (filterPedidos: FilterPedidosColab): void => {
    let filterItems = this._allItems;

    filterItems = filterPedidos.Colaborador
      ? filterItems.filter(i =>
        i.Colaborador?.EMail?.toLowerCase() === filterPedidos.Colaborador
      )
      : filterItems;

    const sortedFilteredItems = _copyAndSort(filterItems, 'Data_Inicio', false);

    this.setState({
      items: sortedFilteredItems,
      _FilterPedidosColab: filterPedidos
    });
  };

  // Clicar na coluna para ordenar
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
    });
    this.setState({
      items: newItems
    });
  }
  public componentWillUnmount(): void {

    Object.values(this.state.userPhotos).forEach(url => {
      if (url && url !== 'no-photo' && url.startsWith('blob:')) {
        URL.revokeObjectURL(url);
      }
    });
  }

  // Função para aprovar pedido
  private _onApprove = async (item: any): Promise<void> => {
    try {
      if (item.Estado !== 'Pendente') {
        alert('Apenas pedidos pendentes podem ser Aprovados.');
        return;
      }
      const confirmDelete = confirm(`Tem a certeza que pretende Aprovar o pedido de ${item.Colaborador?.Title}?`);

      if (!confirmDelete) {
        return;
      }
      item.Estado = "Aprovado";
      this._PedidosDataProvider.updateEstado(item).then((pedidos: IPedidos) => {
        // Recarregar a lista
        this._LoadPedidos();

        // Limpar seleção
        this._selection.setAllSelected(false);
      });


    } catch (error) {
      console.error('Erro ao aprovar pedido:', error);
      alert('Erro ao aprovar o pedido.');
    }
  }

  // Função para rejeitar pedido
  private _onReject = async (item: any): Promise<void> => {
    try {
      if (item.Estado !== 'Pendente') {
        alert('Apenas pedidos pendentes podem ser Rejeitados.');
        return;
      }
      const confirmDelete = confirm(`Tem a certeza que pretende Rejeitar o pedido de ${item.Colaborador?.Title}?`);

      if (!confirmDelete) {
        return;
      }
      item.Estado = "Rejeitado";
      this._PedidosDataProvider.updateEstado(item).then((pedidos: IPedidos) => {
        // Recarregar a lista
        this._LoadPedidos();

        // Limpar seleção
        this._selection.setAllSelected(false);
      });


    } catch (error) {
      console.error('Erro ao rejeitar pedido:', error);
      alert('Erro ao rejeitar o pedido.');
    }
  }

  // Função para apagar pedido
  private _onDelete = async (item: any): Promise<void> => {
    try {
      const confirmDelete = confirm(`Tem a certeza que quer apagar o pedido de ${item.Colaborador?.Title}?`);

      if (!confirmDelete) {
        return;
      }

      await this._PedidosDataProvider.deleteItem(item.Id);

      // Recarregar a lista
      await this._LoadPedidos();

      this._selection.setAllSelected(false);
    } catch (error) {
      console.error('Erro ao apagar pedido:', error);
      alert('Erro ao apagar o pedido.');
    }
  };

  // Metodo de obtencao do URL da foto do usuário via graph
  private _getUserPhotoUrl = (userEmail: string): string | undefined => {
    if (!userEmail) return undefined;

    const email = userEmail.toLowerCase();
    const photoUrl = this.state.userPhotos[email];

    if (photoUrl === undefined) {
      // Tentar carregar a foto (apenas uma vez)
      this._loadUserPhoto(email);
      return undefined;
    }

    // Se está carregando ou não existe foto, retornar undefined
    if (photoUrl === 'loading' || photoUrl === 'no-photo') {
      return undefined;
    }
    // Se temos uma URL válida, retornar
    return photoUrl;
  };

  // Função para carregar foto do usuário via Microsoft Graph
  private _loadUserPhoto = async (userEmail: string): Promise<void> => {
    try {
      // Verificar se já está a ser carregada ou se ja foi carregada
      if (this.state.userPhotos[userEmail] !== undefined) {
        return;
      }

      // Marcar como "loading" para nao haver múltiplos pedidos
      this.setState(prevState => ({
        userPhotos: {
          ...prevState.userPhotos,
          [userEmail]: 'loading'
        }
      }));

      // Obter o MSGraphClientFactory do contexto getClient('3') refere-se a versão da Graph API v1.0
      const graphClient = await this.props.context.msGraphClientFactory.getClient('3');

      try {
        // Tenta obter a foto do usuário
        const photoResponse = await graphClient
          .api(`/users/${userEmail}/photo/$value`)
          .get();

        // Verificar se é válida
        if (!photoResponse) {
          throw new Error('Resposta vazia do Graph API');
        }

        // Converter blob para URL
        let photoBlob;
        if (photoResponse instanceof ArrayBuffer) {
          photoBlob = new Blob([photoResponse], { type: 'image/jpeg' });
        } else if (photoResponse instanceof Blob) {
          photoBlob = photoResponse;
        } else {
          photoBlob = new Blob([photoResponse], { type: 'image/jpeg' });
        }

        const photoUrl = URL.createObjectURL(photoBlob);

        // Atualizar o estado com a nova foto
        this.setState(prevState => ({
          userPhotos: {
            ...prevState.userPhotos,
            [userEmail]: photoUrl
          },
          // Forçar re-render da lista adicionando timestamp
          items: [...prevState.items]
        }), () => {
          // Forçar re-render explícito
          this.forceUpdate();
        });

      } catch (photoError) {

        // Foto não encontrada - isso é normal para muitos usuários
        // Marcar como tentativa realizada para evitar tentativas repetidas
        this.setState(prevState => ({
          userPhotos: {
            ...prevState.userPhotos,
            [userEmail]: 'no-photo'
          },
          // Forçar re-render da lista
          items: [...prevState.items]
        }));
      }

    } catch (error) {
      // Erro geral (problemas de rede, permissões, etc.)
      console.error('Erro ao carregar foto do usuário:', error);
      // Marcar como sem foto para evitar tentativas repetidas
      this.setState(prevState => ({
        userPhotos: {
          ...prevState.userPhotos,
          [userEmail]: 'no-photo'
        },
        // Forçar re-render da lista
        items: [...prevState.items]
      }));
    }
  };


}
function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.sort((a: T, b: T) => {
    if (columnKey === 'Data_Inicio' || columnKey === 'Data_Fim') {
      const dateA = a[key] ? new Date(a[key] as any) : new Date(0);
      const dateB = b[key] ? new Date(b[key] as any) : new Date(0);
      return isSortedDescending ? dateB.getTime() - dateA.getTime() : dateA.getTime() - dateB.getTime();
    }
    return ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1);
  });
}
