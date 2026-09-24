import * as React from 'react';
import FormPedidosCreate from '../create/FormPedidosCreate';
import { ICommandBarPedidosState } from './ICommandBarPedidosState';
import { IPedidosProps } from '../../IPedidosProps';
import { CommandBar, DefaultButton, IPanelProps, IRenderFunction, IStackTokens, Panel, PanelType, Stack } from '@fluentui/react';

interface ICommandBarProps extends IPedidosProps {
  _reload?: () => Promise<void>;
}

const stackTokens: IStackTokens = { childrenGap: 20 };
export class CommandBarPedidos extends React.Component<ICommandBarProps, ICommandBarPedidosState> {
  constructor(props) {
    super(props);
    this.state = {
      isVisible: false,
      items: [],
      _goBack: this._hidePanel,
      _reload: this._hidePanel,
    };
  }

  public render(): JSX.Element {
    return (
      <div>
          <CommandBar
          styles={{
            root: {
              alignItems: 'stretch',
            },
          }}
          items={this.getItems()}
        />        
        <Panel isOpen={this.state.isVisible} onDismiss={this._hidePanel} type={PanelType.smallFluid} headerText={"Novo Pedido de Férias"}
          onRenderNavigationContent={this._onRenderNavigationContent}
        >
          <FormPedidosCreate 
            {...this.props} 
            _goBack={this._hidePanel}
            _reload={this.props._reload || (() => Promise.resolve())}
          />
        </Panel>
      </div>
    );
  }
  private _onRenderNavigationContent: IRenderFunction<IPanelProps> = (
    props?: IPanelProps,
    defaultRender?: IRenderFunction<IPanelProps>
  ): JSX.Element => {
    return (
      <React.Fragment>
        <Stack horizontal styles={{
          root: {
            margin: '10px',
            marginLeft: '30px',
            height: 'auto',
            width: '100%'
          }
        }}
          tokens={stackTokens}>
          <DefaultButton onClick={this.state._goBack} >Cancelar</DefaultButton>
        </Stack>
        {defaultRender!(props)}
      </React.Fragment>
    );
  }

  private getItems = () => {
    return [
      {
        key: 'newItem',
        name: 'Novo',
        cacheKey: 'myCacheKey',
        iconProps: {
          iconName: 'Add'
        },
        ariaLabel: 'Novo',
        subMenuProps: {
          items: [
            {
              key: 'pedidoItem',
              name: 'Pedido de Férias',
              iconProps: {
                iconName: 'SwayLogo16'
              },
              onClick: () => {
                this.setState({ isVisible: true });
              }
            }
          ]
        }
      }
    ];
  }
  public async componentDidMount(): Promise<void> {
  }
  private _hidePanel = () => {
    this.setState({ isVisible: false });
  }
}
