import * as React from 'react';
import { IFormPedidosCreateState } from './IFormPedidosCreateState';

import { PedidosDataProvider } from '../sharePointDataProvider/pedidos/PedidosDataProvider';
import { IPedidosProps } from '../../IPedidosProps';
import { Pedidos } from '../Models/pedidos/Pedidos';

interface IFormPedidosCreateProps extends IPedidosProps {
    _goBack?: () => void;
    _reload?: () => Promise<void>;
}
import {
    DefaultButton, Dialog, DialogFooter, IStackProps, IStackTokens, PrimaryButton, Stack, DialogType, TextField, MessageBarType, MessageBar, DatePicker, DayOfWeek
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

const columnPropsEvid: Partial<IStackProps> = {
    tokens: { childrenGap: 15 },
    styles: {
        root: {
            width: 250
        }
    }
};
const stackTokens: IStackTokens = { childrenGap: 20 };

const customStyles: React.CSSProperties = {
    display: 'flex',
    flexFlow: 'row',
    width: 'auto',
    height: 'auto',
    boxSizing: 'border-box',
    justifyContent: 'flex-start',
    alignItems: 'center',
};

const customStackProps: IStackProps = {
    tokens: { childrenGap: 15 },
    style: customStyles,
};

export default class FormPedidosCreate extends React.Component<IFormPedidosCreateProps, IFormPedidosCreateState> {
    private _PedidosDataProvider: PedidosDataProvider;

    private _createPedidos: Pedidos;
    private mensagemErrosCreatePedidos: string[] = [];
    private isSaving: boolean = false;
    private _siteUrl: string;



    constructor(props) {
        super(props);

        this._PedidosDataProvider = new PedidosDataProvider();
        // Configurar o contexto no PedidosDataProvider
        this._PedidosDataProvider.webPartContext = this.props.context as any;

        this._createPedidos = new Pedidos();


        this._createPedidos.Colaborador = null;
        this._createPedidos.Data_Inicio = null;
        this._createPedidos.Data_Fim = null;
        this._createPedidos.Estado = null;

        this.state = {

            siteUrl: this._siteUrl,
            isCreate: true,
            hideDialog: false,
            isBusy: false,
            _showErrors: false,
            mensagemErrosCreatePedidos: this.mensagemErrosCreatePedidos,

            createPedidos: this._createPedidos,

            _reload: props._reload || (() => Promise.resolve()),
            _goBack: props._goBack || (() => { }),
        };
    }
    public render(): React.ReactElement<{}> {
        const { createPedidos } = this.state;
        return (
            <div>
                {this.state._showErrors ? this._renderMessageBar() : null}
                {this.state._showErrors ? this._renderDialog() : null}
                <form>
                    <br />
                    <br />
                    <Stack horizontal horizontalAlign="space-between" {...customStackProps}>
                        <Stack {...columnPropsEvid}>
                            <PeoplePicker
                                context={{
                                    ...this.props.context,
                                    absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
                                    siteAbsoluteUrl: this.props.context.pageContext.site.absoluteUrl,
                                    webAbsoluteUrl: this.props.context.pageContext.web.absoluteUrl,
                                    msGraphClientFactory: this.props.context.msGraphClientFactory,
                                    spHttpClient: this.props.context.spHttpClient
                                } as any}
                                titleText="Seleciona um colaborador"
                                personSelectionLimit={1}
                                showtooltip={true}
                                required
                                onChange={this._getColaborador}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={1000}
                                ensureUser={true}
                                suggestionsLimit={15}
                            />
                        </Stack>
                        <Stack {...columnPropsEvid}>
                            {createPedidos.Colaborador && (
                                <TextField
                                    label="Colaborador selecionado"
                                    value={`${createPedidos.Colaborador.Title} (${createPedidos.Colaborador.EMail})`}
                                    disabled
                                    styles={{ root: { marginTop: '24px' } }}
                                />
                            )}
                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                    </Stack>
                    <br />
                    <br />
                    <Stack horizontal horizontalAlign="space-between" {...customStackProps}>
                        <Stack {...columnPropsEvid}>
                            <DatePicker
                                label="Seleciona data de inicio das Férias"
                                firstDayOfWeek={DayOfWeek.Monday}
                                placeholder="Escolhe uma data..."
                                ariaLabel="Seleciona uma data"
                                onSelectDate={this._onSelectedData_Inicio}
                                value={createPedidos.Data_Inicio}
                                isRequired
                            />
                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                    </Stack>
                    <br />
                    <br />
                    <Stack horizontal horizontalAlign="space-between" {...customStackProps}>
                        <Stack {...columnPropsEvid}>
                            <DatePicker
                                label="Seleciona data de fim das Férias"
                                firstDayOfWeek={DayOfWeek.Monday}
                                placeholder="Escolhe uma data..."
                                ariaLabel="Seleciona uma data"
                                onSelectDate={this._onSelectedData_Fim}
                                value={createPedidos.Data_Fim}
                                isRequired
                            />
                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                    </Stack>
                    <br />
                    <br />
                    <Stack horizontal horizontalAlign="space-between" {...customStackProps}>
                        <Stack {...columnPropsEvid}>
                            <TextField label="Estado" value="Pendente" disabled />
                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                        <Stack {...columnPropsEvid}>

                        </Stack>
                    </Stack>
                    <br />
                    <br />
                    <Stack horizontal tokens={stackTokens}>
                        <DefaultButton onClick={this.state._goBack} >Voltar</DefaultButton>
                        <PrimaryButton onClick={this._CreatePedidos} disabled={this.isSaving} >Submeter Pedido</PrimaryButton>
                    </Stack>
                </form>
            </div >
        );
    }
    private _renderDialog(): JSX.Element {
        return (
            <Dialog
                hidden={this.state.hideDialog}
                onDismiss={this._closeDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: 'Formulário com Erro',
                    subText: 'Consultar mensagem no topo.'
                }}
                modalProps={{
                    isBlocking: false,
                    styles: { main: { width: 1200 } }
                }}
            >
                <DialogFooter>
                    <PrimaryButton onClick={this._closeDialog} text="Voltar" />
                </DialogFooter>
            </Dialog>
        );
    }
    private _renderMessageBar(): JSX.Element {
        return (
            <MessageBar
                messageBarType={MessageBarType.blocked}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
                truncated={true}
                overflowButtonAriaLabel="See more"
            >
                <b> {this.state.mensagemErrosCreatePedidos}</b>
            </MessageBar>
        );
    }

    private _closeDialog = (): void => {
        this.setState({ hideDialog: true });
    }
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    private _onFormatDate = (date: Date): string => {
        return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    }
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    private _onSelectedData_Inicio = (date: Date | null | undefined): void => {

        const { createPedidos } = this.state;
        if (date) {
            const formatted = this._onFormatDate(date);
            createPedidos.Data_Inicio = date;
            createPedidos.formattedData_Inicio = formatted;

            this.setState({
                createPedidos: createPedidos
            });
        }
    }
    private _onSelectedData_Fim = (date: Date | null | undefined): void => {

        const { createPedidos } = this.state;
        if (date) {
            const formatted = this._onFormatDate(date);
            createPedidos.Data_Fim = date;
            createPedidos.formattedData_Fim = formatted;

            this.setState({
                createPedidos: createPedidos
            });
        }
    }
    private _getColaborador = (items: any[]): void => {
        const { createPedidos } = this.state;
        try {
            if (items.length > 0) {
                const selectedUser = items[0];
                console.log('Usuário selecionado para criar pedido:', selectedUser);

                createPedidos.Colaborador = {
                    Id: selectedUser.id || 0,
                    Title: selectedUser.text || selectedUser.displayName,
                    EMail: selectedUser.secondaryText || selectedUser.mail
                };

                this.setState({ createPedidos: createPedidos });
            } else {
                createPedidos.Colaborador = null;
                this.setState({ createPedidos: createPedidos });
            }
        } catch (error) {
            console.error('Erro ao processar seleção de colaborador:', error);
        }
    };
    private _ValidaFormularioSubmissao = async (): Promise<boolean> => {

        const { createPedidos } = this.state;

        const hoje = new Date();

        let mensagem: string[] = [];
        mensagem.push("Por favor preencher todos os campos assinalados com (*):");
        let validate = true;

        if (!createPedidos) {
            mensagem.push(" Relatorio Final SEG EE ");
            validate = false;
        }
        if (!createPedidos.Colaborador || !createPedidos.Colaborador.Title || !createPedidos.Colaborador.EMail) {
            mensagem.push("| Colaborador invalido / por preencher ");
            validate = false;
        }
        if (createPedidos.Data_Inicio === undefined || createPedidos.Data_Inicio === null) {
            mensagem.push("| Data de Inicio invalida / nula ");
            validate = false;
        }
        if (createPedidos.Data_Inicio && createPedidos.Data_Fim && createPedidos.Data_Inicio > createPedidos.Data_Fim) {
            mensagem.push("| Data de Inicio Posterior a data de Fim ");
            validate = false;
        }
        if (createPedidos.Data_Fim === undefined || createPedidos.Data_Fim === null ) {
            mensagem.push("| Data de Fim invalida / nula  ");
            validate = false;
        }
        if (createPedidos.Data_Fim && createPedidos.Data_Inicio && createPedidos.Data_Fim < createPedidos.Data_Inicio) {
            mensagem.push("| Data de Fim Anterior a data de Inicio ");
            validate = false;
        }
        if (createPedidos.Data_Inicio && createPedidos.Data_Inicio < hoje) {
            mensagem.push("| Data de Início não pode ser anterior a hoje ");
            validate = false;
        }
        if (createPedidos.Data_Fim && createPedidos.Data_Fim < hoje) {
            mensagem.push("| Data de Fim não pode ser anterior a hoje ");
            validate = false;
        }
        this.setState({
            mensagemErrosCreatePedidos: mensagem,
            _showErrors: !validate
        });

        return validate;
    }
    private _CreatePedidos = async (event: React.MouseEvent<HTMLButtonElement, MouseEvent>): Promise<void> => {
        const { createPedidos } = this.state;
        if (await this._ValidaFormularioSubmissao()) {
            try {
                this.isSaving = true;
                this.setState({ createPedidos: createPedidos });

                createPedidos.Estado = 'Pendente';
                await this._PedidosDataProvider.createItem(createPedidos);

                console.log('Pedido criado com sucesso');

                // Recarregar a lista de pedidos
                if (this.state._reload) {
                    await this.state._reload();
                }

                this.state._goBack();
            } catch (error) {
                console.error('Erro ao criar pedido:', error);
            } finally {
                this.isSaving = false;
            }
        }
        else {
            this._showDialog();
        }
    }
    private _showDialog = (): void => {
        this.setState({ hideDialog: false });
        this._renderDialog();
    }
}
