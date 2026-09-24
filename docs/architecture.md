# Arquitetura

Notas sobre como o projeto está montado por dentro — pensado para quem vai continuar o desenvolvimento e precisa de saber onde mexer.

## Visão geral

É uma única web part SPFx. Não há backend próprio, não há servidor, não há base de dados: a lista SharePoint `Pedidos_de_Ferias` **é** a camada de persistência, e o browser fala diretamente com a REST API do SharePoint usando o token do utilizador autenticado na página.

A hierarquia de componentes é curta:

```
PedidosWebPart            (SPFx, não-React)
  └─ Pedidos              (passthrough)
       └─ DetailsListPedidos          ← estado da listagem, ações, Graph
            ├─ CommandBarPedidos      ← botão "Novo" + Panel
            │    └─ FormPedidosCreate ← formulário de submissão
            └─ PeoplePicker           ← filtro por colaborador

PedidosDataProvider       ← instanciado por DetailsListPedidos e por FormPedidosCreate
       └─ PnPjs (SPFI)    ← singleton em pnpjsConfig.ts
```

## Arranque

`PedidosWebPart.ts` é o ponto de entrada declarado em `config/config.json`.

```ts
protected async onInit(): Promise<void> {
  await super.onInit();
  getSP(this.context);
}
```

Este `getSP(context)` é o passo crítico. O `pnpjsConfig.ts` guarda a instância `SPFI` numa variável de módulo:

```ts
let _sp: SPFI;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context && !_sp) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};
```

Repara na guarda `!_sp`: a instância só é criada **uma vez**. Chamadas posteriores com contexto são silenciosamente ignoradas e devolvem a instância existente. É por isso que o `onInit` tem de correr antes de qualquer componente — quando o `PedidosDataProvider` chama `getSP()` no construtor, sem argumentos, conta com o singleton já estar populado.

Há também um `initializeSP(context)` exportado que força a recriação. Não é chamado em lado nenhum; existe como escape se alguma vez for preciso trocar de contexto.

Depois do `onInit`, o `render()` cria o elemento React e faz `ReactDom.render` no `this.domElement`.

## Props

`IPedidosProps` tem seis campos, mas na prática só um interessa:

```ts
export interface IPedidosProps {
  description: string;         // não lido
  isDarkTheme: boolean;        // não lido, sempre false
  environmentMessage: string;  // não lido, sempre ''
  hasTeamsContext: boolean;    // não lido
  userDisplayName: string;     // não lido
  context: WebPartContext;     // ← este é o que importa
}
```

O `context` é usado em três sítios distintos:

1. **PnPjs** — `PedidosDataProvider.webPartContext = this.props.context`
2. **PeoplePicker** — os controlos PnP precisam de um objeto de contexto com `absoluteUrl`, `spHttpClient` e `msGraphClientFactory`. Como a forma esperada pelo controlo não bate certo com o `WebPartContext` do SPFx, o código constrói um objeto híbrido com spread e faz cast para `any`:

   ```tsx
   context={{
     ...this.props.context,
     absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
     siteAbsoluteUrl: this.props.context.pageContext.site.absoluteUrl,
     webAbsoluteUrl: this.props.context.pageContext.web.absoluteUrl,
     msGraphClientFactory: this.props.context.msGraphClientFactory,
     spHttpClient: this.props.context.spHttpClient
   } as any}
   ```

   Este bloco está duplicado em `DetailsListPedidos.tsx` e em `FormPedidosCreate.tsx`. Se mexeres num, mexe no outro.
3. **Microsoft Graph** — `this.props.context.msGraphClientFactory.getClient('3')` para as fotos.

As outras cinco props são herança do scaffold do Yeoman.

## Camada de dados

Tudo o que toca no SharePoint passa por `PedidosDataProvider`. É a única classe com conhecimento da lista e dos nomes das colunas.

```ts
const LIST_PEDIDOS = "Pedidos_de_Ferias";
```

Quatro métodos:

| Método | O que faz |
|---|---|
| `getItems()` | `select` das colunas + `expand("Colaborador")` para trazer Id, Title e EMail do utilizador; mapeia cada item para `IPedidos` |
| `createItem(pedido)` | `items.add({ ColaboradorId, Data_Inicio, Data_Fim, Estado })` |
| `updateEstado(pedido)` | `items.getById(id).update({ Estado })` — escreve **só** a coluna Estado |
| `deleteItem(id)` | `items.getById(id).delete()` |

Três detalhes que valem a pena saber:

- **`updateEstado` não está em `IPedidosDataProvider`.** A interface declara apenas `getItems`, `createItem` e `deleteItem`. O método existe na classe concreta e é chamado diretamente. Se fores programar contra a interface, acrescenta-o lá.
- **O campo de pessoa escreve-se como `ColaboradorId`, não `Colaborador`.** É a convenção da REST API do SharePoint para lookups/pessoas: escreve-se o Id, lê-se o objeto expandido.
- **O Id devolvido pelo `add` é defensivo.** `result?.data?.Id || result?.Id || null` cobre as duas formas que o PnPjs já usou em versões diferentes.

O `webPartContext` é uma propriedade com setter, e o setter volta a chamar `getSP(value)` — que, pela guarda do singleton, não faz nada de novo. Continua a ser importante atribuí-lo, porque é assim que o `DetailsListPedidos` e o `FormPedidosCreate` "declaram" a dependência do contexto.

### Tratamento de erros

O padrão no provider é uniforme: `try/catch`, `console.error` com mensagem em português, e re-`throw`. A decisão do que mostrar ao utilizador fica sempre a cargo do componente.

Nos componentes o tratamento diverge:

- `_LoadPedidos()` apanha e **só regista na consola**. Se a lista não existir ou o utilizador não tiver permissões, a grelha aparece vazia sem qualquer explicação — é o ponto mais fraco do tratamento de erros atual.
- `_onApprove`, `_onReject` e `_onDelete` apanham e mostram um `alert()`.
- `_loadUserPhoto()` tem dois `catch` encaixados de propósito: o interno trata o caso normal de "o utilizador não tem foto" (marca `'no-photo'` e segue), o externo apanha falhas de rede ou de permissões.

## O componente da listagem

`DetailsListPedidos` é onde está quase tudo. Vale a pena perceber três mecanismos.

### Seleção

Usa a `Selection` do Fluent UI, guardada num campo da classe (não no estado):

```ts
this._selection = new Selection({
  onSelectionChanged: () => {
    this.setState({ selectionDetails: this._getSelectionDetails() });
  }
});
```

O `_getSelectionDetails()` faz `setState` **e** devolve `undefined`, que depois é atribuído a `selectionDetails` — que não é usado em lado nenhum. O efeito útil é o re-render, e é isso que faz os botões Aprovar/Rejeitar/Apagar aparecerem e desaparecerem. É confuso mas funciona; se mexeres aqui, o que tens de preservar é o re-render.

A visibilidade dos botões é decidida diretamente no `render`, lendo a seleção:

```tsx
{this._selection.getSelectedCount() === 1 && (
   ... (this._selection.getSelection()[0] as IPedidos)?.Estado === 'Pendente' && (Aprovar / Rejeitar)
   ... Apagar (sempre)
)}
```

### Filtro

Dois arrays, propositadamente:

- `this._allItems` — o conjunto completo, atualizado a cada `_LoadPedidos()`
- `this.state.items` — o que está a ser mostrado

O `_applyFilter` filtra `_allItems` por email em minúsculas e escreve o resultado em `state.items`. Quando o PeoplePicker é limpo, o filtro passa a string vazia e o array completo volta. **Não há nova ida ao SharePoint em nenhum destes passos** — tudo acontece em memória sobre o que já foi carregado.

O email vem do PeoplePicker como `secondaryText` (ou `mail`, conforme a origem do resultado), e é comparado com `Colaborador.EMail` do item.

### Fotos via Graph

`state.userPhotos` é um dicionário `email → valor`, onde o valor é um de quatro estados:

| Valor | Significado |
|---|---|
| `undefined` | ainda não se tentou — dispara o carregamento |
| `'loading'` | pedido em curso, não disparar outro |
| `'no-photo'` | já se tentou e não há foto (ou falhou) — não voltar a tentar |
| `blob:...` | object URL válido, pronto a usar no `Persona` |

O ciclo arranca a partir do `onRender` da coluna Colaborador, que chama `_getUserPhotoUrl(email)`. Se o valor for `undefined`, dispara `_loadUserPhoto` e devolve `undefined` nesse render (o `Persona` mostra as iniciais). Quando a foto chega, o `setState` provoca novo render e a imagem aparece.

A máquina de estados existe precisamente para evitar um pedido Graph por cada render de cada linha. O `componentWillUnmount` percorre o dicionário e faz `URL.revokeObjectURL` de todos os blobs, para não deixar memória presa.

### Ordenação

`_copyAndSort` trata `Data_Inicio` e `Data_Fim` como datas (converte para `Date` e compara timestamps) e tudo o resto com comparação genérica. O nome engana: usa `Array.prototype.sort`, que ordena in-place, por isso o array de entrada é modificado. Como é sempre chamado sobre arrays acabados de construir, não causa bugs — mas se reutilizares a função noutro sítio, conta com isso.

A ordenação por omissão, aplicada logo no `_LoadPedidos`, é `Data_Inicio` ascendente.

## O fluxo de criação

1. `CommandBarPedidos` mostra uma `CommandBar` com um item "Novo" e um submenu "Pedido de Férias". Clicar põe `isVisible: true`.
2. Isso abre um `Panel` (`PanelType.smallFluid`) com o `FormPedidosCreate` lá dentro.
3. O `DetailsListPedidos` passa o seu `_LoadPedidos` como prop `_reload` até ao formulário, através da command bar. É assim que a listagem se atualiza sem conhecer o formulário.
4. No submit, `_ValidaFormularioSubmissao()` corre todas as regras e acumula as mensagens num array; se falhar, escreve-as em `state.mensagemErrosCreatePedidos`, põe `_showErrors: true` e abre o diálogo.
5. Se passar, força `Estado = 'Pendente'`, chama `createItem`, depois `_reload()` e por fim `_goBack()` para fechar o painel.

O `_goBack` que chega ao formulário é o `_hidePanel` da command bar. Repara que há um segundo `_hidePanel`, diferente, no `DetailsListPedidos` — esse faz toggle do `isDataLoaded` para mostrar o `ProgressIndicator` e recarrega. São métodos distintos com o mesmo nome em classes diferentes; não os confundas.

### As regras de validação

Todas em `_ValidaFormularioSubmissao`, todas no cliente, nenhuma replicada no servidor:

- colaborador selecionado, com `Title` e `EMail`
- `Data_Inicio` não nula
- `Data_Fim` não nula
- `Data_Inicio <= Data_Fim` (verificado nos dois sentidos, redundante mas inofensivo)
- nenhuma das datas anterior a `hoje`

A última regra tem um problema real: `hoje` é `new Date()`, logo inclui a hora, enquanto o `DatePicker` devolve a data à meia-noite. Escolher o dia de hoje dá sempre "anterior a hoje". A correção é normalizar `hoje` para `00:00:00` antes de comparar.

## Onde ficam as configurações

| O quê | Onde |
|---|---|
| Nome da lista SharePoint | constante `LIST_PEDIDOS` em `PedidosDataProvider.ts` |
| Id e entrypoint do bundle | `config/config.json` |
| Id, versão e metadados da solução | `config/package-solution.json` |
| Id da web part e hosts suportados | `src/webparts/pedidos/PedidosWebPart.manifest.json` |
| Porta e workbench do `gulp serve` | `config/serve.json` |
| Regras de lint | `.eslintrc.js` (perfil `@microsoft/eslint-config-spfx/react`) |
| Opções do compilador | `tsconfig.json` — `strictNullChecks` e `noImplicitAny` estão **desligados** |

O painel de propriedades da web part está deliberadamente vazio (`pages: []`), por isso **não há nada configurável em runtime**. Mudar de lista implica editar código e refazer o build.

## Integrações externas

Duas, ambas autenticadas pelo contexto SPFx, sem chaves nem segredos em lado nenhum:

**SharePoint REST API**, via PnPjs, com o token do utilizador. As permissões efetivas são as permissões do utilizador na lista — se ele não puder escrever, a chamada falha do lado do servidor. Não há elevação de privilégios.

**Microsoft Graph**, via `MSGraphClientFactory`, apenas para `/users/{email}/photo/$value`. Depende das permissões concedidas ao *SharePoint Online Client Extensibility Web Application Principal* no tenant. Como o `package-solution.json` não declara `webApiPermissionRequests`, não há pedido automático de consentimento — tem de ser configurado à mão no Admin Center. Se não estiver, falha de forma controlada e mostra as iniciais.

## Autenticação e permissões

Não há código de autenticação: a web part corre dentro de uma página SharePoint, onde o utilizador já está autenticado, e herda esse contexto.

**Não há autorização na aplicação.** Não existe o conceito de aprovador. Qualquer pessoa que consiga ver a web part vê os botões Aprovar, Rejeitar e Apagar para os pedidos de toda a gente. O que realmente impede alguém de aprovar é só o SharePoint recusar o `update` por falta de permissões na lista — e nesse caso o utilizador vê um `alert` genérico.

Quem quiser implementar controlo de acessos a sério tem dois caminhos, e o segundo é o que conta:

1. Esconder os botões no cliente, consultando a pertença a um grupo SharePoint (é só cosmética, contorna-se).
2. Configurar permissões ao nível da lista — ou permissões ao nível do item, ou uma coluna que só um grupo possa editar — para que o servidor recuse a escrita. É a única barreira efetiva.

## Pontos de partida para continuar

Se vais pegar neste código, estes são os sítios por onde começar:

- **`DetailsListPedidos.tsx`** concentra 613 linhas e faz demasiado: define colunas, gere seleção, filtra, trata das ações de estado e faz chamadas ao Graph. Se for para crescer, a primeira coisa a fazer é extrair as colunas para um módulo próprio e a lógica de fotos para um hook ou serviço separado.
- **`PedidosDataProvider.ts`** é a fronteira certa se precisares de mudar de fonte de dados ou acrescentar queries. Está bem isolado; nenhum componente fala com o PnPjs diretamente.
- **O bloco de contexto do PeoplePicker** está copiado em dois ficheiros. Extrair para um helper evita que divirjam.
- **Não há testes nem infraestrutura para eles.** O `PedidosDataProvider` é a peça mais fácil de testar primeiro, injetando um `SPFI` falso em vez de o ir buscar ao singleton — o que implicaria passá-lo no construtor em vez de chamar `getSP()` lá dentro.
- **`strictNullChecks` está desligado** no `tsconfig.json`, e o código conta com isso (atribui `null` a campos tipados como `Date`, por exemplo). Ligá-lo é possível, mas obriga a rever praticamente todos os ficheiros.
