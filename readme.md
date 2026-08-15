# Bora Mermã — Gerenciamento da Loja

Aplicação desktop (Windows) para controle de estoque, vendas e financeiro de uma loja, integrada **em tempo real** com a planilha Excel que o cliente já usa no dia a dia — sem precisar trocar de ferramenta ou duplicar dados.

## Como funciona

O app não usa banco de dados próprio. Ele controla a planilha do cliente diretamente, via **COM Interop com o Excel** (a mesma tecnologia de automação do VBA), então:

- Qualquer entrada de estoque, venda ou atualização feita pelo app é gravada direto na planilha.
- Se o cliente mexer na planilha por fora (Excel aberto ao mesmo tempo), os dados continuam consistentes — é o mesmo arquivo, sem sincronização "por trás".
- Não é necessário migrar o histórico: o app se conecta à estrutura de colunas que o cliente já usa.

## Módulos

| Menu | Tela | O que faz |
|---|---|---|
| — | Dashboard | Espaço reservado para o dashboard principal (customizado à parte) |
| Produtos | Produtos disponíveis | Lista o estoque disponível (peças ainda não vendidas) e permite registrar entrada de peças novas ou dar baixa |
| Vendas | Vendas realizadas | Lista todas as vendas já registradas, com total vendido |
| Vendas | Adicionar venda | Marca uma peça em estoque como vendida (valor recebido, cliente) |
| Financeiro | Balanço geral | Soma de entradas (vendas) e saídas (custo de reposição) agrupadas por dia, semana, mês ou ano |

## Conectando com a planilha do cliente

Como cada planilha é organizada de um jeito diferente, a conexão é **assistida**:

1. Selecionar o arquivo `.xlsx`/`.xlsm`.
2. Escolher a aba que tem o controle de peças.
3. Confirmar qual linha é o cabeçalho (o app já sugere automaticamente a linha com mais colunas preenchidas).
4. Mapear cada campo do sistema (SKU, produto, categoria, status, datas, valores, cliente) para a coluna correspondente na planilha do cliente.

Esse mapeamento fica salvo localmente, então nas próximas vezes o app já abre direto conectado.

### Modelo de dados

Cada **linha da planilha representa uma peça individual** (não uma quantidade agregada de um produto), com:

- Data de entrada no estoque
- Status (vazio = em estoque, "VENDIDO" = vendida)
- Data e valor da venda, quando vendida

O estoque disponível, as vendas e o balanço financeiro são todos derivados dessas linhas — sem precisar de nenhuma aba extra.

## Requisitos para rodar

- Windows com **Microsoft Excel instalado** (a integração usa automação COM do Excel).
- Nada mais — o executável publicado é *self-contained* (inclui o runtime do .NET), não precisa instalar o .NET separadamente.

## Rodando o projeto (desenvolvimento)

```bash
cd GerenciamentoLoja
dotnet run
```

## Gerando um executável para distribuir

```bash
cd GerenciamentoLoja
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

O executável final fica em `GerenciamentoLoja/bin/Release/net10.0-windows/win-x64/publish/GerenciamentoLoja.exe` — é um arquivo único, o cliente só precisa copiar e abrir.

## Stack técnica

- **.NET 10 / WPF**, com [Material Design in XAML](https://github.com/MaterialDesignInXAML/MaterialDesignInXamlToolkit) para a interface.
- **MVVM** com [CommunityToolkit.Mvvm](https://github.com/CommunityToolkit/dotnet) (`ObservableObject`, `RelayCommand`).
- **COM Interop com late binding (`dynamic`)** para controlar o Excel, evitando depender de Primary Interop Assemblies (PIAs) que só existem instaladas via MSI tradicional do Office.

### Estrutura de pastas

```
GerenciamentoLoja/
  Models/         Entidades de domínio (ItemEstoque, PlanilhaMapeamento, BalancoPeriodo...)
  Services/       Acesso à planilha (ExcelInteropService) e regras de negócio (LojaDataService)
  ViewModels/     Um ViewModel por tela, MVVM
  Views/          XAML das telas
  Converters/     Conversores de binding usados nas Views
```
