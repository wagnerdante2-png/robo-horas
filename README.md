# Robô de Disparos BI → WhatsApp

MVP local para Windows. O robô:

1. abre o relatório do BI no Chrome;
2. exporta/aguarda o Excel do relatório;
3. separa os dados por loja;
4. cruza cada loja com o telefone cadastrado;
5. gera uma mensagem;
6. envia via WhatsApp Web (temporário) ou apenas gera uma prévia;
7. registra logs e salva um Excel separado por loja.

> **Importante:** a camada de WhatsApp foi isolada de propósito. Quando a Zenvia for confirmada, ela pode substituir o WhatsApp Web sem alterar a coleta do BI nem o tratamento do Excel.

## Segurança

- Não existe GitHub Actions neste repositório.
- Não coloque senhas, tokens ou dados reais de clientes/colaboradores no GitHub.
- O modo padrão é `dry_run: true`: prepara tudo, mas **não envia mensagens**.
- O perfil local do Chrome fica em `.rpa_profile/` e é ignorado pelo Git.

## Requisitos

- Windows 10/11
- Google Chrome instalado
- Python 3.11+ recomendado
- acesso ao BI
- acesso ao WhatsApp Web

## Instalação

Execute:

```bat
setup.bat
```

Depois copie os arquivos de exemplo:

```text
config.example.yaml  -> config.yaml
data/lojas.exemplo.csv -> data/lojas.csv
```

Edite `config.yaml` e `data/lojas.csv`.

## Primeiro acesso / autenticação

Para gravar a sessão do BI e do WhatsApp no perfil local do Chrome:

```bat
run.bat --setup-auth
```

Faça login no BI e no WhatsApp Web quando as janelas forem abertas. Depois pressione ENTER no terminal.

## Execução normal

```bat
run.bat
```

### Testar com um Excel já exportado

```bat
run.bat --input "C:\caminho\relatorio.xlsx"
```

### Forçar somente simulação

```bat
run.bat --dry-run
```

### Liberar envio

Primeiro altere no `config.yaml`:

```yaml
whatsapp:
  dry_run: false
```

Depois execute:

```bat
run.bat
```

## Exportação do BI

Existem dois modos.

### 1. manual

O robô abre o relatório e aguarda surgir um novo arquivo `.xlsx` na pasta de downloads. É o modo inicial enquanto não calibrarmos os botões do relatório.

```yaml
bi:
  export:
    mode: manual
```

### 2. selectors

Depois que soubermos exatamente quais botões/menus precisam ser clicados no seu relatório, basta cadastrar as ações no `config.yaml`.

Exemplo:

```yaml
bi:
  export:
    mode: selectors
    actions:
      - by: xpath
        value: "//button[@aria-label='Mais opções']"
      - by: xpath
        value: "//*[contains(text(),'Exportar dados')]"
      - by: xpath
        value: "//*[contains(text(),'Excel')]"
      - by: xpath
        value: "//button[contains(.,'Exportar')]"
```

Os seletores acima são apenas ilustrativos. Power BI pode mudar de acordo com o relatório/visual.

## Cadastro de lojas

`data/lojas.csv`:

```csv
loja,telefone,ativo
ML01,5511999999999,sim
ML02,5511988888888,sim
```

Use DDI + DDD + número, somente dígitos.

## Estrutura esperada do Excel

Por padrão o robô procura a coluna `Loja`. Isso pode ser alterado em:

```yaml
excel:
  store_column: "Loja"
```

Os campos enviados na mensagem são definidos em:

```yaml
message:
  fields:
    - "Venda"
    - "Meta"
    - "Atingimento"
```

Se `fields` ficar vazio, o robô inclui automaticamente as colunas da primeira linha de cada loja, limitado por `max_auto_fields`.

## Saídas

A cada execução:

```text
output/
  downloads/
  lojas/
  previews/
  logs/
```

## Próximo passo

Para transformar a coleta do BI em 100% automática, precisamos apenas do endereço real do relatório e identificar a sequência exata de cliques para exportar o Excel. A integração futura com Zenvia fica concentrada na classe de envio e não exige refazer o restante do robô.
