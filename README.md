# Formatador-de-Relat-rios-Excel

Ferramenta web para aplicar um padrão visual a relatórios em Excel (`.xls` e `.xlsx`). Você escolhe os arquivos, define o título de cada um e baixa as cópias já formatadas. Tudo roda no navegador, sem instalação e sem enviar dados para nenhum servidor.

## Para que serve

Relatórios exportados de sistemas costumam vir com os dados certos, mas sem formatação: colunas estreitas, texto cortado, sem título, sem cabeçalho destacado. Ajustar isso à mão em cada arquivo toma tempo e o resultado varia de pessoa para pessoa.

O Formatador aplica o mesmo padrão em todos os arquivos, em lote, por meio de uma interface simples que qualquer pessoa consegue usar.

## Funcionalidades

- **Processamento em lote** de vários arquivos de uma vez, com arrastar e soltar
- **Suporte a `.xls` e `.xlsx`**, inclusive misturados na mesma seleção (a saída é sempre `.xlsx`)
- **Título personalizado** por arquivo, inserido na primeira linha, mesclado e centralizado
- **Cabeçalho destacado** com fundo cinza, negrito e centralizado
- **Bordas finas** em todas as células
- **Fonte padronizada** (Arial)
- **Quebra de texto automática** e ajuste da altura das linhas conforme o conteúdo
- **Largura das colunas** ajustada de acordo com o tipo de dado do cabeçalho (descrição, data, valor etc.), com opção de manter as larguras originais
- **Alinhamento automático**: números à direita, datas ao centro, texto à esquerda
- **Formatar só a primeira aba ou todas**
- **Download individual ou em `.zip`**
- **Arquivos originais preservados**: o resultado é sempre uma cópia com o sufixo `_FORMATADO.xlsx`

## Como usar

1. Baixe ou clone o repositório.
2. Abra o arquivo `index.html` no navegador (Chrome ou Edge).
3. Arraste os arquivos para a área indicada ou clique em **Escolher arquivos**.
4. Digite o título de cada relatório (opcional; se ficar vazio, o título existente é mantido).
5. Ajuste as opções de formatação, se precisar.
6. Clique em **Formatar arquivos** e baixe os resultados.

Não é preciso servidor, Node.js ou qualquer instalação.

## Estrutura do projeto

```
Formatador_Relatorios/
├── index.html          # Interface
├── style.css           # Estilos da interface
├── app.js              # Lógica de leitura, formatação e geração dos arquivos
├── lib/
│   └── jszip.min.js    # Leitura e escrita do pacote .xlsx
└── LEIA-ME.txt         # Instruções rápidas de uso
```

## Como funciona

Um arquivo `.xlsx` é, na prática, um pacote ZIP com arquivos XML dentro (padrão OOXML). A ferramenta:

1. Abre o pacote com **JSZip**.
2. Lê as planilhas (`xl/worksheets/sheetN.xml`), os textos compartilhados (`sharedStrings.xml`) e os estilos (`styles.xml`).
3. Identifica a linha de título e a linha de cabeçalho.
4. Cria os novos estilos (fonte, preenchimento, borda, alinhamento) e aplica a cada célula de acordo com o papel dela: título, cabeçalho ou corpo.
5. Ajusta colunas, alturas e mesclagem.
6. Gera um novo `.xlsx` para download.

Arquivos `.xls` (formato binário antigo) são convertidos antes para `.xlsx` com a biblioteca **SheetJS**, carregada sob demanda.

## Tecnologias

- HTML, CSS e JavaScript puro (sem frameworks)
- [JSZip](https://stuk.github.io/jszip/) para manipular o pacote `.xlsx`
- [SheetJS](https://sheetjs.com/) para converter `.xls` em `.xlsx`
- APIs nativas do navegador: `DOMParser`, `XMLSerializer`, `Blob`, `FileReader`

## Uso offline

- **`.xlsx`**: funciona 100% offline, pois o JSZip já está incluído na pasta `lib/`.
- **`.xls`**: o conversor SheetJS é buscado primeiro em `lib/xlsx.full.min.js` e, se não existir, em CDNs públicos. Para usar `.xls` sem internet, coloque uma cópia do `xlsx.full.min.js` na pasta `lib/`.

## Limitações

- Para inserir uma linha de título em planilhas que começam direto pelo cabeçalho, o arquivo precisa ser simples. Se houver fórmulas, tabelas, gráficos, tabelas dinâmicas, formatação condicional ou validação de dados, o arquivo é recusado com uma mensagem explicando o motivo, para não quebrar referências. Nesse caso, basta adicionar uma linha vazia no topo pelo Excel antes de formatar.
- A conversão de `.xls` pode não manter macros, imagens, gráficos e formatações avançadas.
- Recomenda-se conferir o resultado no Excel antes de usar em produção.

## Privacidade

Nenhum dado sai do computador. A leitura, a formatação e a geração dos arquivos acontecem inteiramente no navegador.
