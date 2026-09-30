# Automações em VBA para Excel

## Ferramentas

| Ferramenta | Problema | Solução |
| --- | --- | --- |
| Escala por personagem e episódio | Reunir a quantidade de palavras e os episódios de cada personagem exige percorrer o roteiro. | Conta as palavras por personagem e lista os episódios em uma nova tabela. |
| Contagem com coluna selecionável | Diferentes planilhas podem ter as falas em colunas diferentes. | Permite selecionar a coluna e gera uma tabela ordenada com os totais. |
| Aplicação de glossário | Aplicar um conjunto de termos manualmente exige várias buscas e substituições. | Lê um glossário, substitui termos nas colunas C e D e destaca as alterações em negrito. |

## Exemplo rápido

Com o [roteiro fictício de demonstração](EXEMPLOS.md), a escala por episódio deve produzir:

| Personagem | Palavras | Episódios |
| --- | ---: | --- |
| Ana | 12 | EP01, EP02 |
| Bruno | 9 | EP01, EP02 |
| Clara | 4 | EP02 |

Exemplo de aplicação do glossário:

- Texto original: `O portal fica perto daqui.`
- Entrada no glossário: `portal` → `portao`
- Texto esperado: O **portao** fica perto daqui.

A substituição segue o glossário e exige revisão humana de concordância e contexto.

Os resultados acima são exemplos esperados, conferidos a partir dos dados fictícios. Não são capturas de uma execução no Excel.

## Código e recursos utilizados

- [Escalas.bas](Escalas.bas): contagem e agrupamento por personagem, associação com episódios, ordenação e formatação da tabela.
- [Glossario.bas](Glossario.bas): leitura do glossário e substituições com destaque visual.
- VBA e o modelo de objetos do Excel para leitura e alteração de células.
- `Scripting.Dictionary` para agrupamento de dados e consulta de termos.
- `VBScript.RegExp` para contagem de sequências separadas por espaços e busca de termos.

As macros usam `CreateObject` (late binding), sem exigir que as referências sejam marcadas manualmente no editor VBA. O ambiente previsto é Excel desktop no Windows com esses componentes disponíveis.

## Como executar

1. Abra uma **cópia** da planilha no Excel desktop e salve como **Pasta de Trabalho Habilitada para Macro do Excel (`.xlsm`)**.
2. Abra o editor VBA com `Alt + F11`.
3. Em **Arquivo > Importar arquivo**, importe `Escalas.bas` e `Glossario.bas` para essa pasta de trabalho.
4. Volte ao Excel e selecione a aba com os dados de entrada.
5. Use `Alt + F8` e execute a macro desejada.

| Macro | Entrada |
| --- | --- |
| `Escala_varios_eps_por_doc_e_contagem_de_aparicoes` | Coluna A: horário ou valor numérico nas linhas de fala e marcadores como `EP01` nas linhas de episódio. Coluna B: personagem. Coluna C: fala. |
| `Escala_Simples_Contagem_Palavras` | Coluna B: personagem. A coluna das falas é informada ao executar a macro, por letra ou número. |
| `SubstituirPalavrasDoGlossario` | Aba `GLOSSARIO` ou `GLOSSÁRIO`, com cabeçalho na linha 1, termo de origem na coluna A e substituição na B. Aplica as trocas nas colunas C e D da aba selecionada. |

As duas macros de escala pedem confirmação antes de substituir uma aba `Escala` existente na mesma pasta de trabalho. A macro de glossário altera o texto das células de entrada; use uma cópia para experimentar.

## Comportamento e limites

- A contagem considera cada sequência de caracteres sem espaço (`\S+`) como uma palavra. Pontuação isolada também pode entrar na contagem.
- A escala por episódio lista os episódios em que há texto para o personagem. Ela não conta aparições visuais nem a quantidade de cenas.
- Linhas de fala anteriores ao primeiro marcador `EP` são associadas ao episódio `Indefinido`.
- Na escala por episódio, o horário deve ser reconhecido pelo Excel como data/hora ou número. Não há interpretação própria de timecodes com frames.
- A contagem simples não verifica horários. Use linhas de cabeçalho com `Personagem` ou `Personagens` e deixe a coluna B vazia nas linhas de separação.
- Os nomes dos personagens são agrupados de forma literal após remover espaços nas extremidades. Grafias e maiúsculas diferentes podem gerar grupos diferentes.
- O glossário ignora diferenças de maiúsculas na busca, mas aplica exatamente a grafia do texto de substituição.
- As substituições seguem a ordem do glossário. Termos sobrepostos ou substituições que criem outro termo do glossário precisam de revisão.
- A busca do glossário usa os limites `\b` do mecanismo VBScript, que têm limitações com caracteres acentuados nas extremidades dos termos.
- Células com fórmulas são ignoradas pelo glossário. Uma falha durante a execução pode deixar substituições parciais.
- As macros não fazem tradução automática, não chamam serviços de IA e não enviam dados pela rede. A IA foi utilizada como apoio ao desenvolvimento.

## Sobre esta publicação

Esta versão organiza os códigos para consulta e importação. Inclui ajustes no tratamento de erros, na seleção da coluna e na criação da aba de resultados. Os exemplos usam somente dados fictícios.

A preparação desta publicação incluiu revisão estática do código e conferência dos exemplos. A execução e a compatibilidade desta versão ainda precisam ser verificadas no Excel desktop; não há teste automatizado de integração com o Excel.

Autor: [Leonardo Pazello](https://github.com/leopazello).
