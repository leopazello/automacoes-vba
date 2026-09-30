# Exemplos com dados fictícios

## 1. Roteiro de entrada

Crie uma aba chamada `Roteiro` e reproduza a tabela abaixo a partir da célula A1. Nas linhas com falas, use horários reconhecidos pelo Excel, como `00:00:01`.

| Timecode | Personagem | Fala |
| --- | --- | --- |
| EP01 | | |
| 00:00:01 | Ana | Precisamos chegar antes do amanhecer. |
| 00:00:04 | Bruno | O portal fica perto daqui. |
| 00:00:08 | Ana | Vamos seguir juntos. |
| EP02 | | |
| 00:00:01 | Ana | Encontrei a pedra azul. |
| 00:00:05 | Bruno | O portal esta aberto. |
| 00:00:09 | Clara | Eu cuido da passagem. |

As linhas `EP01` e `EP02` têm as colunas B e C vazias. O texto sem acentos em algumas falas é intencional para simplificar a reprodução do exemplo.

### Escala por episódio

Selecione a aba `Roteiro` e execute `Escala_varios_eps_por_doc_e_contagem_de_aparicoes`.

Resultado esperado na aba `Escala`:

| Personagem | Palavras | Episódios |
| --- | ---: | --- |
| Ana | 12 | EP01, EP02 |
| Bruno | 9 | EP01, EP02 |
| Clara | 4 | EP02 |

Conferência: Ana tem 5 + 3 + 4 palavras; Bruno, 5 + 4; Clara, 4.

### Contagem simples

Volte à aba `Roteiro`, execute `Escala_Simples_Contagem_Palavras` e informe `C` ou `3`. Confirme a substituição de `Escala` caso tenha executado o exemplo anterior.

Resultado esperado:

| Personagem | Palavras |
| --- | ---: |
| Ana | 12 |
| Bruno | 9 |
| Clara | 4 |

## 2. Glossário

Crie uma aba chamada `GLOSSARIO` e preencha a partir de A1:

| Termo de origem | Substituição |
| --- | --- |
| portal | portao |
| pedra azul | joia azul |

Volte à aba `Roteiro` e execute `SubstituirPalavrasDoGlossario`. A macro processa C e D, mesmo quando D está vazia.

| Texto original | Texto esperado |
| --- | --- |
| O portal fica perto daqui. | O **portao** fica perto daqui. |
| Encontrei a pedra azul. | Encontrei a **joia azul**. |
| O portal esta aberto. | O **portao** esta aberto. |

As demais falas permanecem iguais. Os trechos substituídos ficam em negrito no Excel.

A macro substitui os termos sem reescrever o restante da frase. Quando uma troca exige mudanças de artigo, concordância ou contexto, esses ajustes precisam ser feitos na revisão.

## Conferência manual no Excel

- Execute a escala por episódio e compare os três totais e as listas de episódios.
- Execute a contagem simples com `C` e com `3`; compare os resultados.
- Cancele a pergunta da coluna: nenhuma aba deve ser substituída.
- Recuse a substituição de uma aba `Escala` existente: ela deve ser preservada.
- Teste uma planilha só com cabeçalho: a escala deve exibir os cabeçalhos sem tentar formatar linhas de dados inexistentes.
- Execute o glossário e confira texto e negrito nas três células alteradas.
- Coloque uma fórmula em C ou D e confirme que o glossário não altera essa célula.

Este roteiro de conferência está disponível para reprodução. Não representa testes já executados no Excel durante a preparação do repositório.
