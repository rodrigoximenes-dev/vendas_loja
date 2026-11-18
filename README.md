# Relatório Mensal de Vendas

Transforma a planilha de vendas de uma loja pequena em um relatório pronto, em um comando.

> Os dados deste repositório são **fictícios** (loja de artesanato inventada). Nenhuma venda real é usada.

## O problema

Todo pequeno comércio anota as vendas em planilha, mas poucos têm tempo de transformar essas linhas em respostas:

- Quanto vendi no mês?
- Quais produtos puxam o faturamento?
- Qual foi o melhor dia?
- Estou vendendo mais ou menos que no mês passado?

Fazer isso à mão, todo mês, é repetitivo e sujeito a erro. Este script faz em segundos.

## O que o relatório mostra

- **Total do mês** (R$), itens vendidos, dias com venda e média por dia
- **Comparação com o mês anterior**: total e itens, com a variação em % (se a planilha não tiver o mês anterior, o relatório avisa)
- **Produtos mais vendidos** (ranking por quantidade, com o valor de cada um)
- **Dia de maior faturamento**, com o dia da semana

## Formatos de saída

| Formato | Para quê | Conteúdo |
|---|---|---|
| **PDF** | Enviar ao cliente | Relatório de 2 páginas com texto explicativo e tabelas |
| **Excel** | O cliente editar e reaproveitar | 4 abas: Resumo, Produtos, Por dia (com gráfico) e Vendas do mês (com filtro) |
| **Texto** | Conferência rápida na tela | Resumo em texto simples |

### Padrão de apresentação (ABNT)

Os arquivos PDF e Excel não são dados "jogados" na página. Seguem as normas de apresentação da ABNT (NBR 14724 e NBR 6024) e as normas de apresentação tabular do IBGE, no que cabe a um relatório curto:

- Fonte Arial, tamanho 12 no texto e 10 em tabelas, fontes e paginação
- Papel A4, margens de 3 cm (superior e esquerda) e 2 cm (inferior e direita)
- Espaçamento de 1,5 entre linhas no texto, com texto justificado e recuo de parágrafo de 1,25 cm
- Título centralizado; seções numeradas, em maiúsculas e negrito
- Tabelas numeradas (`Tabela 1 – ...`), com título acima, fonte indicada abaixo (`Fonte: ...`) e apenas filetes horizontais
- Número da página no canto superior direito
- Ortografia e acentuação revisadas, com valores em reais no formato brasileiro (`R$ 1.234,56`)

Este relatório não é um trabalho acadêmico completo: não tem capa, sumário ou referências, que a ABNT exige apenas nesse tipo de documento.

## Exemplo de saída (texto)

```text
RELATÓRIO DE VENDAS – Outubro de 2026
============================================
Total do mês:      R$ 2.039,00
Itens vendidos:    97
Dias com venda:    24
Média por dia:     R$ 84,96

Comparação com o mês anterior (setembro de 2026):
  Total:    R$ 1.151,00 -> R$ 2.039,00  (+77,2%)
  Itens:    54 -> 97  (+79,6%)

Produtos mais vendidos (top 3):
  1. Sabonete artesanal      30 un.  R$ 360,00
  2. Chaveiro de crochê      18 un.  R$ 144,00
  3. Vela aromática          15 un.  R$ 420,00

Dia mais forte:
  10/10/2026 (sábado) - R$ 347,00
```

## Como rodar

**Requisito:** Python 3.8 ou superior. O relatório em texto não precisa de nenhuma biblioteca. Para Excel e PDF, instale uma vez:

```bash
pip install openpyxl reportlab
```

Depois, rode:

```bash
python relatorio.py
```

Isso lê `vendas.csv` e gera o relatório do mês mais recente da planilha, na tela e no arquivo `relatorio.txt`.

Para gerar texto, Excel e PDF de uma vez:

```bash
python relatorio.py --formato todos
```

### Opções

| Opção | O que faz | Exemplo |
|---|---|---|
| `arquivo` | Planilha a ler (padrão: `vendas.csv`) | `python relatorio.py minhas_vendas.csv` |
| `--mes AAAA-MM` | Escolhe o mês do relatório | `--mes 2026-10` |
| `--top N` | Quantos produtos listar no ranking (padrão: 3) | `--top 5` |
| `--formato F` | `txt` (padrão), `xlsx`, `pdf` ou `todos` | `--formato pdf` |
| `--saida NOME` | Nome base dos arquivos gerados (padrão: `relatorio`) | `--saida outubro` |

## Formato da planilha

Um CSV com estas quatro colunas, na primeira linha (maiúsculas e minúsculas tanto faz):

| Coluna | Formato | Exemplo |
|---|---|---|
| `data` | AAAA-MM-DD | `2026-10-17` |
| `produto` | texto | `Vela aromática` |
| `quantidade` | número inteiro | `3` |
| `valor` | valor total da linha, em R$ | `84.00` ou `84,00` |

Pontos importantes:

- `valor` é o **total da venda** (quantidade × preço), não o preço unitário.
- Aceita separador por vírgula ou ponto e vírgula (o padrão do Excel em português).
- Aceita valores no formato brasileiro (`1.234,56`).
- Para gerar o CSV no Excel: **Arquivo → Salvar como → CSV UTF-8** (assim os acentos ficam corretos).

### Se a planilha tiver problema

O script avisa com uma mensagem clara, sem travar com erro técnico:

```text
Erro: A planilha não tem a coluna obrigatória: valor.
Colunas encontradas: data, produto, quantidade.
Colunas esperadas: data, produto, quantidade, valor.
```

Linhas com data, quantidade ou valor inválidos são ignoradas, e o script diz qual foi e por quê:

```text
Aviso: linha 3 ignorada (data inválida ('31/10/2026'); use AAAA-MM-DD).
```

## Testes

```bash
python -m unittest -v
```

São 12 testes simples, sem bibliotecas extras: leitura da planilha (coluna faltando, separador, linha inválida), cálculos (totais, variação, virada de ano) e geração dos arquivos Excel e PDF.

## Arquivos

| Arquivo | Função |
|---|---|
| `relatorio.py` | O script |
| `test_relatorio.py` | Testes automatizados |
| `vendas.csv` | Planilha de exemplo (49 vendas fictícias de setembro e outubro de 2026) |
| `relatorio.txt`, `.xlsx`, `.pdf` | Relatórios gerados a partir do exemplo |

## Precisa de algo parecido para o seu negócio?

Posso adaptar o relatório para a sua planilha: outras colunas, outros indicadores, envio automático por e-mail.

*Contato: preencher com o link do perfil (Workana/99Freelas/LinkedIn) antes de publicar.*
