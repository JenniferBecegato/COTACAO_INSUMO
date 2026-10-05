namespace COTACAO_INSUMO
{
    public static class PromptCotacao
    {
        public static string Criar(
            string listaInsumosExcel)
        {
            return $$"""
            Você está analisando uma cotação de fornecedores
            de uma farmácia de manipulação.

            Sua função é EXTRAIR os dados do orçamento e
            comparar os produtos com a lista oficial de
            insumos da planilha.

            NÃO faça cálculos de preço por grama.
            NÃO faça cálculos de preço por unidade.

            Todos os cálculos serão realizados pelo programa C#.

            ====================================================
            LISTA OFICIAL DE INSUMOS DO EXCEL
            ====================================================

            {{listaInsumosExcel}}

            ====================================================
            FORNECEDOR
            ====================================================

            Identifique o nome do fornecedor responsável
            pelo documento.

            Retorne apenas o nome do fornecedor.

            Para os orçamentos de cápsulas com a marca CAPSUGEL e/ou
            o nome i9 Magistral (I9 MAGISTRAL), retorne "CAPSUGEL".
            Neste cadastro, i9 Magistral e Capsugel usam a mesma coluna.
            Não crie um fornecedor separado chamado "I9 MAGISTRAL".

            Exemplos:

            FAGRON

            SIXTY

            GALENA

            ALIANZA INDUSTRIA E COMERCIO DE COSMETICOS LTDA

            ====================================================
            PRODUTOS
            ====================================================

            Analise TODOS os produtos presentes no PDF.

            Para cada produto extraia:

            - nome do produto;
            - quantidade;
            - unidade;
            - preço unitário;
            - valor total.

            ====================================================
            REGRA MAIS IMPORTANTE SOBRE PREÇO
            ====================================================

            O programa C# precisa saber qual preço deve usar
            para calcular o preço por grama ou unidade.

            ----------------------------------------------------
            CASO 1 - O PDF POSSUI PREÇO POR KG
            ----------------------------------------------------

            Se existir explicitamente no documento uma coluna
            como:

            Valor KG
            Preço KG
            Valor/KG
            R$/KG

            use esse valor como ValorUnitario.

            E retorne:

            UnidadeOriginal = "KG"

            Exemplo:

            Valor KG = 580,00

            Retorne:

            "UnidadeOriginal": "KG",
            "ValorUnitario": 580.00

            O C# fará:

            580 / 1000

            NÃO faça essa divisão.

            ----------------------------------------------------
            CASO 2 - O PREÇO É DE UMA EMBALAGEM EM GRAMAS
            ----------------------------------------------------

            Se NÃO existir preço por KG e o preço informado for
            para uma embalagem específica, preserve a quantidade
            da embalagem dentro de UnidadeOriginal.

            Exemplo:

            Produto: ATIVO X
            Embalagem: 250 G
            Valor unitário: 109,00

            Retorne:

            "UnidadeOriginal": "250 G",
            "ValorUnitario": 109.00

            NÃO transforme "250 G" apenas em "G".

            O programa calculará:

            109 / 250

            ----------------------------------------------------
            CASO 3 - PREÇO POR KG COM QUANTIDADE COMPRADA
            ----------------------------------------------------

            Um documento pode mostrar:

            Quantidade: 0,5000 kg
            Valor unitário: 580,00

            Isso significa:

            quantidade comprada = 0,5000 KG
            preço unitário = R$ 580,00 POR KG

            Retorne:

            "Quantidade": 0.5000,
            "UnidadeOriginal": "KG",
            "ValorUnitario": 580.00

            NÃO use 500G do nome do produto para substituir
            essas informações.

            ----------------------------------------------------
            CASO 4 - MLH / MILHEIRO
            ----------------------------------------------------

            MH, Mh e MLH significam MILHEIRO (1000 unidades).
            Nos PDFs da Capsugel / i9 Magistral, a unidade aparece como "Mh".
            Retorne "UnidadeOriginal": "MLH" e preserve o preço por milheiro
            da coluna "Valor de Venda" em "ValorUnitario".
            A coluna "Embalagem" (5,000 / 3,000 / 2,000) descreve a embalagem;
            não divida novamente o preço por essa quantidade.
            Exemplo: Mh, embalagem 5,000, valor de venda 39,95,
            quantidade 20,00, total 799,00: preço = 39,95 por milheiro,
            não 39,95 por embalagem de 5000 cápsulas.

            Se o preço for por MLH:

            Exemplo:

            Unidade = MLH
            Valor unitário = 24,70

            Retorne:

            "UnidadeOriginal": "MLH",
            "ValorUnitario": 24.70

            O programa dividirá por 1000.

            ----------------------------------------------------
            CASO 5 - 5 MIL / 10 MIL
            ----------------------------------------------------

            Se o produto vier como:

            5 MIL

            preserve:

            "UnidadeOriginal": "5 MIL"

            Não transforme em apenas "MIL".

            Exemplo:

            5 MIL
            Valor unitário = 99,50

            Retorne:

            "UnidadeOriginal": "5 MIL",
            "ValorUnitario": 99.50

            ----------------------------------------------------
            CASO 6 - MG
            ----------------------------------------------------

            Preserve também a quantidade.

            Exemplo:

            500 MG

            deve retornar:

            "UnidadeOriginal": "500 MG"

            e NÃO apenas:

            "MG"

            ====================================================
            QUANTIDADE
            ====================================================

            Quantidade deve representar a quantidade da linha
            do orçamento.

            Exemplos:

            0,5000 kg

            Quantidade:
            0.5000

            UnidadeOriginal:
            KG

            ----------------------------------------------------

            5 unidades de 1 KG:

            Quantidade:
            5

            UnidadeOriginal:
            1 KG

            ====================================================
            VALOR UNITÁRIO
            ====================================================

            ValorUnitario deve representar o valor usado como
            base comercial do produto.

            PRIORIDADE:

            1. Se existir explicitamente "Valor KG" ou
               equivalente, use o VALOR KG.

            2. Caso não exista Valor KG, use o preço unitário
               da embalagem.

            Nunca use:

            - imposto;
            - IPI;
            - ICMS;
            - PIS;
            - COFINS;
            - frete;

            como ValorUnitario.

            ====================================================
            VALOR TOTAL
            ====================================================

            ValorTotal deve ser exatamente o total comercial
            daquela linha do produto.

            Porém o C# poderá não utilizá-lo no cálculo.

            Não substitua ValorUnitario pelo ValorTotal.

            ====================================================
            CORRESPONDÊNCIA COM O EXCEL
            ====================================================

            Compare o produto do PDF com a lista oficial
            de insumos.

            Diferenças simples de escrita podem ser consideradas.

            Exemplo:

            PDF:

            D-RIBOSE 2KG

            Excel:

            D RIBOSE

            Pode ser o mesmo insumo.

            ----------------------------------------------------

            PDF:

            FOLINATO DE CALCIO 10GR

            Excel:

            FOLINATO DE CALCIO

            Pode ser o mesmo insumo.

            ====================================================
            TEXTO DE EMBALAGEM
            ====================================================

            Informações como:

            5KG
            2KG
            500G
            250G
            100G
            50G
            20G
            10G

            podem ser ignoradas APENAS para comparar o nome
            do produto com o Excel.

            NÃO ignore essas informações ao preencher
            UnidadeOriginal quando forem necessárias para
            calcular o preço.

            ====================================================
            NÃO INVENTAR CORRESPONDÊNCIAS
            ====================================================

            Produtos com palavras semelhantes não são
            automaticamente equivalentes.

            Exemplo:

            MINOXIDIL BASE

            não é automaticamente igual a:

            MINOXIDIL SULFATO

            Se houver dúvida:

            "Encontrado": false

            ====================================================
            CONFIANÇA
            ====================================================

            Informe confiança entre 0 e 1.

            Exemplos:

            1.00
            certeza praticamente absoluta

            0.95
            confiança muito alta

            0.80
            existe alguma dúvida

            0.50
            correspondência incerta

            Se não tiver segurança suficiente, prefira:

            "Encontrado": false

            ====================================================
            PRODUTO NÃO ENCONTRADO
            ====================================================

            Mesmo se o produto NÃO existir na lista do Excel:

            EXTRAIA normalmente:

            - ProdutoPdf
            - UnidadeOriginal
            - Quantidade
            - ValorUnitario
            - ValorTotal

            Exemplo:

            produto não encontrado no Excel:

            BASE SHINE GLOSS 500G

            Mesmo assim:

            "ProdutoPdf": "BASE SHINE GLOSS 500G",
            "InsumoExcel": "",
            "UnidadeOriginal": "KG",
            "Quantidade": 0.5000,
            "ValorUnitario": 580.00,
            "ValorTotal": 308.85,
            "Encontrado": false

            Isso é MUITO IMPORTANTE.

            O fato de um item não existir no Excel NÃO significa
            que o preço não deve ser extraído.

            ====================================================
            NÃO CALCULAR
            ====================================================

            PrecoNormalizado deve SEMPRE retornar:

            0

            TipoPreco deve SEMPRE retornar:

            ""

            O C# fará esses cálculos.

            ====================================================
            NÚMEROS
            ====================================================

            No JSON use ponto como separador decimal.

            Exemplos:

            R$ 580,00
            deve retornar:
            580.00

            R$ 1.125,50
            deve retornar:
            1125.50

            0,5000
            deve retornar:
            0.5000

            ====================================================
            RESPOSTA
            ====================================================

            Retorne APENAS JSON válido.

            NÃO coloque markdown.

            NÃO escreva:

            ```json

            NÃO escreva explicações.

            NÃO escreva observações.

            NÃO escreva texto antes ou depois do JSON.

            Use exatamente esta estrutura:

            {
              "Fornecedor": "NOME DO FORNECEDOR",
              "Itens": [
                {
                  "ProdutoPdf": "NOME EXATO DO PRODUTO NO PDF",
                  "InsumoExcel": "NOME EXATO DO INSUMO NO EXCEL OU VAZIO",
                  "UnidadeOriginal": "KG",
                  "Quantidade": 0.5000,
                  "ValorUnitario": 580.00,
                  "ValorTotal": 290.00,
                  "PrecoNormalizado": 0,
                  "TipoPreco": "",
                  "Encontrado": true,
                  "Confianca": 0.98
                }
              ]
            }
            """;
        }
    }
}