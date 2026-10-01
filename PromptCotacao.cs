namespace COTACAO_INSUMO
{
    public static class PromptCotacao
    {
        public static string Criar(
            string listaInsumosExcel)
        {
            return """
            Você está analisando uma cotação de fornecedores
            de uma farmácia de manipulação.

            =========================================
            INSUMOS EXISTENTES NA PLANILHA
            =========================================

            """ + listaInsumosExcel + """

            =========================================
            SUA TAREFA
            =========================================

            Analise o PDF da cotação anexado.

            Para CADA produto do orçamento:

            1. Leia o nome exatamente como aparece no PDF.

            2. Identifique:
               - produto;
               - unidade;
               - quantidade;
               - valor unitário;
               - valor total.

            3. Identifique o fornecedor responsável pelo orçamento.

            4. Compare o produto do PDF com a lista oficial
               de insumos da planilha.

            5. Considere diferenças simples de escrita.

            Exemplos:

            PDF:
            D-RIBOSE 2KG

            PLANILHA:
            D RIBOSE

            É o mesmo produto.

            -----------------------------------------

            PDF:
            FOLINATO DE CALCIO 10GR

            PLANILHA:
            FOLINATO DE CALCIO

            É o mesmo produto.

            -----------------------------------------

            Remova mentalmente informações comerciais como:

            5KG
            500GR
            20GR
            10GR
            2KG

            quando essas informações forem apenas
            embalagem/quantidade.

            -----------------------------------------

            ATENÇÃO:

            NÃO considere dois produtos iguais apenas porque
            possuem palavras semelhantes.

            Exemplo:

            MINOXIDIL BASE

            NÃO deve ser automaticamente considerado igual a:

            MINOXIDIL SULFATO

            Em caso de dúvida:

            Encontrado = false

            -----------------------------------------

            UNIDADES:

            KG
            G
            GR
            MG

            são produtos cujo preço será posteriormente
            normalizado por GRAMA pelo programa.

            MLH significa MILHEIRO.

            Para produtos/cápsulas em MLH:

            UnidadeOriginal = "MLH"

            O programa C# calculará posteriormente
            o preço por UNIDADE.

            NÃO faça você mesmo essa conversão.

            -----------------------------------------

            CONFIANÇA:

            Informe uma confiança entre 0 e 1.

            1.00 = certeza
            0.90 = confiança muito alta
            0.50 = dúvida

            Caso não tenha confiança suficiente,
            use Encontrado = false.

            -----------------------------------------

            IMPORTANTE:

            PrecoNormalizado deve permanecer 0.
            TipoPreco deve permanecer vazio.

            Esses valores serão calculados pelo C#,
            e não pela inteligência artificial.

            =========================================
            FORMATO OBRIGATÓRIO
            =========================================

            Retorne SOMENTE um JSON válido.

            Não coloque ```json.
            Não coloque comentários.
            Não coloque explicações antes ou depois.

            Exatamente nesta estrutura:

            {{
              "Fornecedor": "NOME DO FORNECEDOR",
              "Itens": [
                {{
                  "ProdutoPdf": "NOME COMO APARECE NO PDF",
                  "InsumoExcel": "NOME EXATO ENCONTRADO NA PLANILHA",
                  "UnidadeOriginal": "KG",
                  "Quantidade": 0.500,
                  "ValorUnitario": 1750.00,
                  "ValorTotal": 875.00,
                  "PrecoNormalizado": 0,
                  "TipoPreco": "",
                  "Encontrado": true,
                  "Confianca": 0.98
                }}
              ]
            }}
            """;
        }
    }
}