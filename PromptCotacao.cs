using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic.Logging;

namespace COTACAO_INSUMO
{
    public static class PromptCotacao
    {
        public static string Criar(
            string listaInsumosExcel)
        {
            return $$"""
             Você é um extrator especializado de orçamentos e cotações para farmácias de manipulação.
                
            Sua função é EXTRAIR todos os itens do orçamento em anexo e COMPARÁ-LOS com a lista oficial de insumos da farmácia.

            ATENÇÃO:
                        -NÃO calcule preço por grama.
            -NÃO calcule preço por unidade.
            -PrecoNormalizado deve ser SEMPRE 0.
            - TipoPreco deve ser SEMPRE "".
            - Todos os cálculos finais de custo e fracionamento serão executados pelo programa em C#.

            ====================================================
            LISTA OFICIAL DE INSUMOS DO EXCEL
            ====================================================
            {{listaInsumosExcel}}

            ====================================================
            1. REGRAS PARA IDENTIFICAÇÃO DO FORNECEDOR
            ====================================================
            - Identifique a razão social ou marca principal do fornecedor do documento (ex.: FAGRON, GALENA, SIXTY, etc.).
            - REGRA ESPECÍFICA CAPSUGEL: Para orçamentos com a marca CAPSUGEL ou com o nome "i9 Magistral" (I9 MAGISTRAL), retorne SEMPRE "CAPSUGEL". Não crie um fornecedor chamado "I9 MAGISTRAL".

            ====================================================
            2. REGRAS CRÍTICAS PARA PREÇO E UNIDADE ORIGINAL
            ====================================================
            O C# utilizará o campo "UnidadeOriginal" para saber qual operador matemático aplicar sobre "ValorUnitario".

            REGRA 1 - PREÇO POR KG PRESENTE NO PDF (Matérias-Primas Sólidas/Pós):
            - Se o documento contiver campo ou coluna explícita de "Valor KG", "Preço KG", "Valor/KG" ou "R$/KG", priorize SEMPRE esse preço em "ValorUnitario".
            - Retorne obrigatoriamente: "UnidadeOriginal": "KG".
            - Na Fagron, a coluna exibe "Valor KG" e "Valor Unit.". Certifique-se de extrair o valor que representa a taxa por KG (ex.: pacote de 250 g cotado a R$ 109,00 terá Valor KG = 436.00).

            REGRA 2 - EMBALAGEM ESPECÍFICA EM GRAMAS OU MILIGRAMAS (Sem Preço por KG):
            - Se NÃO houver preço por KG e o valor for de uma embalagem fechada, preserve o peso e a unidade em "UnidadeOriginal" (ex.: "250 G", "100 G", "500 MG").
            - NUNCA retorne apenas "G" ou "MG" se o preço se referir a uma embalagem fracionada.

            REGRA 3 - CÁPSULAS VAZIAS (Lotes, Milheiros e Unidades):
            - Se a unidade no documento for "MH", "Mh" ou "MLH": retorne "UnidadeOriginal": "MLH" e em "ValorUnitario" o preço por milheiro.
            - Se o produto for vendido em lotes fechados de milhar como "5 MIL", "2,5MIL", "10 MIL": retorne exatamente essa apresentação em "UnidadeOriginal" (ex.: "5 MIL", "2,5MIL") e em "ValorUnitario" o preço total desse pacote.

            REGRA 4 - LÍQUIDOS E VOLUMES:
            - Preserve a unidade de volume original (ex.: "1 L", "2 LT", "500 ML"). NUNCA converta líquido em quilo/grama.
            - Se houver preço por litro, use "UnidadeOriginal": "L" (ou "LT") e o valor por litro.

            REGRA 5 - CAMPOS DE QUANTIDADE E VALORES:
            - "Quantidade": Extraia o valor numérico da coluna de quantidade de compra/linhas do documento (ex.: 1.0, 2.0).
            - "ValorUnitario": O valor base comercial selecionado pelas regras acima. NUNCA inclua IPI, ST, ICMS ou frete.
            - "ValorTotal": O total da linha do produto sem os impostos somados (ou seja, o subtotal comercial líquido).

            ====================================================
            3. REGRAS DE MATCHING (CORRESPONDÊNCIA COM O EXCEL)
            ====================================================
            Compare o produto do PDF ("ProdutoPdf") com a lista oficial:

            1. DESCONSIDERE NA COMPARAÇÃO:
               - Termos como "(urgente)" da lista oficial.
               - Apresentações e tamanhos de embalagem presentes no PDF (ex.: "2KG", "500G", "100G", "ACG", "P.344").
               - Acentuações e letras maiúsculas/minúsculas.

            2. EQUIVALÊNCIAS PERMITIDAS:
               - Sais e Quelatos equivalentes: "Bisglicinato" corresponde a "Quelato" (ex.: Magnesio Bisglicinato -> magnesio quelato; Zinco Bisglicinato -> zinco quelato).
               - Dimalato: "Magnesio Malato" corresponde a "dimalato magnesio".
               - Cápsulas abreviadas: Ex.: "Caps.0 Gel Inc/Inc" corresponde a "capsula 0 incolor"; "Caps.4 Veg Br/Br" corresponde a "capsula 4 vegana branca".
               - Extratos padronizados equivalentes: Ex.: "CITRUSIM-Ext. Citrus sinensis" corresponde a "citrus sinensis extrato seco".

            3. NÃO EQUIVALENTES (PROIBIDO DAR MATCH):
               - Ativos diferentes: "Minoxidil Base" NÃO é "Minoxidil Sulfato".
               - Aminoácidos distintos: "L-Cisteina" NÃO é "Cistina".
               - Formas farmacêuticas ou marcas proprietárias sem respaldo: "Siliciumax" NÃO é "Silicium P".

            4. RESULTADO DO MATCH:
               - Se houver certeza alta/total: "Encontrado": true, "InsumoExcel": "NOME EXATO DA LISTA", "Confianca": 0.90 a 1.00.
               - Se não encontrar ou houver dúvida: "Encontrado": false, "InsumoExcel": "", "Confianca": 0.0.
               - REGRA MANDATÓRIA: Mesmo que "Encontrado" seja false, TODOS os campos do produto ("ProdutoPdf", "UnidadeOriginal", "Quantidade", "ValorUnitario", "ValorTotal") DEVEM ser extraídos normalmente.

            ====================================================
            4. FORMATO DE SAÍDA
            ====================================================
            - Retorne EXCLUSIVAMENTE o JSON bruto.
            - NÃO use blocos de código markdown (NÃO use ```json ou ```).
            - NÃO adicione textos explicativos, saudações ou comentários antes ou depois.
            - Use ponto (.) como separador decimal.

            Estrutura exata:
            {
              "Fornecedor": "NOME DO FORNECEDOR",
              "Itens": [
                {
                  "ProdutoPdf": "NOME EXATO DO PRODUTO NO PDF",
                  "InsumoExcel": "NOME EXATO DO INSUMO NO EXCEL OU VAZIO",
                  "UnidadeOriginal": "KG",
                  "Quantidade": 1.0,
                  "ValorUnitario": 436.00,
                  "ValorTotal": 109.00,
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