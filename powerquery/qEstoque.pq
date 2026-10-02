let
    Caminho = Excel.CurrentWorkbook(){[Name = "CaminhoSAP"]}[Content]{0}[Column1],
    Fonte = Excel.Workbook(File.Contents(Caminho), null, true),
    Aba = Fonte{[Item = "Estoque atual", Kind = "Sheet"]}[Data],
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    Tipos = Table.TransformColumnTypes(Cabecalho, {{"Material", type text}, {"Lote", type text}, {"Tipo de depósito", type text}, {"Posição no depósito", type text}, {"Estoque total", type number}, {"Último movimento", type date}, {"Depósito", type text}, {"Data do vencimento", type date}}),
    SemVazia = Table.RemoveColumns(Tipos, {"Inventário ativo"}, MissingField.Ignore),
    Primeiras = {"Material", "Lote", "Tipo de depósito", "Posição no depósito"},
    Ordenado = Table.ReorderColumns(SemVazia, Primeiras & List.RemoveItems(Table.ColumnNames(SemVazia), Primeiras)),
    Final = Table.RenameColumns(Ordenado, {{"Material", "SKU"}, {"Estoque total", "Estoque físico"}, {"UM básica", "UM"}})
in
    Final
