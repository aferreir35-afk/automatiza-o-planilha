let
    Caminho = Excel.CurrentWorkbook(){[Name = "CaminhoSAP"]}[Content]{0}[Column1],
    Fonte = Excel.Workbook(File.Contents(Caminho), null, true),
    Aba = Fonte{[Item = "Saldo inicial", Kind = "Sheet"]}[Data],
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    Colunas = Table.SelectColumns(Cabecalho, {"Material", "Lote", "Depósito", "UM básica", "Quantidade"}),
    Tipos = Table.TransformColumnTypes(Colunas, {{"Material", type text}, {"Lote", type text}, {"Depósito", type text}, {"UM básica", type text}, {"Quantidade", type number}}),
    SemVazias = Table.SelectRows(Tipos, each [Material] <> null and [Material] <> ""),
    Obs = Table.AddColumn(SemVazias, "Observação", each "SAP 08/09", type text),
    Final = Table.RenameColumns(Obs, {{"Material", "SKU"}, {"UM básica", "UM"}})
in
    Final
