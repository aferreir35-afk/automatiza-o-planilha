let
    Caminho = Excel.CurrentWorkbook(){[Name = "CaminhoSAP"]}[Content]{0}[Column1],
    Fonte = Excel.Workbook(File.Contents(Caminho), null, true),
    Aba = Fonte{[Item = "Saida", Kind = "Sheet"]}[Data],
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    Tipos = Table.TransformColumnTypes(Cabecalho, {{"Doc.material", type text}, {"Material", type text}, {"Lote", type text}, {"Depósito", type text}, {"Centro", type text}, {"Data de lançamento", type date}, {"Quantidade", type number}, {"Data de entrada", type date}, {"Tipo de movimento", type text}, {"Referência", type text}, {"Cliente", type text}}),
    SemTotais = Table.SelectRows(Tipos, each [Material] <> null and [Material] <> ""),
    Saiu = Table.AddColumn(SemTotais, "Saiu", each -[Quantidade], type number),
    Final = Table.RenameColumns(Saiu, {{"Material", "SKU"}, {"Texto breve de material", "Descrição"}, {"UM básica", "UM"}})
in
    Final
