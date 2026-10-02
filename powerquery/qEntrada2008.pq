let
    Caminho = Excel.CurrentWorkbook(){[Name = "CaminhoSAP"]}[Content]{0}[Column1],
    Fonte = Excel.Workbook(File.Contents(Caminho), null, true),
    DataLimite = Date.From(Excel.CurrentWorkbook(){[Name = "DataLimiteRecebimento"]}[Content]{0}[Column1]),
    Aba = Fonte{[Item = "Entrada ", Kind = "Sheet"]}[Data],
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    Tipos = Table.TransformColumnTypes(Cabecalho, {{"Doc.material", type text}, {"Material", type text}, {"Lote", type text}, {"Depósito", type text}, {"Centro", type text}, {"Data de lançamento", type date}, {"Quantidade", type number}, {"UM básica", type text}, {"Data de entrada", type date}, {"Hora do registro", type time}}),
    SemTotais = Table.SelectRows(Tipos, each [Material] <> null and [Material] <> ""),
    Movimentacao = Table.AddColumn(SemTotais, "Movimentação", each if [Quantidade] > 0 and [Data de lançamento] <= DataLimite then "Recebido no 2008" else if [Quantidade] < 0 then "Entrada" else "Estorno", type text),
    EntrouZBR = Table.AddColumn(Movimentacao, "Entrou no ZBR", each if [Movimentação] = "Recebido no 2008" then 0 else -[Quantidade], type number),
    Final = Table.RenameColumns(EntrouZBR, {{"Material", "SKU"}, {"UM básica", "UM"}})
in
    Final
