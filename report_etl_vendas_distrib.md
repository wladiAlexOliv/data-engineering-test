---### produtos

let

Fonte = Excel.Workbook(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\Produtos.xlsx"), null, true),

produtos\_Sheet = Fonte{[Item="produtos",Kind="Sheet"]}[Data],

#"Tipo Alterado" = Table.TransformColumnTypes(produtos\_Sheet,{{"Column1", type text}}),

#"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado",{{"Column1", "cod\_prod"}, {"Column2", "PRODUTO"}})

in

#"Colunas Renomeadas"

---### EstadosBra

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\EstadosBra.csv"),[Delimiter=";", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None]),

#"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Column1", type text}, {"Column2", type text}}),

#"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado",{{"Column1", "UNIDADE FEDERATIVA"}, {"Column2", "UN. DA FEDERAÇÃO"}})

in

#"Colunas Renomeadas"

---### oleo combustivel

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-oleo-combustivel-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### oleo diesel

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-oleo-diesel-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "}),

#"Consulta Acrescentada" = Table.Combine({#"Colunas Removidas", #"vendas-oleo-combustivel-m3-1990-2020"}),

#"Consulta Acrescentada1" = Table.Combine({#"Consulta Acrescentada", #"vendas-oleo-combustivel-m3-1990-2020"})

in

#"Consulta Acrescentada1"

---### etanol hidratado

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-etanol-hidratado-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### gasolina aviacao

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-etanol-hidratado-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### gasolina C

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-gasolina-c-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### GLP

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-glp-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### querosene aviacao

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-querosene-aviacao-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### querosene iluminante

let

Fonte = Csv.Document(File.Contents("C:\wladiston\Wladi\Docs\cv\Raizen\anp\venda\mesAno\vendas-querosene-aviacao-m3-1990-2020.csv"),[Delimiter=";", Columns=19, Encoding=65001, QuoteStyle=QuoteStyle.None]),

#"Cabeçalhos Promovidos" = Table.PromoteHeaders(Fonte, [PromoteAllScalars=true]),

#"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"COMBUSTÍVEL ", type text}, {" ANO ", Int64.Type}, {" REGIÃO ", type text}, {" ESTADO ", type text}, {" UNIDADE ", type text}, {" JANEIRO ", type number}, {" FEVEREIRO ", type number}, {" MARÇO ", type number}, {" ABRIL ", type number}, {" MAIO ", type number}, {" JUNHO ", type number}, {" JULHO ", type number}, {" AGOSTO ", type number}, {" SETEMBRO ", type number}, {" OUTUBRO ", type number}, {" NOVEMBRO ", type number}, {" DEZEMBRO ", type number}, {" TOTAL ", type number}, {" ", type text}}),

#"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{" REGIÃO ", " UNIDADE ", " "})

in

#"Colunas Removidas"

---### vendas comb distribuidoras

let

Fonte = Table.Combine({#"vendas-querosene-iluminante-m3-1990-2020", #"vendas-oleo-combustivel-m3-1990-2020", #"vendas-oleo-diesel-m3-1990-2020", #"vendas-etanol-hidratado-m3-1990-2020", #"vendas-gasolina-aviacao-m3-1990-2020", #"vendas-gasolina-c-m3-1990-2020", #"vendas-glp-m3-1990-2020", #"vendas-querosene-aviacao-m3-1990-2020"}),

#"Outras Colunas Não Dinâmicas" = Table.UnpivotOtherColumns(Fonte, {"COMBUSTÍVEL ", " ANO ", " ESTADO "}, "Atributo", "Valor"),

#"Colunas Renomeadas" = Table.RenameColumns(#"Outras Colunas Não Dinâmicas",{{"Atributo", "DADOS"}}),

#"Somente as Colunas Selecionadas Foram Transformadas em Linhas" = Table.Unpivot(#"Colunas Renomeadas", {" ANO "}, "Atributo", "Valor.1"),

#"Colunas Renomeadas1" = Table.RenameColumns(#"Somente as Colunas Selecionadas Foram Transformadas em Linhas",{{"Valor.1", "ANO"}}),

#"Valor Substituído" = Table.ReplaceValue(#"Colunas Renomeadas1"," TOTAL ","TOTAL do ano",Replacer.ReplaceText,{"DADOS"})

in

#"Valor Substituído"


