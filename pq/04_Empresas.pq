// =============================================================================
// 04_Empresas
// -----------------------------------------------------------------------------
// Última empresa informada por cada persona en todo el período (usa TODAS
// las fichadas, válidas o no: el dato de empresa viaja con la credencial
// más allá de si el acceso fue otorgado). Si la persona nunca tiene
// empresa, queda "SIN EMPRESA". "--- A DEFINIR ---" se respeta tal cual
// viene. Equivalente a calcular_empresa_por_persona() en referencia.py.
// =============================================================================
let
    Origen = #"03_Limpieza",
    Agrupado = Table.Group(Origen, {"IdPersona"}, {
        {"Empresa", each
            let
                ordenado = Table.Sort(_, {{"FechaHora", Order.Ascending}}),
                valor = ordenado{Table.RowCount(ordenado) - 1}[EMPRESA]
            in
                if valor = null or Text.Trim(valor) = "" then "SIN EMPRESA" else valor,
            type text
        }
    })
in
    Agrupado
