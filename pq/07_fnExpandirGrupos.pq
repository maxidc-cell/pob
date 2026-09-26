// =============================================================================
// 07_fnExpandirGrupos
// -----------------------------------------------------------------------------
// Función reutilizada por los reportes POB_Horario, POB_Franjas y Comedor.
// A partir de una tabla con columnas [...columnasGrupo, IdPersona, Segmento,
// Empresa], arma las filas de salida con Segmento en {valores..., TOTAL} x
// Empresa en {valores..., TODAS}, contando personas DISTINTAS (nunca
// sumando filas). Es el equivalente vectorizado de _combos_segmento_empresa
// + los nunique() de referencia.py.
//
// NOTA para quien audite el M contra Python: a diferencia de
// validacion/logica_pob.py (que en algunos lugares itera combo por combo
// con nunique()), acá se usa Table.Group una sola vez por variante para que
// la consulta escale bien con ~90.000 filas/mes. El resultado numérico es
// idéntico; sólo cambia la forma de calcularlo.
//
// Filas con conteo 0 no se generan (igual que en referencia.py): una
// combinación Fecha/Hora/Segmento/Empresa sin nadie presente simplemente no
// aparece en la tabla. Para una tabla dinámica esto no es un problema
// (Power Pivot no necesita la fila en cero para graficar bien un total).
// =============================================================================
let
    fnExpandirGrupos = (tabla as table, columnasGrupo as list, nombreValor as text) as table =>
        let
            contarDistintos = (t as table) as number => Table.RowCount(Table.Distinct(Table.SelectColumns(t, {"IdPersona"}))),

            PorSegmentoYEmpresa = Table.Group(tabla, columnasGrupo & {"Segmento", "Empresa"}, {{nombreValor, each contarDistintos(_), Int64.Type}}),

            PorSegmentoTodas = Table.Group(tabla, columnasGrupo & {"Segmento"}, {{nombreValor, each contarDistintos(_), Int64.Type}}),
            PorSegmentoTodas2 = Table.AddColumn(PorSegmentoTodas, "Empresa", each "TODAS", type text),

            TotalPorEmpresa = Table.Group(tabla, columnasGrupo & {"Empresa"}, {{nombreValor, each contarDistintos(_), Int64.Type}}),
            TotalPorEmpresa2 = Table.AddColumn(TotalPorEmpresa, "Segmento", each "TOTAL", type text),

            TotalTodas = Table.Group(tabla, columnasGrupo, {{nombreValor, each contarDistintos(_), Int64.Type}}),
            TotalTodas2 = Table.AddColumn(Table.AddColumn(TotalTodas, "Segmento", each "TOTAL", type text), "Empresa", each "TODAS", type text),

            OrdenColumnas = columnasGrupo & {"Segmento", "Empresa", nombreValor},
            Combinado = Table.Combine({
                Table.ReorderColumns(PorSegmentoYEmpresa, OrdenColumnas),
                Table.ReorderColumns(PorSegmentoTodas2, OrdenColumnas),
                Table.ReorderColumns(TotalPorEmpresa2, OrdenColumnas),
                Table.ReorderColumns(TotalTodas2, OrdenColumnas)
            })
        in
            Combinado
in
    fnExpandirGrupos
