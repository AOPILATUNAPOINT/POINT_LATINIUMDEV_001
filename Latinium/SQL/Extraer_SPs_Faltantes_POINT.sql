-- =============================================================================
-- Ejecutar en POINT para extraer los 7 SPs faltantes
-- Copia el resultado (pestaña Results) y envialo para completar el script
-- =============================================================================

SELECT
    'IF OBJECT_ID(''[dbo].[' + p.name + ']'') IS NOT NULL' + CHAR(13)+CHAR(10) +
    '    DROP PROCEDURE [dbo].[' + p.name + ']' + CHAR(13)+CHAR(10) +
    'GO' + CHAR(13)+CHAR(10) +
    CHAR(13)+CHAR(10) +
    OBJECT_DEFINITION(p.object_id) + CHAR(13)+CHAR(10) +
    'GO' + CHAR(13)+CHAR(10) AS ScriptCompleto
FROM sys.procedures p
WHERE p.name IN (
    'Com_ConsultaPresupuestoCarpa',
    'Com_ConsultaPresupuestoAccesorios',
    'Com_ConsultaMetaConsultaCedulas',
    'Com_ConsultaPresupuestoMotos',
    'Com_ActualizaValoresConsultaCedulas',
    'Com_CopiaMesAnterior',
    'Com_AgregaLocalNuevo'
)
ORDER BY p.name
