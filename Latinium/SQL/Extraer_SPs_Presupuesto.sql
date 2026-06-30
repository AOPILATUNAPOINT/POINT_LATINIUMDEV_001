-- ============================================================
--  Ejecutar en SSMS para obtener las definiciones de los SPs
--  de Presupuesto y guardarlas en el archivo
--  METAS_JULIO_2026_CARGA.sql
-- ============================================================

SELECT
    ROUTINE_NAME,
    ROUTINE_DEFINITION
FROM INFORMATION_SCHEMA.ROUTINES
WHERE ROUTINE_TYPE = 'PROCEDURE'
  AND ROUTINE_NAME IN (
        'Com_ActualizaValoresPresupuestoEmpresa',
        'Com_ActualizaValoresPresupuestoEmpresaCredito',
        'Com_ActualizaValoresPresupuestoAccesorios',
        'Com_ActualizaValoresPresupuestoMotos',
        'Com_ActualizaValoresPresupuestoEmpresaGE',
        'Com_ActualizaValoresPresupuestoEmpresaContado',
        'Com_ActualizaValoresPresupuestoEmpresaCarpa',
        'Com_ConsultaPresupuestoEmpresa',
        'Com_ConsultaPresupuestoEmpresaCredito',
        'Com_ConsultaPresupuestoEmpresaGE',
        'Com_ConsultaPresupuestoEmpresaContado',
        'Com_GeneraPresupuestoEmpresa',
        'Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior'
    )
ORDER BY ROUTINE_NAME;

-- ============================================================
--  O usar sp_helptext para ver SPs mas largos de 4000 chars:
-- ============================================================
/*
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoEmpresa'
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoEmpresaCredito'
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoAccesorios'
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoMotos'
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoEmpresaGE'
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoEmpresaContado'
EXEC sp_helptext 'Com_ActualizaValoresPresupuestoEmpresaCarpa'
EXEC sp_helptext 'Com_ConsultaPresupuestoEmpresa'
EXEC sp_helptext 'Com_ConsultaPresupuestoEmpresaCredito'
EXEC sp_helptext 'Com_ConsultaPresupuestoEmpresaGE'
EXEC sp_helptext 'Com_ConsultaPresupuestoEmpresaContado'
EXEC sp_helptext 'Com_GeneraPresupuestoEmpresa'
EXEC sp_helptext 'Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior'
*/
