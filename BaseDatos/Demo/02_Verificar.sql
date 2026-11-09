SET NOCOUNT ON;
IF (SELECT COUNT(*) FROM dbo.MDBDocument)<>29 OR (SELECT COUNT(*) FROM dbo.MDBServices)<>9 OR (SELECT COUNT(*) FROM dbo.MdbUsers)<>6
 BEGIN RAISERROR('La demostracion no tiene las cantidades esperadas.',16,1); RETURN; END;
IF EXISTS(SELECT 1 FROM dbo.MDBDocument WHERE FecEmi<'20110501' OR FecEmi>='20110601') OR
 EXISTS(SELECT 1 FROM dbo.MDBHistory WHERE FecEmi<'20110501' OR FecEmi>='20110601') OR
 EXISTS(SELECT 1 FROM dbo.MDBCostumer WHERE FecApe<'20110501' OR FecApe>='20110601')
 BEGIN RAISERROR('Los datos demo deben corresponder a mayo de 2011.',16,1); RETURN; END;
IF NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE Usuari='ADMIN' AND [Access]=0 AND SitUsu=1 AND ActPas=1)
 BEGIN RAISERROR('Falta el administrador demo.',16,1); RETURN; END;
IF EXISTS(SELECT NroCod FROM dbo.MDBHistory WHERE Status=1 AND Active=1 GROUP BY NroCod HAVING COUNT(*)>1)
 BEGIN RAISERROR('La demo contiene asignaciones vigentes repetidas.',16,1); RETURN; END;
IF EXISTS(SELECT Codigo FROM dbo.ViewControlDelivery GROUP BY Codigo HAVING COUNT(*)>1)
 BEGIN RAISERROR('Delivery duplica una guia.',16,1); RETURN; END;
IF NOT EXISTS(SELECT 1 FROM dbo.IntSaveActualizar WHERE VersionEsquema=2 AND ActSav=1)
 BEGIN RAISERROR('Falta el ejemplo de actualizaciones pendientes.',16,1); RETURN; END;
DBCC CHECKCONSTRAINTS WITH ALL_CONSTRAINTS;
SELECT Usuari,CASE [Access] WHEN 0 THEN 'Administrador' ELSE 'Operador' END Perfil,SitUsu Activo FROM dbo.MdbUsers ORDER BY ITEM;
PRINT 'OK: fechas, cuentas, guias, servicios, reparto y cola de demostracion.';
GO
