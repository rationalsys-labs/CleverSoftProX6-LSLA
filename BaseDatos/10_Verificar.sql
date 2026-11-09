SET NOCOUNT ON;
IF (SELECT COUNT(*) FROM sys.tables WHERE is_ms_shipped=0)<>21
 BEGIN RAISERROR('La base debe contener las 21 tablas de Clever SoftPro.',16,1); RETURN; END;
IF EXISTS(SELECT 1 FROM sys.tables t WHERE t.is_ms_shipped=0 AND NOT EXISTS(SELECT 1 FROM sys.indexes i WHERE i.object_id=t.object_id AND i.is_primary_key=1))
 BEGIN RAISERROR('Hay tablas sin clave primaria.',16,1); RETURN; END;
IF (SELECT COUNT(*) FROM sys.foreign_keys)<16 OR EXISTS(SELECT 1 FROM sys.foreign_keys WHERE is_disabled=1 OR is_not_trusted=1)
 BEGIN RAISERROR('Las relaciones no estan completas o verificadas.',16,1); RETURN; END;
IF EXISTS(SELECT 1 FROM sys.check_constraints WHERE is_disabled=1 OR is_not_trusted=1)
 BEGIN RAISERROR('Hay validaciones deshabilitadas o no verificadas.',16,1); RETURN; END;
IF OBJECT_ID('dbo.GuardarUsuario','P') IS NULL OR OBJECT_ID('dbo.RegistrarReparto','P') IS NULL OR OBJECT_ID('dbo.GuardarConfirmaciones','P') IS NULL OR OBJECT_ID('dbo.AplicarActualizacion','P') IS NULL
 BEGIN RAISERROR('Faltan procedimientos del sistema.',16,1); RETURN; END;
DECLARE @Vista sysname,@Sql nvarchar(max);
DECLARE Vistas CURSOR LOCAL FAST_FORWARD FOR SELECT name FROM sys.views WHERE is_ms_shipped=0;
OPEN Vistas; FETCH NEXT FROM Vistas INTO @Vista;
WHILE @@FETCH_STATUS=0 BEGIN
 SET @Sql=N'SELECT TOP(0) * FROM dbo.'+QUOTENAME(@Vista)+N';';
 EXEC sp_executesql @Sql;
 FETCH NEXT FROM Vistas INTO @Vista;
END;
CLOSE Vistas; DEALLOCATE Vistas;
DBCC CHECKCONSTRAINTS WITH ALL_CONSTRAINTS;
PRINT 'OK: tablas, claves, relaciones, restricciones, procedimientos y vistas.';
GO
