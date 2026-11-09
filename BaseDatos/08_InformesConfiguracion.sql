SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF OBJECT_ID(N'dbo.AplicarActualizacion',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.AplicarActualizacion AS RETURN 0;');
GO
ALTER PROCEDURE dbo.AplicarActualizacion @Tabla sysname,@Datos nvarchar(max)
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    DECLARE @Clave sysname;
    SET @Clave=CASE @Tabla WHEN N'MDBCostumer' THEN N'CodCli' WHEN N'MDBServices' THEN N'CodSer'
        WHEN N'MDBDocument' THEN N'NroCod' WHEN N'MDBHistory' THEN N'ITEM'
        WHEN N'MDBHisEmp' THEN N'CodEmp' WHEN N'IntCorrel' THEN N'Item'
        WHEN N'MdbGruposServ' THEN N'Item' WHEN N'IntStatus' THEN N'ITEM' END;
    IF @Clave IS NULL RAISERROR('Tabla de actualizacion no valida.',16,1);
    IF @Clave IS NULL RETURN;
    DECLARE @Xml xml;
    SET @Xml=CONVERT(xml,@Datos);
    IF @Xml.value('count(/Fila)','int')<>1 RAISERROR('La actualizacion debe contener una fila.',16,1);
    IF @Xml.value('count(/Fila)','int')<>1 RETURN;
    IF @Tabla=N'IntStatus' AND @Xml.value('(/Fila/ITEM)[1]','int') IN(1,2) AND @Xml.value('(/Fila/Active)[1]','bit')<>1
        BEGIN RAISERROR('Los estados basicos deben permanecer activos.',16,1); RETURN; END;
    DECLARE @Objeto int,@Nombre sysname,@Tipo nvarchar(100),@Expresion nvarchar(max);
    DECLARE @Columnas nvarchar(max),@Seleccion nvarchar(max),@Cambios nvarchar(max),@Sql nvarchar(max),@Identidad bit;
    SET @Objeto=OBJECT_ID(N'dbo.'+QUOTENAME(@Tabla));
    SET @Columnas=N''; SET @Seleccion=N''; SET @Cambios=N'';
    SET @Identidad=CASE WHEN EXISTS(SELECT 1 FROM sys.identity_columns WHERE object_id=@Objeto) THEN 1 ELSE 0 END;
    DECLARE Columnas CURSOR LOCAL FAST_FORWARD FOR
        SELECT c.name,t.name+CASE WHEN t.name IN(N'nvarchar',N'nchar') THEN N'('+CASE WHEN c.max_length=-1 THEN N'max' ELSE CONVERT(nvarchar(10),c.max_length/2) END+N')'
            WHEN t.name IN(N'numeric',N'decimal') THEN N'('+CONVERT(nvarchar(10),c.precision)+N','+CONVERT(nvarchar(10),c.scale)+N')' ELSE N'' END
        FROM sys.columns c JOIN sys.types t ON t.user_type_id=c.user_type_id WHERE c.object_id=@Objeto ORDER BY c.column_id;
    OPEN Columnas; FETCH NEXT FROM Columnas INTO @Nombre,@Tipo;
    WHILE @@FETCH_STATUS=0 BEGIN
        SET @Expresion=N'CASE WHEN @Xml.exist(''/Fila/'+@Nombre+N''')=0 THEN NULL ELSE CONVERT('+@Tipo+N',@Xml.value(''string((/Fila/'+@Nombre+N')[1])'',''nvarchar(max)''),126) END';
        SET @Columnas=@Columnas+CASE WHEN @Columnas=N'' THEN N'' ELSE N',' END+QUOTENAME(@Nombre);
        SET @Seleccion=@Seleccion+CASE WHEN @Seleccion=N'' THEN N'' ELSE N',' END+@Expresion+N' AS '+QUOTENAME(@Nombre);
        IF @Nombre<>@Clave AND NOT EXISTS(SELECT 1 FROM sys.identity_columns WHERE object_id=@Objeto AND name=@Nombre)
            SET @Cambios=@Cambios+CASE WHEN @Cambios=N'' THEN N'' ELSE N',' END+QUOTENAME(@Nombre)+N'='+
                CASE WHEN @Tabla=N'IntCorrel' AND @Nombre=N'Correl' THEN N'CASE WHEN d.Correl>o.Correl THEN d.Correl ELSE o.Correl END' ELSE N'o.'+QUOTENAME(@Nombre) END;
        FETCH NEXT FROM Columnas INTO @Nombre,@Tipo;
    END;
    CLOSE Columnas; DEALLOCATE Columnas;
    SET @Sql=N'UPDATE d WITH(UPDLOCK,HOLDLOCK) SET '+@Cambios+N' FROM dbo.'+QUOTENAME(@Tabla)+N' d CROSS JOIN(SELECT '+@Seleccion+N') o WHERE d.'+QUOTENAME(@Clave)+N'=o.'+QUOTENAME(@Clave)+N';
        IF @@ROWCOUNT=0 BEGIN '+CASE WHEN @Identidad=1 THEN N'SET IDENTITY_INSERT dbo.'+QUOTENAME(@Tabla)+N' ON;' ELSE N'' END+
        N'INSERT dbo.'+QUOTENAME(@Tabla)+N'('+@Columnas+N') SELECT '+@Columnas+N' FROM(SELECT '+@Seleccion+N')o;'+
        CASE WHEN @Identidad=1 THEN N'SET IDENTITY_INSERT dbo.'+QUOTENAME(@Tabla)+N' OFF;' ELSE N'' END+N' END;';
    BEGIN TRY
        BEGIN TRANSACTION;
        EXEC sp_executesql @Sql,N'@Xml xml',@Xml;
        COMMIT;
    END TRY
    BEGIN CATCH
        DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
        IF XACT_STATE()<>0 ROLLBACK;
        IF @Identidad=1 BEGIN
            SET @Sql=N'SET IDENTITY_INSERT dbo.'+QUOTENAME(@Tabla)+N' OFF;'; EXEC sp_executesql @Sql;
        END;
        RAISERROR(@Error,16,1);
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.RegistrarActualizacion',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.RegistrarActualizacion AS RETURN 0;');
GO
ALTER PROCEDURE dbo.RegistrarActualizacion @Tabla sysname,@Codigo numeric(18,0)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Clave sysname,@Datos nvarchar(max),@Sql nvarchar(max);
    SET @Clave=CASE @Tabla WHEN N'MDBCostumer' THEN N'CodCli' WHEN N'MDBServices' THEN N'CodSer'
        WHEN N'MDBDocument' THEN N'NroCod' WHEN N'MDBHistory' THEN N'ITEM'
        WHEN N'MDBHisEmp' THEN N'CodEmp' WHEN N'IntCorrel' THEN N'Item'
        WHEN N'MdbGruposServ' THEN N'Item' WHEN N'IntStatus' THEN N'ITEM' END;
    IF @Clave IS NULL RAISERROR('Tabla de actualizacion no admitida.',16,1);
    IF @Clave IS NULL RETURN;
    SET @Sql=N'SELECT @Datos=CONVERT(nvarchar(max),(SELECT * FROM dbo.'+QUOTENAME(@Tabla)+N' WHERE '+QUOTENAME(@Clave)+N'=@Codigo FOR XML PATH(''Fila''),TYPE));';
    EXEC sp_executesql @Sql,N'@Codigo numeric(18,0),@Datos nvarchar(max) OUTPUT',@Codigo,@Datos OUTPUT;
    IF ISNULL(@Datos,N'')=N'' RAISERROR('No existe la fila que se quiere registrar.',16,1);
    IF ISNULL(@Datos,N'')=N'' RETURN;
    INSERT dbo.IntSaveActualizar(Consulta,ActSav,VersionEsquema)
    VALUES(N'EXEC dbo.AplicarActualizacion @Tabla=N'''+@Tabla+N''',@Datos=N'''+REPLACE(@Datos,N'''',N'''''')+N''';',1,2);
END;
GO
IF OBJECT_ID(N'dbo.GuardarEstado',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.GuardarEstado AS RETURN 0;');
GO
ALTER PROCEDURE dbo.GuardarEstado @Codigo int,@Accion tinyint,@Nombre nvarchar(4000)=N''
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  IF @Accion IS NULL OR @Accion NOT IN(1,2,3) RAISERROR('Accion de estado no valida.',16,1);
  IF @Accion IN(1,2) BEGIN
   SET @Nombre=LTRIM(RTRIM(@Nombre));
   IF @Nombre IS NULL OR LEN(@Nombre)=0 OR DATALENGTH(@Nombre)>40 RAISERROR('El estado admite entre 1 y 20 caracteres.',16,1);
   IF EXISTS(SELECT 1 FROM dbo.IntStatus WITH(UPDLOCK,HOLDLOCK) WHERE Status=@Nombre AND (@Accion=1 OR ITEM<>@Codigo))
    RAISERROR('Ya existe un estado con ese nombre.',16,1);
  END;
  IF @Accion=1 BEGIN
   INSERT dbo.IntStatus(Status,ActEst,Active) VALUES(@Nombre,1,1);
   SET @Codigo=CONVERT(int,SCOPE_IDENTITY());
  END ELSE BEGIN
   IF NOT EXISTS(SELECT 1 FROM dbo.IntStatus WITH(UPDLOCK,HOLDLOCK) WHERE ITEM=@Codigo AND ActEst=1 AND Active=1)
    RAISERROR('El estado no existe, esta de baja o no es editable.',16,1);
   IF @Accion=3 AND @Codigo IN(1,2) RAISERROR('Reparto y Entregado deben permanecer activos.',16,1);
   IF @Accion=2 UPDATE dbo.IntStatus SET Status=@Nombre WHERE ITEM=@Codigo;
   ELSE UPDATE dbo.IntStatus SET Active=0 WHERE ITEM=@Codigo;
  END;
  DECLARE @Clave numeric(18,0); SET @Clave=@Codigo;
  EXEC dbo.RegistrarActualizacion N'IntStatus',@Clave;
  COMMIT;
  SELECT @Codigo Codigo;
 END TRY
 BEGIN CATCH
  DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
  IF XACT_STATE()<>0 ROLLBACK;
  RAISERROR(@Error,16,1);
 END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.MarcarActualizaciones',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.MarcarActualizaciones AS RETURN 0;');
GO
ALTER PROCEDURE dbo.MarcarActualizaciones @Seleccion nvarchar(max)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  DECLARE @Xml xml; SET @Xml=CONVERT(xml,@Seleccion);
  DECLARE @Items TABLE(Codigo int PRIMARY KEY);
  INSERT @Items SELECT Fila.value('@Codigo','int') FROM @Xml.nodes('/Actualizaciones/Item') Lista(Fila);
  DECLARE @Cantidad int; SELECT @Cantidad=COUNT(*) FROM @Items;
  IF @Cantidad=0 RAISERROR('No se seleccionaron actualizaciones.',16,1);
  UPDATE a SET ActSav=0 FROM dbo.IntSaveActualizar a JOIN @Items i ON i.Codigo=a.Item WHERE a.VersionEsquema=2 AND a.ActSav IN(0,1);
  IF @@ROWCOUNT<>@Cantidad RAISERROR('La seleccion de actualizaciones ya no esta disponible.',16,1);
  COMMIT;
 END TRY
 BEGIN CATCH
  DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
  IF XACT_STATE()<>0 ROLLBACK;
  RAISERROR(@Error,16,1);
 END CATCH;
END;
GO