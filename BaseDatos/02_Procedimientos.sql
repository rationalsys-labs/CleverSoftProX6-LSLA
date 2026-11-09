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
IF OBJECT_ID(N'dbo.ObtenerCorrelativo',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.ObtenerCorrelativo AS RETURN 0;');
GO
ALTER PROCEDURE dbo.ObtenerCorrelativo @Divisor nvarchar(3),@Numero numeric(18,0) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    IF @@TRANCOUNT=0 BEGIN RAISERROR('El correlativo debe reservarse dentro de una transaccion.',16,1); RETURN; END;
    UPDATE dbo.IntCorrel WITH(UPDLOCK,HOLDLOCK) SET Correl=Correl+1 WHERE DivCor=@Divisor;
    IF @@ROWCOUNT<>1 BEGIN RAISERROR('No existe el divisor del correlativo.',16,1); RETURN; END;
    SELECT @Numero=Correl FROM dbo.IntCorrel WHERE DivCor=@Divisor;
    DECLARE @Item int=(SELECT Item FROM dbo.IntCorrel WHERE DivCor=@Divisor);
    EXEC dbo.RegistrarActualizacion N'IntCorrel',@Item;
END;
GO
-- Consulta previa para las pantallas originales. La reserva real ocurre al guardar.
IF OBJECT_ID(N'dbo.Correlativo',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.Correlativo AS RETURN 0;');
GO
ALTER PROCEDURE dbo.Correlativo @Divisor varchar(3)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Numero numeric(18,0)=(SELECT Correl+1 FROM dbo.IntCorrel WHERE DivCor=@Divisor);
    IF @Numero IS NULL BEGIN RAISERROR('No existe el divisor del correlativo.',16,1); RETURN; END;
    IF @Divisor='CLI'
        SELECT CONVERT(varchar(2),MONTH(GETDATE()))+CONVERT(varchar(2),DAY(GETDATE()))+RIGHT('000000'+CONVERT(varchar(20),@Numero),6) Correl;
    ELSE SELECT CONVERT(varchar(20),@Numero) Correl;
END;
GO
IF OBJECT_ID(N'dbo.Valid_LoginClev',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.Valid_LoginClev AS RETURN 0;');
GO
ALTER PROCEDURE dbo.Valid_LoginClev @Usuario varchar(5),@Pasword varchar(150)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT 'MenSQL' Modal,CASE WHEN EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE Usuari=@Usuario AND Paswor=@Pasword) THEN 1 ELSE 0 END Mensaje;
END;
GO
IF OBJECT_ID(N'dbo.Valid_LoginPrim',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.Valid_LoginPrim AS RETURN 0;');
GO
ALTER PROCEDURE dbo.Valid_LoginPrim @Usuario varchar(5),@Pasword varchar(150)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT 'MenSQL' Modal,CASE WHEN EXISTS(SELECT 1 FROM dbo.ADMUsers WHERE Usuari=@Usuario AND Paswor=@Pasword) THEN 1 ELSE 0 END Mensaje;
END;
GO
IF OBJECT_ID(N'dbo.SaveNewModClient',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveNewModClient AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveNewModClient
    @Codigo numeric(18,0),@Cliente nvarchar(300),@Tipo int,@RucDni nvarchar(20),
    @Direccion nvarchar(300),@Distrito int,@Representante nvarchar(150),@MailRepre nvarchar(150),
    @Contacto nvarchar(150),@Telefono nvarchar(50),@Teltrabajo nvarchar(50),@TelMovil nvarchar(50),
    @TelFax nvarchar(50),@EmailCont nvarchar(150),@Activate bit,@Modal tinyint
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    IF @Modal NOT IN(1,2) BEGIN RAISERROR('Modalidad de cliente no valida (1 nuevo, 2 modificar).',16,1); RETURN; END;
    BEGIN TRY
        BEGIN TRANSACTION;
        IF @Modal=1 BEGIN
            DECLARE @Numero numeric(18,0);
            EXEC dbo.ObtenerCorrelativo N'CLI',@Numero OUTPUT;
            IF @Numero>999999 BEGIN RAISERROR('El correlativo de clientes excedio su formato.',16,1); RETURN; END;
            SET @Codigo=CONVERT(numeric(18,0),CONVERT(varchar(2),MONTH(GETDATE()))+CONVERT(varchar(2),DAY(GETDATE()))+RIGHT('000000'+CONVERT(varchar(20),@Numero),6));
            INSERT dbo.MDBCostumer(CodCli,NomCli,TipDoc,RucDni,FecApe,DirCli,DisCli,RepLeg,MaiRep,ConCli,TelCli,TelTra,TelMov,TelFax,Email,ActCli)
            VALUES(@Codigo,@Cliente,@Tipo,@RucDni,GETDATE(),@Direccion,@Distrito,@Representante,@MailRepre,@Contacto,@Telefono,@Teltrabajo,@TelMovil,@TelFax,@EmailCont,@Activate);
        END ELSE BEGIN
            UPDATE dbo.MDBCostumer SET NomCli=@Cliente,TipDoc=@Tipo,RucDni=@RucDni,DirCli=@Direccion,DisCli=@Distrito,
                RepLeg=@Representante,MaiRep=@MailRepre,ConCli=@Contacto,TelCli=@Telefono,TelTra=@Teltrabajo,TelMov=@TelMovil,TelFax=@TelFax,Email=@EmailCont,ActCli=@Activate
            WHERE CodCli=@Codigo;
            IF @@ROWCOUNT<>1 BEGIN RAISERROR('No existe el cliente.',16,1); RETURN; END;
        END;
        EXEC dbo.RegistrarActualizacion N'MDBCostumer',@Codigo;
        COMMIT;
        SELECT @Codigo Codigo;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK;
        BEGIN DECLARE @MensajeError1 nvarchar(4000); SET @MensajeError1=ERROR_MESSAGE(); RAISERROR(@MensajeError1,16,1); RETURN; END;
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.SaveNewService',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveNewService AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveNewService
    @Cliente numeric(18,0),@TipoServicio int,@Cantidad numeric(18,0),@Producto int,@Boleta numeric(18,0)=NULL
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        IF @Cantidad<=0 BEGIN RAISERROR('La cantidad del servicio debe ser mayor que cero.',16,1); RETURN; END;
        IF NOT EXISTS(SELECT 1 FROM dbo.MDBCostumer WHERE CodCli=@Cliente AND Active=1 AND ActCli=1)
            BEGIN RAISERROR('El cliente no esta disponible para nuevos servicios.',16,1); RETURN; END;
        DECLARE @Codigo numeric(18,0);
        EXEC dbo.ObtenerCorrelativo N'CSV',@Codigo OUTPUT;
        INSERT dbo.MDBServices(CodSer,NroBol,CodCLi,TipSer,CanSer,Tipo)
        VALUES(@Codigo,ISNULL(@Boleta,@Codigo),@Cliente,@TipoServicio,@Cantidad,@Producto);
        EXEC dbo.RegistrarActualizacion N'MDBServices',@Codigo;
        COMMIT;
        SELECT @Codigo Codigo;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK;
        BEGIN DECLARE @MensajeError2 nvarchar(4000); SET @MensajeError2=ERROR_MESSAGE(); RAISERROR(@MensajeError2,16,1); RETURN; END;
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.SaveNewHisClient',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveNewHisClient AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveNewHisClient
    @Codigo numeric(18,0),@Nombre nvarchar(500),@Direccion nvarchar(500),@Distrito int,@Contacto nvarchar(150),
    @CodigoNuevo numeric(18,0)=NULL OUTPUT,@DevolverCodigo bit=1
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        -- Una direccion ya utilizada no se sobrescribe: se busca o crea la nueva version.
        SELECT @CodigoNuevo=MIN(CodEmp) FROM dbo.MDBHisEmp WITH(UPDLOCK,HOLDLOCK)
        WHERE NomEmp=@Nombre AND DirEmp=@Direccion AND DisEmp=@Distrito AND ConEmp=@Contacto;
        IF @CodigoNuevo IS NULL BEGIN
            EXEC dbo.ObtenerCorrelativo N'CHC',@CodigoNuevo OUTPUT;
            INSERT dbo.MDBHisEmp(CodEmp,NomEmp,DirEmp,DisEmp,ConEmp) VALUES(@CodigoNuevo,@Nombre,@Direccion,@Distrito,@Contacto);
            EXEC dbo.RegistrarActualizacion N'MDBHisEmp',@CodigoNuevo;
        END;
        COMMIT;
        IF @DevolverCodigo=1 SELECT @CodigoNuevo Codigo;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK;
        BEGIN DECLARE @MensajeError3 nvarchar(4000); SET @MensajeError3=ERROR_MESSAGE(); RAISERROR(@MensajeError3,16,1); RETURN; END;
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.GuardarDocumento',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.GuardarDocumento AS RETURN 0;');
GO
ALTER PROCEDURE dbo.GuardarDocumento
    @Servicio numeric(18,0),@Documento nvarchar(50),@Producto int,@Cliente numeric(18,0),
    @Destinatario nvarchar(500),@Direccion nvarchar(500),@Distrito int,@Contacto nvarchar(150),
    @Descripcion nvarchar(150),@CodSave numeric(18,0),@CodigoNuevo numeric(18,0) OUTPUT
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        DECLARE @CantidadServicio numeric(18,0),@EstadoServicio bit;
        SELECT @CantidadServicio=CanSer,@EstadoServicio=Estado FROM dbo.MDBServices WITH(UPDLOCK,HOLDLOCK) WHERE CodSer=@Servicio AND CodCLi=@Cliente;
        IF @CantidadServicio IS NULL
            BEGIN RAISERROR('El cliente no corresponde al servicio indicado.',16,1); RETURN; END;
        IF @EstadoServicio=0 BEGIN RAISERROR('El servicio esta finalizado. Debe reactivarlo antes de registrar documentos.',16,1); RETURN; END;
        IF @Servicio<>105 AND (SELECT COUNT_BIG(*) FROM dbo.MDBDocument WHERE CoServ=@Servicio AND Active=1)>=@CantidadServicio
            BEGIN RAISERROR('Se alcanzo la cantidad de documentos del servicio.',16,1); RETURN; END;
        IF LEN(LTRIM(RTRIM(ISNULL(@Destinatario,N''))))=0 OR LEN(LTRIM(RTRIM(ISNULL(@Direccion,N''))))=0
            BEGIN RAISERROR('El destinatario y su direccion son obligatorios.',16,1); RETURN; END;
        DECLARE @Empresa numeric(18,0);
        EXEC dbo.SaveNewHisClient @CodSave,@Destinatario,@Direccion,@Distrito,@Contacto,@Empresa OUTPUT,0;
        EXEC dbo.ObtenerCorrelativo N'CDO',@CodigoNuevo OUTPUT;
        INSERT dbo.MDBDocument(NroCod,CoServ,Docmto,CoProd,CodEmp,Descrip)
        VALUES(@CodigoNuevo,@Servicio,ISNULL(@Documento,N''),@Producto,@Empresa,ISNULL(@Descripcion,N''));
        EXEC dbo.RegistrarActualizacion N'MDBDocument',@CodigoNuevo;
        COMMIT;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK;
        BEGIN DECLARE @MensajeError4 nvarchar(4000); SET @MensajeError4=ERROR_MESSAGE(); RAISERROR(@MensajeError4,16,1); RETURN; END;
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.SaveNewRegister',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveNewRegister AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveNewRegister
    @Servicio numeric(18,0),@Documento nvarchar(50),@Producto int,@Cliente numeric(18,0),
    @Destinatario nvarchar(500),@Direccion nvarchar(500),@Distrito int,@Contacto nvarchar(150),
    @Descrip nvarchar(150)=N'',@CodSave numeric(18,0)=0
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Codigo numeric(18,0);
    EXEC dbo.GuardarDocumento @Servicio,@Documento,@Producto,@Cliente,@Destinatario,@Direccion,@Distrito,@Contacto,@Descrip,@CodSave,@Codigo OUTPUT;
    SELECT @Codigo Guia;
END;
GO
IF OBJECT_ID(N'dbo.SaveNewRegisterExcel',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveNewRegisterExcel AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveNewRegisterExcel
    @Servicio numeric(18,0),@Documento nvarchar(50),@Producto int,@Cliente numeric(18,0),
    @Destinatario nvarchar(500),@Direccion nvarchar(500),@Distrito int,@Contacto nvarchar(150),@Descripcion nvarchar(150)=N''
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Codigo numeric(18,0);
    EXEC dbo.GuardarDocumento @Servicio,@Documento,@Producto,@Cliente,@Destinatario,@Direccion,@Distrito,@Contacto,@Descripcion,0,@Codigo OUTPUT;
    SELECT @Codigo Info;
END;
GO
IF OBJECT_ID(N'dbo.SaveStatusConfirm',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveStatusConfirm AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveStatusConfirm
    @Codigo numeric(18,0),@Usuario int,@Status int,@Fecha datetime,@Informe nvarchar(500),@CodMen int=NULL,@Ranche numeric(18,0)=0
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        IF NOT EXISTS(SELECT 1 FROM dbo.MDBDocument WITH(UPDLOCK,HOLDLOCK) WHERE NroCod=@Codigo AND Active=1)
            RAISERROR('El documento no existe o fue anulado.',16,1);
        IF NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WITH(HOLDLOCK) WHERE ITEM=@Usuario AND SitUsu=1)
            RAISERROR('El usuario operativo no esta disponible.',16,1);
        IF NOT EXISTS(SELECT 1 FROM dbo.IntStatus WITH(HOLDLOCK) WHERE ITEM=@Status AND Active=1)
            RAISERROR('El estado no esta disponible.',16,1);
        IF @CodMen IS NOT NULL AND NOT EXISTS(SELECT 1 FROM dbo.Mdbmensajeros WITH(HOLDLOCK) WHERE CodMen=@CodMen AND Active=1)
            RAISERROR('El mensajero no esta disponible.',16,1);
        IF @Fecha IS NULL OR @Ranche IS NULL OR @Ranche<0 RAISERROR('La fecha o el cargo no son validos.',16,1);
        INSERT dbo.MDBHistory(NroCod,Status,FecEmi,Usuari,Ranche,Informe,CodMen)
        VALUES(@Codigo,@Status,@Fecha,@Usuario,@Ranche,ISNULL(@Informe,N''),@CodMen);
        DECLARE @Item numeric(18,0)=SCOPE_IDENTITY();
        EXEC dbo.RegistrarActualizacion N'MDBHistory',@Item;
        COMMIT;
        SELECT @Item Item;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK;
        BEGIN DECLARE @MensajeError5 nvarchar(4000); SET @MensajeError5=ERROR_MESSAGE(); RAISERROR(@MensajeError5,16,1); RETURN; END;
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.SaveHistoryStatus',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveHistoryStatus AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveHistoryStatus @Codigo numeric(18,0),@Usuario int,@ranche numeric(18,0),@mensajero int
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Fecha datetime=GETDATE();
    EXEC dbo.SaveStatusConfirm @Codigo,@Usuario,1,@Fecha,N'',@mensajero,@ranche;
END;
GO
IF OBJECT_ID(N'dbo.SaveHistoryStatusFin',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveHistoryStatusFin AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveHistoryStatusFin @Codigo numeric(18,0),@Usuario int,@ranche numeric(18,0),@Status int
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Fecha datetime=GETDATE();
    EXEC dbo.SaveStatusConfirm @Codigo,@Usuario,@Status,@Fecha,N'',NULL,@ranche;
END;
GO
IF OBJECT_ID(N'dbo.SaveHistoryConfirmSta',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.SaveHistoryConfirmSta AS RETURN 0;');
GO
ALTER PROCEDURE dbo.SaveHistoryConfirmSta @Codigo numeric(18,0),@Status int,@Usuari int,@Inform nvarchar(500)
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @Fecha datetime=GETDATE();
    EXEC dbo.SaveStatusConfirm @Codigo,@Usuari,@Status,@Fecha,@Inform;
END;
GO
-- Los documentos conservan la direccion con la que se registraron.
-- Las correcciones del catalogo crean otra fila mediante SaveNewHisClient.
IF OBJECT_ID(N'dbo.TR_MDBHisEmp_DireccionUtilizada',N'TR') IS NOT NULL DROP TRIGGER dbo.TR_MDBHisEmp_DireccionUtilizada;
GO
CREATE TRIGGER dbo.TR_MDBHisEmp_DireccionUtilizada ON dbo.MDBHisEmp AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS(SELECT 1 FROM inserted i JOIN deleted d ON d.CodEmp=i.CodEmp
        WHERE EXISTS(SELECT 1 FROM dbo.MDBDocument doc WHERE doc.CodEmp=i.CodEmp)
        AND (i.NomEmp<>d.NomEmp OR i.DirEmp<>d.DirEmp OR i.DisEmp<>d.DisEmp OR i.ConEmp<>d.ConEmp))
        BEGIN IF XACT_STATE()<>0 ROLLBACK; RAISERROR('La direccion ya fue utilizada. Cree otra version para conservar los documentos anteriores.',16,1); RETURN; END;
END;
GO
