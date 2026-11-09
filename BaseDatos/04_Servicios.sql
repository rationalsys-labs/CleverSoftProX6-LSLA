SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF OBJECT_ID(N'dbo.ActualizarServicio',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.ActualizarServicio AS RETURN 0;');
GO
ALTER PROCEDURE dbo.ActualizarServicio @Codigo numeric(18,0),@Accion tinyint,@Cantidad numeric(18,0)=NULL
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        DECLARE @Actual numeric(18,0),@Estado bit,@Registrados bigint;
        SELECT @Actual=CanSer,@Estado=Estado FROM dbo.MDBServices WITH(UPDLOCK,HOLDLOCK) WHERE CodSer=@Codigo;
        IF @Actual IS NULL BEGIN RAISERROR('No existe el servicio.',16,1); RETURN; END;
        SELECT @Registrados=COUNT_BIG(*) FROM dbo.MDBDocument WHERE CoServ=@Codigo AND Active=1;
        IF @Accion=1 BEGIN
            IF @Estado=0 OR @Cantidad IS NULL OR @Cantidad<=0 OR @Cantidad<@Registrados
                BEGIN RAISERROR('La cantidad no puede ser menor que los documentos registrados ni modificarse en un servicio finalizado.',16,1); RETURN; END;
            UPDATE dbo.MDBServices SET CanSer=@Cantidad WHERE CodSer=@Codigo;
        END ELSE IF @Accion=2 BEGIN
            IF @Codigo=105 OR @Actual<>@Registrados
                BEGIN RAISERROR('La cantidad registrada debe coincidir con el servicio para finalizarlo.',16,1); RETURN; END;
            UPDATE dbo.MDBServices SET Estado=0 WHERE CodSer=@Codigo;
        END ELSE IF @Accion=3 BEGIN
            UPDATE dbo.MDBServices SET Estado=1 WHERE CodSer=@Codigo;
        END ELSE BEGIN RAISERROR('Operacion de servicio no valida.',16,1); RETURN; END;
        EXEC dbo.RegistrarActualizacion N'MDBServices',@Codigo;
        COMMIT;
    END TRY
    BEGIN CATCH
        DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
        IF XACT_STATE()<>0 ROLLBACK;
        RAISERROR(@Error,16,1);
    END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.AnularDocumento',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.AnularDocumento AS RETURN 0;');
GO
ALTER PROCEDURE dbo.AnularDocumento @Codigo numeric(18,0)
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        DECLARE @Servicio numeric(18,0),@Estado bit;
        SELECT @Servicio=CoServ FROM dbo.MDBDocument WHERE NroCod=@Codigo AND Active=1;
        SELECT @Estado=Estado FROM dbo.MDBServices WITH(UPDLOCK,HOLDLOCK) WHERE CodSer=@Servicio;
        IF @Servicio IS NULL OR @Estado IS NULL OR @Estado=0
            BEGIN RAISERROR('El documento no existe o pertenece a un servicio finalizado.',16,1); RETURN; END;
        UPDATE dbo.MDBDocument SET Active=0 WHERE NroCod=@Codigo AND Active=1;
        IF @@ROWCOUNT<>1 BEGIN RAISERROR('El documento ya fue anulado.',16,1); RETURN; END;
        EXEC dbo.RegistrarActualizacion N'MDBDocument',@Codigo;
        COMMIT;
    END TRY
    BEGIN CATCH
        DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
        IF XACT_STATE()<>0 ROLLBACK;
        RAISERROR(@Error,16,1);
    END CATCH;
END;
GO
