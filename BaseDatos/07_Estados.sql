SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF OBJECT_ID(N'dbo.GuardarConfirmaciones',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.GuardarConfirmaciones AS RETURN 0;');
GO
ALTER PROCEDURE dbo.GuardarConfirmaciones @Usuario int,@Confirmaciones nvarchar(max)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  IF NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WITH(HOLDLOCK) WHERE ITEM=@Usuario AND SitUsu=1)
   RAISERROR('El usuario operativo no esta disponible.',16,1);
  DECLARE @Xml xml; SET @Xml=CONVERT(xml,@Confirmaciones);
  DECLARE @Lista TABLE(Codigo numeric(18,0) PRIMARY KEY,Estado int NOT NULL,Fecha datetime NOT NULL,Informe nvarchar(max) NOT NULL);
  INSERT @Lista
   SELECT Fila.value('@Codigo','numeric(18,0)'),Fila.value('@Estado','int'),Fila.value('@Fecha','datetime'),
    Fila.value('string((Informe)[1])','nvarchar(max)') FROM @Xml.nodes('/Confirmaciones/Confirmacion') Lista(Fila);
  IF NOT EXISTS(SELECT 1 FROM @Lista) OR EXISTS(SELECT 1 FROM @Lista WHERE Codigo<=0 OR Estado<=1 OR DATALENGTH(Informe)>1000)
   RAISERROR('Ingrese confirmaciones validas; la observacion admite hasta 500 caracteres.',16,1);
  IF EXISTS(SELECT 1 FROM @Lista c LEFT JOIN dbo.IntStatus s WITH(HOLDLOCK) ON s.ITEM=c.Estado WHERE s.ITEM IS NULL OR s.Active<>1)
   RAISERROR('Un estado ya no esta disponible.',16,1);
  DECLARE @Esperados int,@Actuales int;
  SELECT @Esperados=COUNT(*) FROM @Lista;
  SELECT @Actuales=COUNT(*) FROM dbo.MDBDocument d WITH(UPDLOCK,HOLDLOCK) JOIN @Lista c ON c.Codigo=d.NroCod WHERE d.Active=1;
  IF @Actuales<>@Esperados RAISERROR('Un documento no existe o fue anulado. Revise el lote.',16,1);
  DECLARE @Nuevos TABLE(Item int);
  INSERT dbo.MDBHistory(NroCod,Status,FecEmi,Usuari,Informe)
   OUTPUT inserted.ITEM INTO @Nuevos
   SELECT c.Codigo,c.Estado,c.Fecha,@Usuario,c.Informe FROM @Lista c
   WHERE NOT EXISTS(SELECT 1 FROM dbo.MDBHistory h WHERE h.NroCod=c.Codigo AND h.Status=c.Estado
    AND h.FecEmi=c.Fecha AND h.Usuari=@Usuario AND h.Informe=c.Informe AND h.Active=1);
  DECLARE @Item numeric(18,0);
  DECLARE Historial CURSOR LOCAL FAST_FORWARD FOR SELECT Item FROM @Nuevos ORDER BY Item;
  OPEN Historial; FETCH NEXT FROM Historial INTO @Item;
  WHILE @@FETCH_STATUS=0 BEGIN
   EXEC dbo.RegistrarActualizacion N'MDBHistory',@Item;
   FETCH NEXT FROM Historial INTO @Item;
  END;
  CLOSE Historial; DEALLOCATE Historial;
  COMMIT;
 END TRY
 BEGIN CATCH
  DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
  IF XACT_STATE()<>0 ROLLBACK;
  RAISERROR(@Error,16,1);
 END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.ConsultarDelivery',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.ConsultarDelivery AS RETURN 0;');
GO
ALTER PROCEDURE dbo.ConsultarDelivery @Mensajero int,@Desde datetime,@Hasta datetime
AS
BEGIN
 SET NOCOUNT ON;
 IF @Desde IS NULL OR @Hasta IS NULL OR CONVERT(date,@Desde)>CONVERT(date,@Hasta)
  BEGIN RAISERROR('El intervalo de fechas no es valido.',16,1); RETURN; END;
 SET @Desde=CONVERT(datetime,CONVERT(date,@Desde));
 SET @Hasta=CONVERT(datetime,CONVERT(date,@Hasta));
 DECLARE @Limite datetime;
 IF @Hasta<CONVERT(datetime,'99991231',112) SET @Limite=DATEADD(day,1,@Hasta);
 SELECT Codigo,Descrip,Empresa,Contacto,Emision,Informe,Status,Fecha,CodEstado
 FROM dbo.ViewControlDelivery WHERE CodMen=@Mensajero AND Emision>=@Desde
  AND (@Limite IS NULL OR Emision<@Limite)
 ORDER BY Emision,Codigo;
END;
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
