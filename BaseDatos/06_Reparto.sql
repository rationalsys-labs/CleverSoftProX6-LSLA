SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF OBJECT_ID(N'dbo.RegistrarReparto',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.RegistrarReparto AS RETURN 0;');
GO
ALTER PROCEDURE dbo.RegistrarReparto @Usuario int,@Mensajero int,@Modalidad tinyint,@Documentos nvarchar(max)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  IF @Modalidad IS NULL OR @Modalidad NOT IN(1,2) RAISERROR('Modalidad de reparto no valida.',16,1);
  IF NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WITH(HOLDLOCK) WHERE ITEM=@Usuario AND SitUsu=1)
   RAISERROR('El usuario operativo no esta disponible.',16,1);
  IF NOT EXISTS(SELECT 1 FROM dbo.Mdbmensajeros WITH(HOLDLOCK) WHERE CodMen=@Mensajero AND Active=1)
   RAISERROR('Seleccione un mensajero activo.',16,1);
  IF NOT EXISTS(SELECT 1 FROM dbo.IntStatus WHERE ITEM=1 AND Active=1)
   RAISERROR('El estado Reparto no esta disponible.',16,1);
  DECLARE @Lista xml; SET @Lista=CONVERT(xml,@Documentos);
  DECLARE @Codigos TABLE(Codigo numeric(18,0) PRIMARY KEY);
  INSERT @Codigos SELECT Fila.value('@Codigo','numeric(18,0)') FROM @Lista.nodes('/Documentos/Documento') Lista(Fila);
  IF NOT EXISTS(SELECT 1 FROM @Codigos) OR EXISTS(SELECT 1 FROM @Codigos WHERE Codigo<=0)
   RAISERROR('Ingrese documentos validos para el reparto.',16,1);
  DECLARE @Total int,@Esperados int,@Ranche numeric(18,0),@Fecha datetime;
  SELECT @Esperados=COUNT(*) FROM @Codigos;
  SELECT @Total=COUNT(*) FROM dbo.MDBDocument d WITH(UPDLOCK,HOLDLOCK) JOIN @Codigos c ON c.Codigo=d.NroCod WHERE d.Active=1;
  IF @Total<>@Esperados RAISERROR('Un documento no existe o fue anulado. Actualice la lista.',16,1);
  IF EXISTS(SELECT 1 FROM dbo.MDBHistory h JOIN @Codigos c ON c.Codigo=h.NroCod WHERE h.Status=1 AND h.Active=1)
   RAISERROR('Una guia ya esta asignada. Debe liberarla antes de volver a repartirla.',16,1);
  SET @Ranche=0;
  IF @Modalidad=1 EXEC dbo.ObtenerCorrelativo N'CRA',@Ranche OUTPUT;
  SET @Fecha=GETDATE();
  DECLARE @Nuevos TABLE(Item int);
  INSERT dbo.MDBHistory(NroCod,Status,FecEmi,Usuari,Ranche,Informe,CodMen)
   OUTPUT inserted.ITEM INTO @Nuevos
   SELECT Codigo,1,@Fecha,@Usuario,@Ranche,N'',@Mensajero FROM @Codigos;
  DECLARE @Item numeric(18,0);
  DECLARE Historial CURSOR LOCAL FAST_FORWARD FOR SELECT Item FROM @Nuevos ORDER BY Item;
  OPEN Historial; FETCH NEXT FROM Historial INTO @Item;
  WHILE @@FETCH_STATUS=0 BEGIN
   EXEC dbo.RegistrarActualizacion N'MDBHistory',@Item;
   FETCH NEXT FROM Historial INTO @Item;
  END;
  CLOSE Historial; DEALLOCATE Historial;
  COMMIT;
  SELECT @Ranche Ranche,@Fecha Fecha;
 END TRY
 BEGIN CATCH
  DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
  IF XACT_STATE()<>0 ROLLBACK;
  RAISERROR(@Error,16,1);
 END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.LiberarGuia',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.LiberarGuia AS RETURN 0;');
GO
ALTER PROCEDURE dbo.LiberarGuia @Codigo numeric(18,0),@Usuario int
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  IF NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE ITEM=@Usuario AND SitUsu=1)
   RAISERROR('El usuario operativo no esta disponible.',16,1);
  IF NOT EXISTS(SELECT 1 FROM dbo.MDBDocument WITH(UPDLOCK,HOLDLOCK) WHERE NroCod=@Codigo AND Active=1)
   RAISERROR('La guia no existe o fue anulada.',16,1);
  DECLARE @Historial TABLE(Item int);
  UPDATE dbo.MDBHistory SET Active=0 OUTPUT inserted.ITEM INTO @Historial WHERE NroCod=@Codigo AND Status=1 AND Active=1;
  IF @@ROWCOUNT=0 RAISERROR('La guia no tiene una asignacion activa de reparto.',16,1);
  DECLARE @Item numeric(18,0);
  DECLARE Historial CURSOR LOCAL FAST_FORWARD FOR SELECT Item FROM @Historial ORDER BY Item;
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