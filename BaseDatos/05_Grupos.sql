SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF OBJECT_ID(N'dbo.CrearGrupo',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.CrearGrupo AS RETURN 0;');
GO
ALTER PROCEDURE dbo.CrearGrupo @Nombre nvarchar(150)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  SET @Nombre=UPPER(LTRIM(RTRIM(@Nombre)));
  IF @Nombre IS NULL OR LEN(@Nombre)=0 RAISERROR('Debe indicar el nombre del grupo.',16,1);
  IF EXISTS(SELECT 1 FROM dbo.MdbGruposServ WITH(UPDLOCK,HOLDLOCK) WHERE NomGrp=@Nombre)
   RAISERROR('Ya existe un grupo con ese nombre.',16,1);
  INSERT dbo.MdbGruposServ(NomGrp) VALUES(@Nombre);
  DECLARE @Codigo numeric(18,0); SET @Codigo=SCOPE_IDENTITY();
  EXEC dbo.RegistrarActualizacion N'MdbGruposServ',@Codigo;
  COMMIT;
  SELECT CONVERT(int,@Codigo) Codigo;
 END TRY
 BEGIN CATCH
  DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
  IF XACT_STATE()<>0 ROLLBACK;
  RAISERROR(@Error,16,1);
 END CATCH;
END;
GO
IF OBJECT_ID(N'dbo.AgruparServicios',N'P') IS NULL EXEC(N'CREATE PROCEDURE dbo.AgruparServicios AS RETURN 0;');
GO
ALTER PROCEDURE dbo.AgruparServicios @Grupo int,@Accion tinyint,@Servicios nvarchar(max)
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  IF @Accion NOT IN(1,2) OR @Accion IS NULL RAISERROR('Operacion de grupo no valida.',16,1);
  IF NOT EXISTS(SELECT 1 FROM dbo.MdbGruposServ WITH(UPDLOCK,HOLDLOCK) WHERE Item=@Grupo)
   RAISERROR('El grupo no existe.',16,1);
  DECLARE @Lista xml; SET @Lista=CONVERT(xml,@Servicios);
  DECLARE @Codigos TABLE(Codigo numeric(18,0) PRIMARY KEY);
  INSERT @Codigos SELECT Fila.value('@Codigo','numeric(18,0)') FROM @Lista.nodes('/Servicios/Servicio') AS Lista(Fila);
  IF NOT EXISTS(SELECT 1 FROM @Codigos) OR EXISTS(SELECT 1 FROM @Codigos WHERE Codigo<=0)
   RAISERROR('Seleccione servicios validos.',16,1);
  DECLARE @Esperados int; SELECT @Esperados=COUNT(*) FROM @Codigos;
  UPDATE s WITH(UPDLOCK,HOLDLOCK) SET Grupo=CASE WHEN @Accion=1 THEN @Grupo ELSE NULL END
   FROM dbo.MDBServices s JOIN @Codigos c ON c.Codigo=s.CodSer
   WHERE (@Accion=1 AND s.Grupo IS NULL) OR (@Accion=2 AND s.Grupo=@Grupo);
  IF @@ROWCOUNT<>@Esperados RAISERROR('Un servicio ya cambio de grupo o no existe. Actualice la lista.',16,1);
  DECLARE @Codigo numeric(18,0);
  DECLARE Servicios CURSOR LOCAL FAST_FORWARD FOR SELECT Codigo FROM @Codigos ORDER BY Codigo;
  OPEN Servicios; FETCH NEXT FROM Servicios INTO @Codigo;
  WHILE @@FETCH_STATUS=0 BEGIN
   EXEC dbo.RegistrarActualizacion N'MDBServices',@Codigo;
   FETCH NEXT FROM Servicios INTO @Codigo;
  END;
  CLOSE Servicios; DEALLOCATE Servicios;
  COMMIT;
 END TRY
 BEGIN CATCH
  DECLARE @Error nvarchar(4000); SET @Error=ERROR_MESSAGE();
  IF XACT_STATE()<>0 ROLLBACK;
  RAISERROR(@Error,16,1);
 END CATCH;
END;
GO
