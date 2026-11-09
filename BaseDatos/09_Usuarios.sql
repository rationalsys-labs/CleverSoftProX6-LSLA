-- Usuarios, perfiles, permisos y autenticacion de Clever SoftPro.
SET NOCOUNT ON;
SET XACT_ABORT ON;
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
BEGIN TRANSACTION;
IF COL_LENGTH('dbo.MdbUsers','Email') IS NULL
 ALTER TABLE dbo.MdbUsers ADD Email nvarchar(150) NOT NULL CONSTRAINT DF_MdbUsers_Email DEFAULT(N'');
IF OBJECT_ID('dbo.IntAccesos','U') IS NULL
 CREATE TABLE dbo.IntAccesos (
  CodUsu int NOT NULL REFERENCES dbo.MdbUsers(ITEM),
  CodMod int NOT NULL CHECK(CodMod BETWEEN 1 AND 11),
  Permitido bit NOT NULL,
  CONSTRAINT PK_IntAccesos PRIMARY KEY(CodUsu,CodMod)
 );
GO
IF OBJECT_ID('dbo.GuardarUsuario','P') IS NULL EXEC('CREATE PROCEDURE dbo.GuardarUsuario AS RETURN 0');
GO
ALTER PROCEDURE dbo.GuardarUsuario
 @Actor int,@Codigo int,@Accion int,@Nombre nvarchar(50)=N'',@Apellido nvarchar(50)=N'',
 @Usuario nvarchar(5)=N'',@Clave varchar(32)='',@Perfil tinyint=1,@Activo bit=1,
 @Correo nvarchar(150)=N'',@Permisos xml=NULL
AS
BEGIN
 SET NOCOUNT ON; SET XACT_ABORT ON;
 BEGIN TRY
  BEGIN TRANSACTION;
  -- Serializa altas/bajas y cambios de perfil para proteger al ultimo administrador.
  DECLARE @Administradores int;
  SELECT @Administradores=COUNT(*) FROM dbo.MdbUsers WITH(UPDLOCK,HOLDLOCK) WHERE [Access]=0 AND SitUsu=1 AND ActPas=1;
  IF NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE ITEM=@Actor AND [Access]=0 AND SitUsu=1 AND ActPas=1)
   RAISERROR('Solo un administrador activo puede gestionar usuarios.',16,1);
  IF @Accion NOT BETWEEN 1 AND 5 RAISERROR('Operacion de usuario no valida.',16,1);
  IF @Accion<>1 AND NOT EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE ITEM=@Codigo) RAISERROR('Seleccione un usuario existente.',16,1);
  IF @Accion IN(1,2) BEGIN
   SET @Nombre=LTRIM(RTRIM(@Nombre)); SET @Apellido=LTRIM(RTRIM(@Apellido));
   IF @Nombre=N'' OR @Apellido=N'' OR @Perfil NOT IN(0,1) RAISERROR('Complete nombres, apellidos y perfil.',16,1);
  END;
  IF @Accion IN(1,4) AND (LEN(@Clave)<>32 OR @Clave LIKE '%[^0-9a-f]%' COLLATE Latin1_General_BIN)
   RAISERROR('La clave de usuario no es valida.',16,1);
  IF @Accion=1 BEGIN
   SET @Usuario=UPPER(LTRIM(RTRIM(@Usuario)));
   IF LEN(@Usuario) NOT BETWEEN 1 AND 5 OR @Usuario LIKE N'%[^A-Z0-9]%' COLLATE Latin1_General_BIN RAISERROR('Usuario: de 1 a 5 letras o numeros.',16,1);
   INSERT dbo.MdbUsers(NomUsu,ApeUsu,Usuari,Paswor,ActPas,[Access],SitUsu,Email)
    VALUES(@Nombre,@Apellido,@Usuario,@Clave,1,@Perfil,1,@Correo);
   SET @Codigo=CONVERT(int,SCOPE_IDENTITY());
  END;
  IF @Accion IN(2,3) BEGIN
   IF @Accion=3 SET @Activo=0;
   IF @Codigo=@Actor AND (@Activo=0 OR (@Accion=2 AND @Perfil<>0)) RAISERROR('No puede dar de baja su cuenta ni quitarse la administracion durante su sesion.',16,1);
   IF EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE ITEM=@Codigo AND [Access]=0 AND SitUsu=1 AND ActPas=1)
    AND (@Activo=0 OR (@Accion=2 AND @Perfil<>0)) AND @Administradores<=1
    RAISERROR('Debe conservar al menos un administrador activo.',16,1);
   IF @Accion=2 UPDATE dbo.MdbUsers SET NomUsu=@Nombre,ApeUsu=@Apellido,[Access]=@Perfil,SitUsu=@Activo,Email=@Correo WHERE ITEM=@Codigo;
   ELSE UPDATE dbo.MdbUsers SET SitUsu=0 WHERE ITEM=@Codigo;
  END;
  IF @Accion=4 UPDATE dbo.MdbUsers SET Paswor=@Clave,ActPas=1 WHERE ITEM=@Codigo;
  IF @Accion=5 BEGIN
   IF @Permisos IS NULL OR @Permisos.value('count(/Permisos/Modulo)','int')<>11
    OR EXISTS(SELECT 1 FROM @Permisos.nodes('/Permisos/Modulo') t(n) WHERE n.value('@Codigo','int') NOT BETWEEN 1 AND 11 OR n.value('@Permitir','int') NOT IN(0,1))
    OR (SELECT COUNT(DISTINCT n.value('@Codigo','int')) FROM @Permisos.nodes('/Permisos/Modulo') t(n))<>11
    RAISERROR('El listado de permisos esta incompleto.',16,1);
   DELETE dbo.IntAccesos WHERE CodUsu=@Codigo;
   INSERT dbo.IntAccesos(CodUsu,CodMod,Permitido)
    SELECT @Codigo,n.value('@Codigo','int'),n.value('@Permitir','bit') FROM @Permisos.nodes('/Permisos/Modulo') t(n);
  END;
  COMMIT TRANSACTION;
  SELECT @Codigo Codigo;
 END TRY
 BEGIN CATCH
  IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
  DECLARE @Error nvarchar(2048); SET @Error=ERROR_MESSAGE(); RAISERROR('%s',16,1,@Error);
 END CATCH;
END;
GO
ALTER PROCEDURE dbo.Valid_LoginClev @Usuario varchar(5),@Pasword varchar(150)
AS
BEGIN
 SET NOCOUNT ON;
 SELECT 'MenSQL' Modal,CASE WHEN EXISTS(SELECT 1 FROM dbo.MdbUsers WHERE Usuari=@Usuario AND Paswor=@Pasword AND SitUsu=1 AND ActPas=1) THEN 1 ELSE 0 END Mensaje;
END;
GO
COMMIT TRANSACTION;
GO
