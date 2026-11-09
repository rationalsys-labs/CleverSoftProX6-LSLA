-- NombreBase se proporciona mediante sqlcmd -v. Nunca sobrescribe una base.
USE master;
GO
IF LEN(N'$(NombreBase)')=0 OR LEN(N'$(NombreBase)')>100 OR N'$(NombreBase)' LIKE N'%[^A-Za-z0-9_]%' COLLATE Latin1_General_BIN
BEGIN RAISERROR('NombreBase admite letras, numeros y guion bajo.',16,1); RETURN; END;
IF DB_ID(N'$(NombreBase)') IS NOT NULL
BEGIN RAISERROR('La base ya existe. La instalacion requiere un nombre nuevo y no reemplaza datos.',16,1); RETURN; END;
GO
CREATE DATABASE [$(NombreBase)] COLLATE Modern_Spanish_CI_AS;
GO
ALTER DATABASE [$(NombreBase)] SET COMPATIBILITY_LEVEL = 100;
GO
