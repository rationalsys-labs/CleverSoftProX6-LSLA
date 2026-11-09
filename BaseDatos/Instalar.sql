:on error exit
:r BaseDatos/00_CrearBase.sql
USE [$(NombreBase)];
GO
:r BaseDatos/01_Estructura.sql
:r BaseDatos/02_Procedimientos.sql
:r BaseDatos/03_Vistas.sql
:r BaseDatos/04_Servicios.sql
:r BaseDatos/05_Grupos.sql
:r BaseDatos/06_Reparto.sql
:r BaseDatos/07_Estados.sql
:r BaseDatos/08_InformesConfiguracion.sql
:r BaseDatos/09_Usuarios.sql
:r BaseDatos/10_Verificar.sql
PRINT 'Estructura de Clever SoftPro instalada. Cargue los datos de demostracion para iniciar sesion.';
GO
