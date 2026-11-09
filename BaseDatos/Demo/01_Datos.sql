-- Datos ficticios de demostracion. Periodo: mayo de 2011.
-- Solo admite una base instalada sin datos. No borra ni reemplaza informacion.
SET NOCOUNT ON;
SET XACT_ABORT ON;
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
IF DB_NAME() IN(N'master',N'model',N'msdb',N'tempdb')
 BEGIN RAISERROR('Seleccione una base de Clever SoftPro.',16,1); RETURN; END;
DECLARE @Comprobar nvarchar(max),@Ocupada bit; SET @Ocupada=0; SET @Comprobar=N'';
SELECT @Comprobar=@Comprobar+N'IF EXISTS(SELECT 1 FROM dbo.'+QUOTENAME(name)+N') SET @Ocupada=1;'
 FROM sys.tables WHERE is_ms_shipped=0;
EXEC sp_executesql @Comprobar,N'@Ocupada bit OUTPUT',@Ocupada OUTPUT;
IF @Ocupada=1 BEGIN RAISERROR('La carga demo exige todas las tablas vacias. No se modificaron los datos existentes.',16,1); RETURN; END;
BEGIN TRY
 BEGIN TRANSACTION;
 DECLARE @Clave varchar(32); SET @Clave=LOWER(CONVERT(varchar(32),HASHBYTES('MD5','Demo2011'),2));
 SET IDENTITY_INSERT dbo.ADMUsers ON;
 INSERT dbo.ADMUsers(ITEM,NomUsu,ApeUsu,Usuari,Paswor,ActPas,SitUsu) VALUES(1,N'Mateo',N'Salazar','ADMIN',@Clave,1,1);
 SET IDENTITY_INSERT dbo.ADMUsers OFF;
 SET IDENTITY_INSERT dbo.MdbUsers ON;
 INSERT dbo.MdbUsers(ITEM,NomUsu,ApeUsu,Usuari,Paswor,ActPas,[Access],SitUsu,Email) VALUES
 (1,N'Mateo',N'Salazar','ADMIN',@Clave,1,0,1,N'admin@demo.example'),
 (2,N'Clara',N'Valverde','GEST1',@Clave,1,0,1,N'gestion@demo.example'),
 (3,N'Lucia',N'Montalvo','OPER1',@Clave,1,1,1,N'operaciones@demo.example'),
 (4,N'Ramon',N'Cabrera','REPA1',@Clave,1,1,1,N'reparto@demo.example'),
 (5,N'Elena',N'Zarate','INFO1',@Clave,1,1,1,N'informes@demo.example'),
 (6,N'Victor',N'Aranda','BAJA1',@Clave,1,1,0,N'inactivo@demo.example');
 SET IDENTITY_INSERT dbo.MdbUsers OFF;
 DECLARE @Modulo int; SET @Modulo=1;
 WHILE @Modulo<=11 BEGIN
  INSERT dbo.IntAccesos(CodUsu,CodMod,Permitido) VALUES
  (3,@Modulo,1),(4,@Modulo,CASE WHEN @Modulo IN(5,6,7,8) THEN 1 ELSE 0 END),
  (5,@Modulo,CASE WHEN @Modulo IN(2,8,9) THEN 1 ELSE 0 END),(6,@Modulo,0);
  SET @Modulo=@Modulo+1;
 END;
 SET IDENTITY_INSERT dbo.IntTipDocument ON;
 INSERT dbo.IntTipDocument(CodTip,DesTip) VALUES(1,N'RUC'),(2,N'DNI');
 SET IDENTITY_INSERT dbo.IntTipDocument OFF;
 SET IDENTITY_INSERT dbo.IntTipInvoice ON;
 INSERT dbo.IntTipInvoice(TipDoc,DesTip,NamVar) VALUES(1,N'BO',N'Boleta'),(2,N'FA',N'Factura');
 SET IDENTITY_INSERT dbo.IntTipInvoice OFF;
 SET IDENTITY_INSERT dbo.IntTipService ON;
 INSERT dbo.IntTipService(CodTip,DesTip) VALUES(1,N'Distribucion local'),(2,N'Entrega expresa'),(3,N'Campana programada');
 SET IDENTITY_INSERT dbo.IntTipService OFF;
 SET IDENTITY_INSERT dbo.MDBProduct ON;
 INSERT dbo.MDBProduct(CodPro,DesPro) VALUES(1,N'Documentacion comercial'),(2,N'Correspondencia empresarial'),(3,N'Paquetes livianos'),(4,N'Invitaciones');
 SET IDENTITY_INSERT dbo.MDBProduct OFF;
 SET IDENTITY_INSERT dbo.IntStatus ON;
 INSERT dbo.IntStatus(ITEM,Status,ActEst,Active) VALUES
 (1,N'Reparto',1,1),(2,N'Entregado',1,1),(3,N'Ausente',1,1),(4,N'Direccion Incorrecta',1,1),
 (5,N'Rechazado',1,1),(6,N'Reprogramado',1,1),(7,N'Estado de muestra',1,0);
 SET IDENTITY_INSERT dbo.IntStatus OFF;
 SET IDENTITY_INSERT dbo.MDBDistritos ON;
 INSERT dbo.MDBDistritos(CoDist,CodPos,NomDis) VALUES
 (1,N'LIMA01',N'Lima'),(2,N'LIMA02',N'Lince'),(3,N'LIMA03',N'Jesus Maria'),(4,N'LIMA04',N'San Borja'),
 (5,N'LIMA05',N'Miraflores'),(6,N'LIMA06',N'Surquillo'),(7,N'LIMA07',N'Pueblo Libre'),(8,N'LIMA08',N'San Miguel');
 SET IDENTITY_INSERT dbo.MDBDistritos OFF;
 SET IDENTITY_INSERT dbo.Mdbmensajeros ON;
 INSERT dbo.Mdbmensajeros(CodMen,NomMen,Active) VALUES
 (1,N'Andrés Robles - Demo',1),(2,N'César Palacios - Demo',1),(3,N'Marina Torres - Demo',1),
 (4,N'Daniel Paredes - Demo',0),(5,N'Sergio Oliva - Demo',1);
 SET IDENTITY_INSERT dbo.Mdbmensajeros OFF;
 SET IDENTITY_INSERT dbo.MdbGruposServ ON;
 INSERT dbo.MdbGruposServ(Item,NomGrp) VALUES(1,N'RUTA COMERCIAL MAYO'),(2,N'CAMPANA INSTITUCIONAL'),(3,N'PEDIDO ESPECIAL MAYO');
 SET IDENTITY_INSERT dbo.MdbGruposServ OFF;
 SET IDENTITY_INSERT dbo.IntCorrel ON;
 INSERT dbo.IntCorrel(Item,DivCor,Correl,Concep) VALUES
 (1,N'CLI',7,N'Clientes'),(2,N'CSV',1108,N'Servicios'),(3,N'CHC',1029,N'Destinatarios'),
 (4,N'CDO',511010029,N'Documentos'),(5,N'CRA',505,N'Cargos de reparto');
 SET IDENTITY_INSERT dbo.IntCorrel OFF;
 SET IDENTITY_INSERT dbo.IntInfoSys ON;
 INSERT dbo.IntInfoSys(Item,VarMod,Infor1) VALUES(1,N'Empresa',N'LSLA - Lima'),(2,N'Ciudad',N'Lima, Peru'),(3,N'PeriodoDemo',N'Mayo de 2011');
 SET IDENTITY_INSERT dbo.IntInfoSys OFF;
 SET IDENTITY_INSERT dbo.IntBackup ON;
 INSERT dbo.IntBackup(Item,Archivo) VALUES(1,N'Demostracion_Mayo_2011.sql');
 SET IDENTITY_INSERT dbo.IntBackup OFF;
 SET IDENTITY_INSERT dbo.IntTarifasCourier ON;
 INSERT dbo.IntTarifasCourier(Item,Tarifa,Distri,TimExp,TimHot,Tim1Di,Tim2Di,Tim3Di,Tim4Di)
 SELECT CoDist,1,CoDist,9.00,7.50,5.00,4.50,4.00,3.50 FROM dbo.MDBDistritos;
 SET IDENTITY_INSERT dbo.IntTarifasCourier OFF;
 INSERT dbo.MDBCostumer(CodCli,NomCli,TipDoc,RucDni,FecApe,DirCli,DisCli,RepLeg,MaiRep,ConCli,TelCli,TelTra,TelMov,TelFax,Email,ActCli,Active)
 SELECT 511000000+n,Nombre,1,N'2099100000'+CONVERT(nvarchar(1),n),CONVERT(datetime,'20110502',112),
 N'Av. Demostracion '+CONVERT(nvarchar(10),100+n),n,N'Representante Demo '+CONVERT(nvarchar(1),n),
 N'representante'+CONVERT(nvarchar(1),n)+N'@demo.example',N'Contacto Demo '+CONVERT(nvarchar(1),n),
 N'01-555-010'+CONVERT(nvarchar(1),n),N'01-555-020'+CONVERT(nvarchar(1),n),N'90000000'+CONVERT(nvarchar(1),n),N'',
 N'cliente'+CONVERT(nvarchar(1),n)+N'@demo.example',CASE WHEN n=6 THEN 0 ELSE 1 END,CASE WHEN n=7 THEN 0 ELSE 1 END
 FROM (VALUES(1,N'Editorial Loma Azul S.A.C. - Demo'),(2,N'Comercial Valle Claro S.A.C. - Demo'),
 (3,N'Instituto Horizonte de Lima - Demo'),(4,N'Distribuidora Sendero S.R.L. - Demo'),
 (5,N'Taller Grafico Brisa S.A.C. - Demo'),(6,N'Oficina Prado Sereno - Demo'),(7,N'Comercial Archivo de Muestra - Demo')) c(n,Nombre);
 INSERT dbo.MDBServices(CodSer,NroBol,CodCli,TipSer,CanSer,Tipo,Estado,Grupo) VALUES
 (105,1100105,511000001,1,0,1,1,NULL),
 (1101,1101101,511000001,1,6,1,1,1),(1102,1101102,511000001,3,4,4,1,1),
 (1103,1101103,511000002,2,5,2,0,NULL),(1104,1101104,511000003,3,6,2,1,2),
 (1105,1101105,511000004,1,3,3,1,NULL),(1106,1101106,511000005,2,4,1,1,NULL),
 (1107,1101107,511000006,1,2,4,1,3),(1108,1101108,511000001,1,3,1,1,NULL);
 DECLARE @Destinos TABLE(n int PRIMARY KEY,Nombre nvarchar(150));
 INSERT @Destinos VALUES
 (1,N'Papelera Alameda'),(2,N'Almacen Brisa'),(3,N'Libreria Camino Claro'),(4,N'Oficina Laguna Azul'),
 (5,N'Galeria Las Lomas'),(6,N'Taller Punto Norte'),(7,N'Oficina Los Cedros'),(8,N'Consultorio Jardin'),
 (9,N'Academia Puente Claro'),(10,N'Asociacion El Olivar'),(11,N'Oficina Plaza Serena'),(12,N'Comercial La Ribera'),
 (13,N'Bodega Tres Estrellas'),(14,N'Estudio Amanecer'),(15,N'Centro Cultural Ronda'),(16,N'Oficina El Mirador'),
 (17,N'Almacen Nube Azul'),(18,N'Instituto Las Acacias'),(19,N'Libreria Mar de Papel'),(20,N'Oficina Pampa Verde'),
 (21,N'Centro de Idiomas Sendero'),(22,N'Comercial Naranjal'),(23,N'Oficina La Arboleda'),(24,N'Taller Horizonte'),
 (25,N'Galeria Portal del Sol'),(26,N'Oficina Las Violetas'),(27,N'Almacen Documento Anulado'),
 (28,N'Asociacion Colina Serena'),(29,N'Oficina Nuevo Camino');
 INSERT dbo.MDBHisEmp(CodEmp,NomEmp,DirEmp,DisEmp,ConEmp)
 SELECT 1000+n,Nombre+N' - Demo',N'Jr. Circuito de Muestra '+CONVERT(nvarchar(10),200+n),1+((n-1)%8),N'Contacto de muestra '+CONVERT(nvarchar(10),n) FROM @Destinos;
 INSERT dbo.MDBDocument(NroCod,CoServ,Docmto,CoProd,Descrip,Inform,Active,CodEmp,FecEmi)
 SELECT 511010000+n,
 CASE WHEN n<=2 THEN 105 WHEN n<=6 THEN 1101 WHEN n<=10 THEN 1102 WHEN n<=15 THEN 1103
 WHEN n<=21 THEN 1104 WHEN n<=23 THEN 1105 WHEN n<=27 THEN 1106 ELSE 1107 END,
 N'REF-2011-'+RIGHT(N'000'+CONVERT(nvarchar(3),n),3),
 CASE WHEN n=2 OR n BETWEEN 22 AND 23 THEN 3 WHEN n BETWEEN 7 AND 10 OR n>=28 THEN 4
 WHEN n BETWEEN 11 AND 21 THEN 2 ELSE 1 END,
 N'Distribucion de mayo de 2011 - Demo',N'',CASE WHEN n=27 THEN 0 ELSE 1 END,1000+n,
 DATEADD(minute,n,CONVERT(datetime,'2011-05-05T08:00:00',126)) FROM @Destinos;
 SET IDENTITY_INSERT dbo.MDBHistory ON;
 INSERT dbo.MDBHistory(ITEM,NroCod,Status,FecEmi,Usuari,Ranche,Informe,Active,CodMen)
 SELECT ROW_NUMBER() OVER(ORDER BY Fecha,Cargo,n),511010000+n,1,Fecha,4,Cargo,N'',Activo,Mensajero FROM
 (VALUES(3,501,1,1,'2011-05-10T09:00:00'),(4,501,1,1,'2011-05-10T09:00:00'),(5,501,1,1,'2011-05-10T09:00:00'),
 (11,502,2,1,'2011-05-12T08:30:00'),(12,502,2,1,'2011-05-12T08:30:00'),(13,502,2,1,'2011-05-12T08:30:00'),
 (14,502,2,1,'2011-05-12T08:30:00'),(15,502,2,1,'2011-05-12T08:30:00'),
 (16,0,3,1,'2011-05-13T09:15:00'),(17,0,3,1,'2011-05-13T09:15:00'),(18,0,3,1,'2011-05-13T09:15:00'),
 (19,503,1,1,'2011-05-13T10:00:00'),(20,503,1,1,'2011-05-13T10:00:00'),
 (24,500,4,0,'2011-05-06T09:00:00'),(24,504,2,1,'2011-05-16T08:00:00'),(25,504,2,1,'2011-05-16T08:00:00'),
 (26,505,4,0,'2011-05-06T10:00:00'),(27,505,4,0,'2011-05-06T10:00:00')) h(n,Cargo,Mensajero,Activo,Fecha);
 INSERT dbo.MDBHistory(ITEM,NroCod,Status,FecEmi,Usuari,Informe) VALUES
 (104,511010004,2,'2011-05-11T12:15:00',3,N'Recibido por el contacto de muestra.'),
 (105,511010005,3,'2011-05-11T14:30:00',3,N'Destinatario ausente. Se deja aviso.'),
 (111,511010011,2,'2011-05-12T11:30:00',3,N'Entrega conforme.'),(112,511010012,2,'2011-05-12T12:00:00',3,N'Recibido en recepcion.'),
 (113,511010013,2,'2011-05-12T12:45:00',3,N'Constancia firmada.'),(114,511010014,4,'2011-05-12T15:00:00',3,N'Numeracion no localizada.'),
 (116,511010016,2,'2011-05-14T10:20:00',3,N'Entrega de campana conforme.'),(117,511010017,4,'2011-05-14T11:15:00',3,N'Corregir referencia del destino.'),
 (119,511010019,5,'2011-05-14T15:10:00',3,N'El destinatario no acepta el documento.'),
 (124,511010024,2,'2011-05-18T09:40:00',3,N'Entrega conforme despues de reasignacion.');
 SET IDENTITY_INSERT dbo.MDBHistory OFF;
 INSERT dbo.IntSaveActualizar(Consulta,ActSav,VersionEsquema) VALUES(N'-- Archivo de demostracion: mayo de 2011',0,1);
 EXEC dbo.RegistrarActualizacion N'MDBCostumer',511000001;
 EXEC dbo.RegistrarActualizacion N'IntStatus',6;
 COMMIT TRANSACTION;
 PRINT 'OK: demostracion ficticia de mayo de 2011. Usuario ADMIN; clave Demo2011.';
END TRY
BEGIN CATCH
 IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
 DECLARE @Error nvarchar(2048); SET @Error=ERROR_MESSAGE(); RAISERROR('%s',16,1,@Error);
END CATCH;
GO
