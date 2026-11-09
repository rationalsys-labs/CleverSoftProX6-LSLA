SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
IF OBJECT_ID(N'dbo.ViewDatosServicios',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDatosServicios AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDatosServicios AS
SELECT s.CodSer,s.NroBol,c.NomCli,s.CodCli,c.RucDni,c.DirCli,c.DisCli,s.TipSer,s.CanSer,s.Tipo,s.Estado,ISNULL(s.Grupo,0) Grupo
FROM dbo.MDBServices s JOIN dbo.MDBCostumer c ON c.CodCli=s.CodCli;
GO
IF OBJECT_ID(N'dbo.ViewDatosDocumentos',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDatosDocumentos AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDatosDocumentos AS
SELECT d.FecEmi,d.NroCod,d.CoServ,d.Docmto,d.CoProd,s.CodCli,e.NomEmp,e.DirEmp,e.DisEmp,e.ConEmp,d.Descrip,d.Inform,d.Active,d.CodEmp
FROM dbo.MDBDocument d JOIN dbo.MDBServices s ON s.CodSer=d.CoServ JOIN dbo.MDBHisEmp e ON e.CodEmp=d.CodEmp;
GO
IF OBJECT_ID(N'dbo.ViewDatosHistorial',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDatosHistorial AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDatosHistorial AS
SELECT h.ITEM,h.NroCod,h.Status,h.FecEmi,h.Usuari,h.Ranche,
CASE WHEN h.CodMen IS NOT NULL AND h.Informe=N'' THEN m.NomMen ELSE h.Informe END Informe,
h.Active,h.CodMen,h.Informe Observacion
FROM dbo.MDBHistory h LEFT JOIN dbo.Mdbmensajeros m ON m.CodMen=h.CodMen;
GO
IF OBJECT_ID(N'dbo.ViewUltimoStatus',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewUltimoStatus AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewUltimoStatus AS
SELECT d.NroCod Codigo,ISNULL(s.Status,N'') Status,h.FecEmi Fecha,
ISNULL(CASE WHEN h.Informe=N'REGISTERPRIVATE' THEN N'' ELSE h.Informe END,N'') Informe,
ISNULL(h.Active,CONVERT(bit,0)) Active,ISNULL(h.ITEM,0) Item,h.Status CodEstado,h.CodMen
FROM dbo.MDBDocument d
OUTER APPLY(SELECT TOP(1) x.* FROM dbo.ViewDatosHistorial x WHERE x.NroCod=d.NroCod AND x.Active=1 ORDER BY x.FecEmi DESC,x.ITEM DESC)h
LEFT JOIN dbo.IntStatus s ON s.ITEM=h.Status;
GO
IF OBJECT_ID(N'dbo.ViewUltimosStatus',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewUltimosStatus AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewUltimosStatus AS
SELECT Codigo NroCod,Active,CONVERT(varchar(19),Fecha,120)+N' - '+Status+N' - '+Informe Status,Informe,Fecha FecEmi,Status Estado,Item,CodEstado,CodMen
FROM dbo.ViewUltimoStatus WHERE Item<>0;
GO
IF OBJECT_ID(N'dbo.InforCostumer',N'V') IS NULL EXEC(N'CREATE VIEW dbo.InforCostumer AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.InforCostumer
AS
SELECT     dbo.MDBCostumer.CodCli AS Asiento, UPPER(dbo.MDBCostumer.NomCli) AS Cliente, dbo.MDBCostumer.RucDni AS RUC, 
                      UPPER(dbo.MDBCostumer.DirCli) AS Direccion, dbo.MDBDistritos.NomDis AS Distrito
FROM         dbo.MDBCostumer INNER JOIN
                      dbo.MDBDistritos ON dbo.MDBCostumer.DisCli = dbo.MDBDistritos.CoDist
WHERE     (dbo.MDBCostumer.ActCli = 1) AND (dbo.MDBCostumer.Active = 1)

GO
IF OBJECT_ID(N'dbo.Prueba',N'V') IS NULL EXEC(N'CREATE VIEW dbo.Prueba AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.Prueba
AS
SELECT     dbo.ViewDatosDocumentos.NroCod AS Codigo, dbo.ViewDatosDocumentos.CoServ AS Servicio, dbo.ViewDatosServicios.NomCli AS Cliente, 
                      dbo.MDBDistritos.NomDis AS Distrito, ISNULL(dbo.IntStatus.Status, 'NONE') AS Status, ISNULL(dbo.ViewDatosHistorial.FecEmi, '') AS Fecha
FROM         dbo.ViewDatosDocumentos INNER JOIN
                      dbo.ViewDatosServicios ON dbo.ViewDatosDocumentos.CoServ = dbo.ViewDatosServicios.CodSer INNER JOIN
                      dbo.MDBDistritos ON dbo.ViewDatosDocumentos.DisEmp = dbo.MDBDistritos.CoDist LEFT OUTER JOIN
                      dbo.ViewDatosHistorial ON dbo.ViewDatosDocumentos.NroCod = dbo.ViewDatosHistorial.NroCod LEFT OUTER JOIN
                      dbo.IntStatus ON dbo.ViewDatosHistorial.Status = dbo.IntStatus.ITEM

GO
IF OBJECT_ID(N'dbo.View_1',N'V') IS NULL EXEC(N'CREATE VIEW dbo.View_1 AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.View_1
AS
SELECT     dbo.ViewDatosDocumentos.NroCod, dbo.ViewDatosDocumentos.NomEmp, dbo.ViewUltimoStatus.Status, dbo.ViewUltimoStatus.Fecha, dbo.ViewUltimoStatus.Informe, 
                      dbo.ViewDatosDocumentos.DirEmp
FROM         dbo.ViewDatosDocumentos INNER JOIN
                      dbo.ViewUltimoStatus ON dbo.ViewDatosDocumentos.NroCod = dbo.ViewUltimoStatus.Codigo
WHERE     (dbo.ViewDatosDocumentos.NomEmp LIKE '%cmsj%')

GO
IF OBJECT_ID(N'dbo.ViewCliente',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewCliente AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewCliente
AS
SELECT     UPPER(dbo.MDBCostumer.NomCli) AS Cliente, dbo.MDBCostumer.CodCli AS Codigo, UPPER(dbo.MDBCostumer.DirCli) AS Direccion, 
                      dbo.MDBDistritos.NomDis AS Distrito, dbo.IntTipDocument.DesTip AS Tipo, dbo.MDBCostumer.RucDni AS Documento, 
                      dbo.MDBCostumer.FecApe AS Apertura, dbo.MDBCostumer.RepLeg AS Representante, dbo.MDBCostumer.MaiRep AS Mail, 
                      dbo.MDBCostumer.ConCli AS Contacto, dbo.MDBCostumer.Email AS MailContacto, dbo.MDBCostumer.TelCli AS Telefono, 
                      dbo.MDBCostumer.TelTra AS TelTrabajo, dbo.MDBCostumer.TelMov AS TelMovil, dbo.MDBCostumer.TelFax, 
                      CASE dbo.MDBCostumer.ActCli WHEN 1 THEN 'Activo' ELSE 'Inactivo' END AS Estado
FROM         dbo.MDBCostumer INNER JOIN
                      dbo.IntTipDocument ON dbo.MDBCostumer.TipDoc = dbo.IntTipDocument.CodTip INNER JOIN
                      dbo.MDBDistritos ON dbo.MDBCostumer.DisCli = dbo.MDBDistritos.CoDist
WHERE     (dbo.MDBCostumer.Active = 1)

GO
IF OBJECT_ID(N'dbo.ViewConfirmDocument',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewConfirmDocument AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewConfirmDocument
AS
SELECT     dbo.ViewDatosDocumentos.NroCod AS Codigo, dbo.MDBProduct.DesPro AS Producto, dbo.ViewDatosDocumentos.NomEmp AS Empresa, 
                      dbo.ViewDatosDocumentos.DirEmp AS Direccion, dbo.MDBDistritos.NomDis AS Distrito, dbo.ViewDatosDocumentos.ConEmp AS Contacto, 
                      dbo.ViewDatosDocumentos.CoServ AS Servicio, ISNULL(dbo.ViewUltimosStatus.Status, '') AS Status
FROM         dbo.MDBDistritos INNER JOIN
                      dbo.ViewDatosDocumentos INNER JOIN
                      dbo.MDBProduct ON dbo.ViewDatosDocumentos.CoProd = dbo.MDBProduct.CodPro ON dbo.MDBDistritos.CoDist = dbo.ViewDatosDocumentos.DisEmp INNER JOIN
                      dbo.ViewDatosServicios ON dbo.ViewDatosDocumentos.CoServ = dbo.ViewDatosServicios.CodSer LEFT OUTER JOIN
                      dbo.ViewUltimosStatus ON dbo.ViewDatosDocumentos.NroCod = dbo.ViewUltimosStatus.NroCod
WHERE     (dbo.ViewDatosServicios.Estado = 0)

GO
IF OBJECT_ID(N'dbo.ViewControlDelivery',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewControlDelivery AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewControlDelivery AS
SELECT d.NroCod Codigo,d.Descrip,d.NomEmp Empresa,d.ConEmp Contacto,h.FecEmi Emision,
 ISNULL(m.NomMen,h.Informe) Informe,u.Status,u.Fecha,h.CodMen,u.CodEstado
FROM dbo.ViewDatosDocumentos d
CROSS APPLY(SELECT TOP(1) x.* FROM dbo.MDBHistory x WHERE x.NroCod=d.NroCod AND x.Status=1 AND x.Active=1 ORDER BY x.FecEmi DESC,x.ITEM DESC) h
LEFT JOIN dbo.Mdbmensajeros m ON m.CodMen=h.CodMen
JOIN dbo.ViewUltimoStatus u ON u.Codigo=d.NroCod
WHERE d.Active=1;
GO
IF OBJECT_ID(N'dbo.ViewDateClient',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDateClient AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDateClient
AS
SELECT     'Distritos' AS Base, NomDis AS Codice
FROM         dbo.MDBDistritos
UNION
SELECT     'Tipo' AS Base, DesTip
FROM         dbo.IntTipDocument
UNION
SELECT     'Producto' AS Base, DesPro
FROM         dbo.MDBProduct
UNION
SELECT     'Servicio' AS Base, DesTip
FROM         dbo.IntTipService
UNION
SELECT     'TipDoc' AS Base, DesTip
FROM         dbo.IntTipInvoice

GO
IF OBJECT_ID(N'dbo.ViewDatos',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDatos AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDatos
AS
SELECT     'Productos' AS Origen, DesPro
FROM         dbo.MDBProduct
WHERE     (Active = 1)
UNION
SELECT     'TipoServicio' AS Origen, DesTip
FROM         dbo.IntTipService
WHERE     (Active = 1)
UNION
SELECT     'Distrito' AS Origen, NomDis
FROM         dbo.MDBDistritos
WHERE     (Active = 1)

GO
IF OBJECT_ID(N'dbo.ViewDocument',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDocument AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDocument
AS
SELECT     dbo.ViewDatosDocumentos.NroCod AS Codigo, dbo.MDBProduct.DesPro AS Producto, dbo.ViewDatosDocumentos.NomEmp AS Empresa, 
                      dbo.ViewDatosDocumentos.DirEmp AS Direccion, dbo.MDBDistritos.NomDis AS Distrito, dbo.ViewDatosDocumentos.ConEmp AS Contacto, 
                      dbo.ViewDatosDocumentos.FecEmi AS Fecha, dbo.ViewDatosServicios.NomCli AS Cliente, dbo.ViewDatosDocumentos.Descrip AS Descripcion, 
                      ISNULL(dbo.ViewUltimoStatus.Status, '') AS Status, dbo.ViewDatosServicios.NroBol AS Boleta, ISNULL
                          ((SELECT     TOP (1) Informe
                              FROM         dbo.ViewDatosHistorial
                              WHERE     (NroCod = dbo.ViewDatosDocumentos.NroCod) AND (Status = 1) AND (Active = 1) ORDER BY FecEmi DESC, ITEM DESC), '') AS Reparto, dbo.ViewDatosDocumentos.Docmto AS Documento, 
                      dbo.ViewDatosDocumentos.Inform AS Informacion
FROM         dbo.ViewDatosDocumentos INNER JOIN
                      dbo.MDBProduct ON dbo.ViewDatosDocumentos.CoProd = dbo.MDBProduct.CodPro INNER JOIN
                      dbo.MDBDistritos ON dbo.ViewDatosDocumentos.DisEmp = dbo.MDBDistritos.CoDist INNER JOIN
                      dbo.ViewDatosServicios ON dbo.ViewDatosDocumentos.CoServ = dbo.ViewDatosServicios.CodSer LEFT OUTER JOIN
                      dbo.ViewUltimoStatus ON dbo.ViewDatosDocumentos.NroCod = dbo.ViewUltimoStatus.Codigo

GO
IF OBJECT_ID(N'dbo.ViewGrupos',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewGrupos AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewGrupos
AS
SELECT     dbo.MdbGruposServ.Item AS Grupo, CONVERT(nvarchar(50), dbo.ViewDatosDocumentos.NroCod) AS Codigo, dbo.MDBProduct.DesPro AS Producto, 
                      dbo.ViewDatosDocumentos.NomEmp AS Empresa, dbo.ViewDatosDocumentos.DirEmp AS Direccion, dbo.MDBDistritos.NomDis AS Distrito, 
                      dbo.ViewDatosDocumentos.ConEmp AS Contacto, dbo.ViewDatosDocumentos.Descrip AS Descripcion, CONVERT(nvarchar(50), dbo.ViewDatosServicios.CodSer) AS Servicio, 
                      CONVERT(nvarchar(50), dbo.ViewDatosServicios.NroBol) AS Boleta, dbo.ViewDatosServicios.NomCli AS Cliente, dbo.ViewUltimoStatus.Status AS UltimoSta, 
                      dbo.ViewUltimoStatus.Fecha, dbo.ViewUltimoStatus.Informe, dbo.ViewDatosDocumentos.Docmto
FROM         dbo.MdbGruposServ INNER JOIN
                      dbo.ViewDatosServicios ON dbo.MdbGruposServ.Item = dbo.ViewDatosServicios.Grupo INNER JOIN
                      dbo.ViewDatosDocumentos ON dbo.ViewDatosServicios.CodSer = dbo.ViewDatosDocumentos.CoServ INNER JOIN
                      dbo.MDBProduct ON dbo.ViewDatosDocumentos.CoProd = dbo.MDBProduct.CodPro INNER JOIN
                      dbo.MDBDistritos ON dbo.ViewDatosDocumentos.DisEmp = dbo.MDBDistritos.CoDist LEFT OUTER JOIN
                      dbo.ViewUltimoStatus ON dbo.ViewDatosDocumentos.NroCod = dbo.ViewUltimoStatus.Codigo
WHERE dbo.ViewDatosDocumentos.Active=1;

GO
IF OBJECT_ID(N'dbo.ViewHistory',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewHistory AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewHistory
AS
SELECT     TOP (100) PERCENT dbo.ViewDatosHistorial.NroCod AS Codigo, dbo.IntStatus.Status, dbo.ViewDatosHistorial.FecEmi AS Emision, dbo.MdbUsers.Usuari AS Usuario, 
                      dbo.ViewDatosHistorial.Ranche, CASE dbo.ViewDatosHistorial.Informe WHEN 'REGISTERPRIVATE' THEN '' ELSE dbo.ViewDatosHistorial.Informe END AS Informe, 
                      dbo.ViewDatosHistorial.Active, dbo.ViewDatosDocumentos.Docmto AS Documento, dbo.ViewDatosDocumentos.CoServ AS Servicio, dbo.ViewDatosDocumentos.NomEmp AS Empresa, 
                      dbo.ViewDatosDocumentos.Descrip AS Descripcion, dbo.ViewDatosHistorial.ITEM AS Item, dbo.ViewDatosHistorial.Status AS CodEstado, dbo.ViewDatosHistorial.CodMen
FROM         dbo.ViewDatosHistorial INNER JOIN
                      dbo.IntStatus ON dbo.ViewDatosHistorial.Status = dbo.IntStatus.ITEM INNER JOIN
                      dbo.MdbUsers ON dbo.ViewDatosHistorial.Usuari = dbo.MdbUsers.ITEM INNER JOIN
                      dbo.ViewDatosDocumentos ON dbo.ViewDatosHistorial.NroCod = dbo.ViewDatosDocumentos.NroCod
WHERE     (dbo.ViewDatosHistorial.Active = 1)

GO
IF OBJECT_ID(N'dbo.ViewHistoryEmp',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewHistoryEmp AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewHistoryEmp
AS
SELECT     dbo.MDBHisEmp.CodEmp AS Codigo, dbo.MDBHisEmp.NomEmp AS Nombre, dbo.MDBHisEmp.DirEmp AS Direccion, 
                      dbo.MDBDistritos.NomDis AS Distrito, dbo.MDBHisEmp.ConEmp AS Contacto
FROM         dbo.MDBHisEmp INNER JOIN
                      dbo.MDBDistritos ON dbo.MDBHisEmp.DisEmp = dbo.MDBDistritos.CoDist

GO
IF OBJECT_ID(N'dbo.ViewRegenRanche',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewRegenRanche AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewRegenRanche AS
SELECT h.NroCod Codigo,p.DesPro Producto,d.NomEmp Empresa,d.DirEmp Direccion,
 dis.NomDis Distrito,d.ConEmp Contacto,h.FecEmi Fecha,h.Ranche,
 ISNULL(m.NomMen,h.Informe) Informe,s.NomCli Cliente,h.Active
FROM dbo.MDBHistory h
JOIN dbo.ViewDatosDocumentos d ON d.NroCod=h.NroCod
JOIN dbo.MDBProduct p ON p.CodPro=d.CoProd
JOIN dbo.MDBDistritos dis ON dis.CoDist=d.DisEmp
JOIN dbo.ViewDatosServicios s ON s.CodSer=d.CoServ
LEFT JOIN dbo.Mdbmensajeros m ON m.CodMen=h.CodMen
WHERE h.Ranche>0 AND h.Active=1 AND h.Status=1 AND d.Active=1
AND NOT EXISTS(SELECT 1 FROM dbo.MDBHistory otro WHERE otro.NroCod=h.NroCod AND otro.Ranche=h.Ranche
 AND otro.Status=1 AND otro.Active=1 AND (otro.FecEmi>h.FecEmi OR (otro.FecEmi=h.FecEmi AND otro.ITEM>h.ITEM)));
GO
IF OBJECT_ID(N'dbo.ViewRegistAct',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewRegistAct AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewRegistAct
AS
SELECT     TOP (100) PERCENT dbo.ViewDatosDocumentos.NroCod AS Codigo, dbo.ViewDatosServicios.NomCli AS Cliente, dbo.ViewDatosDocumentos.CodCli AS CodCliente, 
                      dbo.MDBProduct.DesPro AS Producto, dbo.ViewDatosDocumentos.Docmto AS Registro, dbo.ViewDatosDocumentos.NomEmp AS Empresa, 
                      dbo.ViewDatosDocumentos.DirEmp AS Direccion, dbo.MDBDistritos.NomDis AS Distrito, dbo.ViewDatosDocumentos.ConEmp AS Contacto, 
                      dbo.ViewDatosDocumentos.CoServ AS Servicio, dbo.ViewDatosDocumentos.FecEmi AS Fecha
FROM         dbo.ViewDatosDocumentos INNER JOIN
                      dbo.ViewDatosServicios ON dbo.ViewDatosDocumentos.CoServ = dbo.ViewDatosServicios.CodSer INNER JOIN
                      dbo.MDBDistritos ON dbo.ViewDatosDocumentos.DisEmp = dbo.MDBDistritos.CoDist INNER JOIN
                      dbo.MDBProduct ON dbo.ViewDatosDocumentos.CoProd = dbo.MDBProduct.CodPro
WHERE dbo.ViewDatosDocumentos.Active = 1

GO
IF OBJECT_ID(N'dbo.ViewServices',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewServices AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewServices
AS
SELECT     dbo.ViewDatosServicios.CodSer AS Codigo, dbo.ViewDatosServicios.NroBol AS Boleta, dbo.ViewDatosServicios.NomCli AS Cliente, 
                      dbo.ViewDatosServicios.CodCLi AS CodCLiente, dbo.MDBProduct.DesPro AS Producto, dbo.ViewDatosServicios.CanSer AS Cantidad
FROM         dbo.ViewDatosServicios INNER JOIN
                      dbo.MDBProduct ON dbo.ViewDatosServicios.Tipo = dbo.MDBProduct.CodPro
WHERE     (dbo.ViewDatosServicios.Estado = 1) AND (dbo.ViewDatosServicios.CodSer <> 105)

GO
IF OBJECT_ID(N'dbo.ViewVisDocument',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewVisDocument AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewVisDocument
AS
SELECT     dbo.ViewDatosServicios.CodSer AS Servicio, dbo.ViewDatosServicios.NomCli AS Cliente, dbo.ViewDatosServicios.RucDni, dbo.ViewDatosServicios.DirCli AS Direccion, 
                      dbo.MDBDistritos.NomDis AS Distrito, dbo.IntTipService.DesTip AS Reparto, dbo.MDBProduct.DesPro AS Producto, 
                      dbo.ViewDatosServicios.CanSer AS Cantidad, CONVERT(nvarchar(20), dbo.ViewDatosDocumentos.NroCod) AS Codigo, dbo.ViewDatosDocumentos.Docmto AS Documento, 
                      MDBProduct_1.DesPro AS ProdCliente, dbo.ViewDatosDocumentos.NomEmp AS Empresa, dbo.ViewDatosDocumentos.DirEmp AS Direccion2, 
                      MDBDistritos_1.NomDis AS Distrito2, dbo.ViewDatosDocumentos.ConEmp AS Contacto, dbo.ViewDatosDocumentos.FecEmi AS Fecha, dbo.ViewDatosServicios.Estado, 
                      dbo.ViewDatosDocumentos.Descrip AS Descripcion, dbo.ViewDatosDocumentos.Inform, dbo.ViewDatosServicios.NroBol AS Boleta, 
                      dbo.ViewDatosDocumentos.CodCli AS CodCliente
FROM         dbo.ViewDatosServicios INNER JOIN
                      dbo.MDBDistritos ON dbo.ViewDatosServicios.DisCli = dbo.MDBDistritos.CoDist INNER JOIN
                      dbo.IntTipService ON dbo.ViewDatosServicios.TipSer = dbo.IntTipService.CodTip INNER JOIN
                      dbo.MDBProduct ON dbo.ViewDatosServicios.Tipo = dbo.MDBProduct.CodPro INNER JOIN
                      dbo.ViewDatosDocumentos ON dbo.ViewDatosServicios.CodSer = dbo.ViewDatosDocumentos.CoServ INNER JOIN
                      dbo.MDBProduct AS MDBProduct_1 ON dbo.ViewDatosDocumentos.CoProd = MDBProduct_1.CodPro INNER JOIN
                      dbo.MDBDistritos AS MDBDistritos_1 ON dbo.ViewDatosDocumentos.DisEmp = MDBDistritos_1.CoDist
WHERE dbo.ViewDatosDocumentos.Active = 1

GO
IF OBJECT_ID(N'dbo.ViewInformeStatus',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewInformeStatus AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewInformeStatus AS
SELECT d.NroCod Codigo,d.FecEmi Emision,p.DesPro Producto,d.Docmto,s.CodSer Servicio,s.NroBol Boleta,ps.DesPro ProdServ,
c.NomCli Cliente,s.CanSer Cantidad,t.DesTip Reparto,e.NomEmp Empresa,e.DirEmp Direccion,di.NomDis Distrito,e.ConEmp Contacto,
u.Status,u.Fecha,u.Informe Descripcion
FROM dbo.MDBDocument d JOIN dbo.MDBServices s ON s.CodSer=d.CoServ JOIN dbo.MDBCostumer c ON c.CodCli=s.CodCli
JOIN dbo.MDBHisEmp e ON e.CodEmp=d.CodEmp JOIN dbo.MDBDistritos di ON di.CoDist=e.DisEmp
JOIN dbo.MDBProduct p ON p.CodPro=d.CoProd JOIN dbo.MDBProduct ps ON ps.CodPro=s.Tipo
JOIN dbo.IntTipService t ON t.CodTip=s.TipSer LEFT JOIN dbo.ViewUltimoStatus u ON u.Codigo=d.NroCod WHERE d.Active=1;
GO
-- Etiquetas para mostrar y claves numericas para ComboBox.ItemData.
IF OBJECT_ID(N'dbo.ViewDateClient',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDateClient AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDateClient AS
SELECT N'Distritos' Base,NomDis Codice,CoDist Codigo FROM dbo.MDBDistritos
UNION ALL SELECT N'Tipo',DesTip,CodTip FROM dbo.IntTipDocument
UNION ALL SELECT N'Producto',DesPro,CodPro FROM dbo.MDBProduct
UNION ALL SELECT N'Servicio',DesTip,CodTip FROM dbo.IntTipService
UNION ALL SELECT N'TipDoc',DesTip,TipDoc FROM dbo.IntTipInvoice;
GO
IF OBJECT_ID(N'dbo.ViewDatos',N'V') IS NULL EXEC(N'CREATE VIEW dbo.ViewDatos AS SELECT 1 AS Inicial;');
GO
ALTER VIEW dbo.ViewDatos AS
SELECT N'Productos' Origen,DesPro,CodPro Codigo FROM dbo.MDBProduct WHERE Active=1
UNION ALL SELECT N'TipoServicio',DesTip,CodTip FROM dbo.IntTipService WHERE Active=1
UNION ALL SELECT N'Distrito',NomDis,CoDist FROM dbo.MDBDistritos WHERE Active=1;
GO
