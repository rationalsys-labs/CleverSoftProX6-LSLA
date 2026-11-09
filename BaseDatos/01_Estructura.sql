-- Estructura relacional de Clever SoftPro X6 - LSLA.
-- Ejecutar en una base vacia; los datos se cargan por separado.
SET XACT_ABORT ON;
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO
CREATE TABLE dbo.[ADMUsers] (
    [ITEM] int IDENTITY(1,1) NOT NULL,
    [NomUsu] nvarchar(50) NOT NULL,
    [ApeUsu] nvarchar(50) NOT NULL,
    [Usuari] nvarchar(5) NOT NULL,
    [Paswor] nvarchar(150) NOT NULL,
    [ActPas] bit NOT NULL CONSTRAINT [DF_ADMUsers_ActPas] DEFAULT ((1)),
    [SitUsu] bit NOT NULL CONSTRAINT [DF_ADMUsers_SitUsu] DEFAULT ((1)),
    CONSTRAINT [PK_ADMUsers] PRIMARY KEY CLUSTERED ([ITEM])
);
GO
CREATE TABLE dbo.[IntBackup] (
    [Item] int IDENTITY(1,1) NOT NULL,
    [Archivo] nvarchar(50) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_IntBackup_Active] DEFAULT ((1)),
    CONSTRAINT [PK_IntBackup] PRIMARY KEY CLUSTERED ([Item])
);
GO
CREATE TABLE dbo.[IntCorrel] (
    [Item] int IDENTITY(1,1) NOT NULL,
    [DivCor] nvarchar(3) NOT NULL,
    [Correl] numeric(18,0) NOT NULL,
    [Concep] nvarchar(50) NOT NULL,
    CONSTRAINT [PK_IntCorrel] PRIMARY KEY CLUSTERED ([Item])
);
GO
CREATE TABLE dbo.[IntInfoSys] (
    [Item] int IDENTITY(1,1) NOT NULL,
    [VarMod] nvarchar(50) NOT NULL,
    [Infor1] nvarchar(50) NOT NULL,
    CONSTRAINT [PK_IntInfoSys] PRIMARY KEY CLUSTERED ([Item])
);
GO
CREATE TABLE dbo.[IntSaveActualizar] (
    [Item] int IDENTITY(1,1) NOT NULL,
    [Consulta] nvarchar(max) NOT NULL,
    [ActSav] tinyint NOT NULL CONSTRAINT [DF_IntSaveActualizar_ActSav] DEFAULT ((1)),
    [VersionEsquema] smallint NOT NULL CONSTRAINT [DF_IntSaveActualizar_VersionEsquema] DEFAULT (2),
    CONSTRAINT [PK_IntSaveActualizar] PRIMARY KEY CLUSTERED ([Item])
);
GO
CREATE TABLE dbo.[IntStatus] (
    [ITEM] int IDENTITY(1,1) NOT NULL,
    [Status] nvarchar(20) NOT NULL,
    [ActEst] bit NOT NULL CONSTRAINT [DF_IntStatus_ActEst] DEFAULT ((1)),
    [Active] bit NOT NULL CONSTRAINT [DF_IntStatus_Active] DEFAULT ((1)),
    CONSTRAINT [PK_IntStatus] PRIMARY KEY CLUSTERED ([ITEM])
);
GO
CREATE TABLE dbo.[IntTarifasCourier] (
    [Item] int IDENTITY(1,1) NOT NULL,
    [Tarifa] int NOT NULL,
    [Distri] int NOT NULL,
    [TimExp] money NOT NULL CONSTRAINT [DF_IntTarifasCourier_TimExp] DEFAULT ((0.00)),
    [TimHot] money NOT NULL CONSTRAINT [DF_IntTarifasCourier_TimHot] DEFAULT ((0.00)),
    [Tim1Di] money NOT NULL CONSTRAINT [DF_IntTarifasCourier_Tim1Di] DEFAULT ((0.00)),
    [Tim2Di] money NOT NULL CONSTRAINT [DF_IntTarifasCourier_Tim2Di] DEFAULT ((0.00)),
    [Tim3Di] money NOT NULL CONSTRAINT [DF_IntTarifasCourier_Tim3Di] DEFAULT ((0.00)),
    [Tim4Di] money NOT NULL CONSTRAINT [DF_IntTarifasCourier_Tim4Di] DEFAULT ((0.00)),
    CONSTRAINT [PK_IntTarifasCourier] PRIMARY KEY CLUSTERED ([Item])
);
GO
CREATE TABLE dbo.[IntTipDocument] (
    [CodTip] int IDENTITY(1,1) NOT NULL,
    [DesTip] nvarchar(100) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_IntTipDocument_Active] DEFAULT ((1)),
    CONSTRAINT [PK_IntTipDocument] PRIMARY KEY CLUSTERED ([CodTip])
);
GO
CREATE TABLE dbo.[IntTipInvoice] (
    [TipDoc] int IDENTITY(1,1) NOT NULL,
    [DesTip] nvarchar(2) NOT NULL,
    [NamVar] nvarchar(50) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_IntTipInvoice_Active] DEFAULT ((1)),
    CONSTRAINT [PK_IntTipInvoice] PRIMARY KEY CLUSTERED ([TipDoc])
);
GO
CREATE TABLE dbo.[IntTipService] (
    [CodTip] int IDENTITY(1,1) NOT NULL,
    [DesTip] nvarchar(100) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_IntTipService_Active] DEFAULT ((1)),
    CONSTRAINT [PK_IntTipService] PRIMARY KEY CLUSTERED ([CodTip])
);
GO
CREATE TABLE dbo.[MDBCostumer] (
    [CodCli] numeric(18,0) NOT NULL,
    [NomCli] nvarchar(300) NOT NULL,
    [TipDoc] int NOT NULL,
    [RucDni] nvarchar(20) NOT NULL,
    [FecApe] datetime NOT NULL,
    [DirCli] nvarchar(300) NOT NULL,
    [DisCli] int NOT NULL,
    [RepLeg] nvarchar(150) NOT NULL,
    [MaiRep] nvarchar(150) NOT NULL,
    [ConCli] nvarchar(150) NOT NULL,
    [TelCli] nvarchar(50) NOT NULL,
    [TelTra] nvarchar(50) NOT NULL,
    [TelMov] nvarchar(50) NOT NULL,
    [TelFax] nvarchar(50) NOT NULL,
    [Email] nvarchar(150) NOT NULL,
    [ActCli] bit NOT NULL CONSTRAINT [DF_MDBCostumer_ActCli] DEFAULT ((1)),
    [Active] bit NOT NULL CONSTRAINT [DF_MDBCostumer_Active] DEFAULT ((1)),
    CONSTRAINT [PK_MDBCostumer] PRIMARY KEY CLUSTERED ([CodCli])
);
GO
CREATE TABLE dbo.[MDBDistritos] (
    [CoDist] int IDENTITY(1,1) NOT NULL,
    [CodPos] nvarchar(50) NOT NULL,
    [NomDis] nvarchar(150) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_MDBDistritos_Active] DEFAULT ((1)),
    CONSTRAINT [PK_MDBDistritos] PRIMARY KEY CLUSTERED ([CoDist])
);
GO
CREATE TABLE dbo.[MDBDocument] (
    [FecEmi] datetime NOT NULL CONSTRAINT [DF_MDBDocument_FecEmi] DEFAULT (getdate()),
    [NroCod] numeric(18,0) NOT NULL,
    [CoServ] numeric(18,0) NOT NULL,
    [Docmto] nvarchar(50) NOT NULL,
    [CoProd] int NOT NULL,
    [Descrip] nvarchar(150) NOT NULL CONSTRAINT [DF_MDBDocument_Descrip] DEFAULT (''),
    [Inform] nvarchar(150) NOT NULL CONSTRAINT [DF_MDBDocument_Inform] DEFAULT (''),
    [Active] bit NOT NULL CONSTRAINT [DF_MDBDocument_Active] DEFAULT ((1)),
    [CodEmp] numeric(18,0) NOT NULL,
    CONSTRAINT [PK_MDBDocument] PRIMARY KEY CLUSTERED ([NroCod])
);
GO
CREATE TABLE dbo.[MdbGruposServ] (
    [Item] int IDENTITY(1,1) NOT NULL,
    [NomGrp] nvarchar(150) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_MdbGruposServ_Active] DEFAULT ((0)),
    CONSTRAINT [PK_MdbGruposServ] PRIMARY KEY CLUSTERED ([Item])
);
GO
CREATE TABLE dbo.[MDBHisEmp] (
    [CodEmp] numeric(18,0) NOT NULL,
    [NomEmp] nvarchar(500) NOT NULL,
    [DirEmp] nvarchar(500) NOT NULL,
    [DisEmp] int NOT NULL,
    [ConEmp] nvarchar(150) NOT NULL,
    CONSTRAINT [PK_MDBHisEmp] PRIMARY KEY CLUSTERED ([CodEmp])
);
GO
CREATE TABLE dbo.[MDBHistory] (
    [ITEM] int IDENTITY(1,1) NOT NULL,
    [NroCod] numeric(18,0) NOT NULL,
    [Status] int NOT NULL,
    [FecEmi] datetime NOT NULL,
    [Usuari] int NOT NULL,
    [Ranche] numeric(18,0) NOT NULL CONSTRAINT [DF_MDBHistory_Ranche] DEFAULT ((0)),
    [Informe] nvarchar(500) NOT NULL CONSTRAINT [DF_MDBHistory_Informe] DEFAULT (''),
    [Active] bit NOT NULL CONSTRAINT [DF_MDBHistory_Active] DEFAULT ((1)),
    [CodMen] int NULL,
    CONSTRAINT [PK_MDBHistory] PRIMARY KEY CLUSTERED ([ITEM])
);
GO
CREATE TABLE dbo.[Mdbmensajeros] (
    [CodMen] int IDENTITY(1,1) NOT NULL,
    [NomMen] nvarchar(150) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_Mdbmensajeros_Active] DEFAULT ((1)),
    CONSTRAINT [PK_Mdbmensajeros] PRIMARY KEY CLUSTERED ([CodMen])
);
GO
CREATE TABLE dbo.[MDBProduct] (
    [CodPro] int IDENTITY(1,1) NOT NULL,
    [DesPro] nvarchar(150) NOT NULL,
    [Active] bit NOT NULL CONSTRAINT [DF_MDBProduct_Active] DEFAULT ((1)),
    CONSTRAINT [PK_MDBProduct] PRIMARY KEY CLUSTERED ([CodPro])
);
GO
CREATE TABLE dbo.[MDBServices] (
    [CodSer] numeric(18,0) NOT NULL,
    [NroBol] numeric(18,0) NOT NULL,
    [CodCLi] numeric(18,0) NOT NULL,
    [TipSer] int NOT NULL,
    [CanSer] numeric(18,0) NOT NULL,
    [Tipo] int NOT NULL,
    [Estado] bit NOT NULL CONSTRAINT [DF_MDBServices_Estado] DEFAULT ((1)),
    [Grupo] int NULL,
    CONSTRAINT [PK_MDBServices] PRIMARY KEY CLUSTERED ([CodSer])
);
GO
CREATE TABLE dbo.[MdbUsers] (
    [ITEM] int IDENTITY(1,1) NOT NULL,
    [NomUsu] nvarchar(50) NOT NULL,
    [ApeUsu] nvarchar(50) NOT NULL,
    [Usuari] nvarchar(5) NOT NULL,
    [Paswor] nvarchar(150) NOT NULL,
    [ActPas] bit NOT NULL CONSTRAINT [DF_MdbUsers_ActPas] DEFAULT ((1)),
    [Access] tinyint NOT NULL,
    [SitUsu] bit NOT NULL CONSTRAINT [DF_MdbUsers_SitUsu] DEFAULT ((1)),
    CONSTRAINT [PK_MdbUsers] PRIMARY KEY CLUSTERED ([ITEM])
);
GO
ALTER TABLE dbo.[ADMUsers] ADD CONSTRAINT [UQ_ADMUsers_Usuari] UNIQUE ([Usuari]);
ALTER TABLE dbo.[MdbUsers] ADD CONSTRAINT [UQ_MdbUsers_Usuari] UNIQUE ([Usuari]);
ALTER TABLE dbo.[IntBackup] ADD CONSTRAINT [UQ_IntBackup_Archivo] UNIQUE ([Archivo]);
ALTER TABLE dbo.[IntCorrel] ADD CONSTRAINT [UQ_IntCorrel_DivCor] UNIQUE ([DivCor]);
ALTER TABLE dbo.[IntInfoSys] ADD CONSTRAINT [UQ_IntInfoSys_VarMod] UNIQUE ([VarMod]);
ALTER TABLE dbo.[IntStatus] ADD CONSTRAINT [UQ_IntStatus_Status] UNIQUE ([Status]);
ALTER TABLE dbo.[IntTipDocument] ADD CONSTRAINT [UQ_IntTipDocument_DesTip] UNIQUE ([DesTip]);
ALTER TABLE dbo.[IntTipInvoice] ADD CONSTRAINT [UQ_IntTipInvoice_DesTip] UNIQUE ([DesTip]);
ALTER TABLE dbo.[IntTipService] ADD CONSTRAINT [UQ_IntTipService_DesTip] UNIQUE ([DesTip]);
ALTER TABLE dbo.[MDBDistritos] ADD CONSTRAINT [UQ_MDBDistritos_CodPos] UNIQUE ([CodPos]);
ALTER TABLE dbo.[MDBDistritos] ADD CONSTRAINT [UQ_MDBDistritos_NomDis] UNIQUE ([NomDis]);
ALTER TABLE dbo.[MdbGruposServ] ADD CONSTRAINT [UQ_MdbGruposServ_NomGrp] UNIQUE ([NomGrp]);
ALTER TABLE dbo.[Mdbmensajeros] ADD CONSTRAINT [UQ_Mdbmensajeros_NomMen] UNIQUE ([NomMen]);
ALTER TABLE dbo.[MDBProduct] ADD CONSTRAINT [UQ_MDBProduct_DesPro] UNIQUE ([DesPro]);
ALTER TABLE dbo.[MDBServices] ADD CONSTRAINT [UQ_MDBServices_NroBol] UNIQUE ([NroBol]);
ALTER TABLE dbo.[MDBCostumer] WITH CHECK ADD CONSTRAINT [FK_MDBCostumer_TipDoc] FOREIGN KEY ([TipDoc]) REFERENCES dbo.[IntTipDocument] ([CodTip]);
CREATE INDEX [IX_MDBCostumer_TipDoc] ON dbo.[MDBCostumer] ([TipDoc]);
ALTER TABLE dbo.[MDBCostumer] WITH CHECK ADD CONSTRAINT [FK_MDBCostumer_DisCli] FOREIGN KEY ([DisCli]) REFERENCES dbo.[MDBDistritos] ([CoDist]);
CREATE INDEX [IX_MDBCostumer_DisCli] ON dbo.[MDBCostumer] ([DisCli]);
ALTER TABLE dbo.[MDBHisEmp] WITH CHECK ADD CONSTRAINT [FK_MDBHisEmp_DisEmp] FOREIGN KEY ([DisEmp]) REFERENCES dbo.[MDBDistritos] ([CoDist]);
CREATE INDEX [IX_MDBHisEmp_DisEmp] ON dbo.[MDBHisEmp] ([DisEmp]);
ALTER TABLE dbo.[MDBServices] WITH CHECK ADD CONSTRAINT [FK_MDBServices_CodCLi] FOREIGN KEY ([CodCLi]) REFERENCES dbo.[MDBCostumer] ([CodCli]);
CREATE INDEX [IX_MDBServices_CodCLi] ON dbo.[MDBServices] ([CodCLi]);
ALTER TABLE dbo.[MDBServices] WITH CHECK ADD CONSTRAINT [FK_MDBServices_TipSer] FOREIGN KEY ([TipSer]) REFERENCES dbo.[IntTipService] ([CodTip]);
CREATE INDEX [IX_MDBServices_TipSer] ON dbo.[MDBServices] ([TipSer]);
ALTER TABLE dbo.[MDBServices] WITH CHECK ADD CONSTRAINT [FK_MDBServices_Tipo] FOREIGN KEY ([Tipo]) REFERENCES dbo.[MDBProduct] ([CodPro]);
CREATE INDEX [IX_MDBServices_Tipo] ON dbo.[MDBServices] ([Tipo]);
ALTER TABLE dbo.[MDBServices] WITH CHECK ADD CONSTRAINT [FK_MDBServices_Grupo] FOREIGN KEY ([Grupo]) REFERENCES dbo.[MdbGruposServ] ([Item]);
CREATE INDEX [IX_MDBServices_Grupo] ON dbo.[MDBServices] ([Grupo]);
ALTER TABLE dbo.[MDBDocument] WITH CHECK ADD CONSTRAINT [FK_MDBDocument_CoServ] FOREIGN KEY ([CoServ]) REFERENCES dbo.[MDBServices] ([CodSer]);
CREATE INDEX [IX_MDBDocument_CoServ] ON dbo.[MDBDocument] ([CoServ]);
ALTER TABLE dbo.[MDBDocument] WITH CHECK ADD CONSTRAINT [FK_MDBDocument_CoProd] FOREIGN KEY ([CoProd]) REFERENCES dbo.[MDBProduct] ([CodPro]);
CREATE INDEX [IX_MDBDocument_CoProd] ON dbo.[MDBDocument] ([CoProd]);
ALTER TABLE dbo.[MDBDocument] WITH CHECK ADD CONSTRAINT [FK_MDBDocument_CodEmp] FOREIGN KEY ([CodEmp]) REFERENCES dbo.[MDBHisEmp] ([CodEmp]);
CREATE INDEX [IX_MDBDocument_CodEmp] ON dbo.[MDBDocument] ([CodEmp]);
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_NroCod] FOREIGN KEY ([NroCod]) REFERENCES dbo.[MDBDocument] ([NroCod]);
CREATE INDEX [IX_MDBHistory_NroCod] ON dbo.[MDBHistory] ([NroCod]);
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_Status] FOREIGN KEY ([Status]) REFERENCES dbo.[IntStatus] ([ITEM]);
CREATE INDEX [IX_MDBHistory_Status] ON dbo.[MDBHistory] ([Status]);
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_Usuari] FOREIGN KEY ([Usuari]) REFERENCES dbo.[MdbUsers] ([ITEM]);
CREATE INDEX [IX_MDBHistory_Usuari] ON dbo.[MDBHistory] ([Usuari]);
ALTER TABLE dbo.[MDBHistory] WITH CHECK ADD CONSTRAINT [FK_MDBHistory_CodMen] FOREIGN KEY ([CodMen]) REFERENCES dbo.[Mdbmensajeros] ([CodMen]);
CREATE INDEX [IX_MDBHistory_CodMen] ON dbo.[MDBHistory] ([CodMen]);
ALTER TABLE dbo.[IntTarifasCourier] WITH CHECK ADD CONSTRAINT [FK_IntTarifasCourier_Distri] FOREIGN KEY ([Distri]) REFERENCES dbo.[MDBDistritos] ([CoDist]);
CREATE INDEX [IX_IntTarifasCourier_Distri] ON dbo.[IntTarifasCourier] ([Distri]);
CREATE INDEX [IX_MDBHistory_UltimoEstado] ON dbo.MDBHistory (NroCod,Active,FecEmi DESC,ITEM DESC) INCLUDE (Status,Usuari,Ranche,CodMen,Informe);
CREATE INDEX [IX_MDBHistory_Ranche] ON dbo.MDBHistory (Ranche,Active) INCLUDE (NroCod,CodMen);
CREATE INDEX [IX_MDBHisEmp_NomEmp] ON dbo.MDBHisEmp (NomEmp);
CREATE INDEX [IX_IntSaveActualizar_Pendientes] ON dbo.IntSaveActualizar (VersionEsquema,ActSav,Item);
ALTER TABLE dbo.[IntCorrel] ADD CONSTRAINT [CK_IntCorrel_Valores] CHECK (Correl>=0);
ALTER TABLE dbo.[MDBCostumer] ADD CONSTRAINT [CK_MDBCostumer_Valores] CHECK (CodCli>0 AND LEN(LTRIM(RTRIM(RucDni)))>0);
ALTER TABLE dbo.[MDBServices] ADD CONSTRAINT [CK_MDBServices_Valores] CHECK (CodSer>0 AND NroBol>0 AND CanSer>=0);
ALTER TABLE dbo.[MDBDocument] ADD CONSTRAINT [CK_MDBDocument_Valores] CHECK (NroCod>0);
ALTER TABLE dbo.[MDBHistory] ADD CONSTRAINT [CK_MDBHistory_Valores] CHECK (Ranche>=0);
ALTER TABLE dbo.[IntSaveActualizar] ADD CONSTRAINT [CK_IntSaveActualizar_Valores] CHECK (VersionEsquema IN(1,2));
ALTER TABLE dbo.[IntTarifasCourier] ADD CONSTRAINT [CK_IntTarifasCourier_Valores] CHECK (TimExp>=0 AND TimHot>=0 AND Tim1Di>=0 AND Tim2Di>=0 AND Tim3Di>=0 AND Tim4Di>=0);
ALTER TABLE dbo.MDBCostumer ADD CONSTRAINT UQ_MDBCostumer_Documento UNIQUE(TipDoc,RucDni);
GO
