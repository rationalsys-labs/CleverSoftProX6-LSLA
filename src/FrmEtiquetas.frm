VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Registros y etiquetas"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   13305
   Begin VB.TextBox BodyPar 
      Height          =   3300
      Left            =   7515
      MultiLine       =   -1  'True
      TabIndex        =   91
      Text            =   "FrmEtiquetas.frx":0000
      Top             =   3015
      Visible         =   0   'False
      Width           =   4650
   End
   Begin VB.TextBox Body 
      Height          =   3300
      Left            =   2700
      MultiLine       =   -1  'True
      TabIndex        =   90
      Text            =   "FrmEtiquetas.frx":06AA
      Top             =   2970
      Visible         =   0   'False
      Width           =   4650
   End
   Begin VB.Frame FraInfServ 
      Caption         =   "Servicio"
      Height          =   3615
      Left            =   90
      TabIndex        =   2
      Top             =   360
      Width           =   2220
      Begin VB.CheckBox ChkGen 
         Caption         =   "Usar Servicio General"
         Height          =   195
         Left            =   90
         TabIndex        =   70
         Top             =   945
         Width           =   1995
      End
      Begin VB.CommandButton CmdBusServ 
         Caption         =   "Buscar Serv"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   18
         Top             =   540
         Width           =   1020
      End
      Begin VB.CommandButton CmdNewServ 
         Caption         =   "Agregar Nuevo Servicio"
         Height          =   330
         Left            =   90
         TabIndex        =   17
         Top             =   180
         Width           =   2010
      End
      Begin VB.CommandButton CmdOkCan 
         Caption         =   "OK"
         Height          =   285
         Left            =   1665
         TabIndex        =   14
         Top             =   1620
         Width           =   435
      End
      Begin VB.TextBox TxtCan 
         Height          =   285
         Left            =   990
         TabIndex        =   13
         Top             =   1620
         Width           =   645
      End
      Begin VB.ComboBox CmbServPro 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   2520
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label LblCodCli 
         BackColor       =   &H00808080&
         Height          =   240
         Left            =   900
         TabIndex        =   42
         Top             =   2880
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label LblServ 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   2295
         Width           =   810
      End
      Begin VB.Label LblProd 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   2565
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label LblCod 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   1620
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label LblCli 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   3105
         Width           =   1950
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boleta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   1260
         Width           =   585
      End
      Begin VB.Label LblBol 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   3
         Top             =   1260
         Width           =   1095
      End
   End
   Begin VB.CheckBox ChkImp 
      Caption         =   "Auto Imprimir Etiqueta"
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   4050
      Value           =   1  'Checked
      Width           =   2220
   End
   Begin VB.Frame FraInf 
      Caption         =   "Informaciones"
      Height          =   1050
      Left            =   90
      TabIndex        =   12
      Top             =   4320
      Width           =   2220
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   65
         Top             =   630
         Width           =   795
      End
      Begin VB.Label lblCantidad 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   64
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label LblCodice 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   63
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Registro:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   62
         Top             =   270
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar BarHer 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   635
      ButtonWidth     =   3016
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ICO"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Crear Registros    "
            Key             =   "Crear"
            ImageKey        =   "a4"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anular Registros  "
            Key             =   "Anular"
            ImageKey        =   "a5"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Finalizar Registros"
            Key             =   "Finalizar"
            ImageKey        =   "a6"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Regenerar Masivo"
            Key             =   "Regenerar"
            ImageKey        =   "a7"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   6735
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3704
            MinWidth        =   3704
            TextSave        =   "31/05/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraImp 
      Caption         =   "Re-Imprimir Etiquetas"
      Height          =   1275
      Left            =   90
      TabIndex        =   66
      Top             =   5355
      Width           =   2220
      Begin VB.CommandButton CmdReImp 
         Caption         =   "Imprimir Etiquetas"
         Enabled         =   0   'False
         Height          =   240
         Left            =   135
         TabIndex        =   68
         Top             =   900
         Width           =   1950
      End
      Begin VB.ListBox LstReg 
         Enabled         =   0   'False
         Height          =   645
         Left            =   135
         TabIndex        =   67
         Top             =   225
         Width           =   1950
      End
   End
   Begin MSComctlLib.ImageList ICO 
      Left            =   2745
      Top             =   6210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":12A2
            Key             =   "A1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":183C
            Key             =   "A2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":1DD6
            Key             =   "a3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":23A6
            Key             =   "a4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":2DB8
            Key             =   "a5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":37CA
            Key             =   "a6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas.frx":3D64
            Key             =   "a7"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraNew 
      Caption         =   "Ingresar Nuevo Registro"
      Height          =   2265
      Left            =   2430
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   10770
      Begin VB.CommandButton CmdCancelarReg 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8460
         TabIndex        =   97
         Top             =   855
         Width           =   1995
      End
      Begin VB.ComboBox CmbModalNew 
         Height          =   315
         ItemData        =   "FrmEtiquetas.frx":4776
         Left            =   6525
         List            =   "FrmEtiquetas.frx":4780
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   315
         Width           =   1185
      End
      Begin VB.TextBox TxtDescrip 
         Height          =   285
         Left            =   1080
         TabIndex        =   85
         Top             =   675
         Width           =   2760
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Importar Datos de Excel"
         Height          =   375
         Left            =   8460
         TabIndex        =   69
         Top             =   1305
         Width           =   1995
      End
      Begin VB.CommandButton CmdAgreNew 
         Caption         =   "Agregar Registro"
         Height          =   375
         Left            =   8460
         TabIndex        =   59
         Top             =   1710
         Width           =   1995
      End
      Begin VB.TextBox TxtNewDat 
         Height          =   285
         Left            =   5670
         TabIndex        =   53
         Top             =   675
         Width           =   1995
      End
      Begin VB.CheckBox ChkNewDat 
         Caption         =   "Sin Data del Cliente"
         Height          =   195
         Left            =   3960
         TabIndex        =   61
         Top             =   720
         Width           =   1725
      End
      Begin VB.TextBox TxtNewEmp 
         Height          =   285
         Left            =   1080
         TabIndex        =   54
         Top             =   1035
         Width           =   4695
      End
      Begin VB.TextBox TxtNewDir 
         Height          =   285
         Left            =   1080
         TabIndex        =   55
         Top             =   1395
         Width           =   3885
      End
      Begin VB.ComboBox CmbNewDis 
         Height          =   315
         Left            =   5850
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   1395
         Width           =   1815
      End
      Begin VB.TextBox txtNewCont 
         Height          =   285
         Left            =   1080
         TabIndex        =   58
         Top             =   1755
         Width           =   6585
      End
      Begin VB.Label Label30 
         Caption         =   "Modal:"
         Height          =   240
         Left            =   5985
         TabIndex        =   95
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion:"
         Height          =   195
         Left            =   180
         TabIndex        =   84
         Top             =   675
         Width           =   885
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   180
         TabIndex        =   60
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Labetitlecod 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   7920
         TabIndex        =   57
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Empresas:"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   180
         TabIndex        =   51
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Distrito:"
         Height          =   195
         Left            =   5130
         TabIndex        =   50
         Top             =   1395
         Width           =   525
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Contacto:"
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label LblNewCod 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8640
         TabIndex        =   48
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   3870
         TabIndex        =   47
         Top             =   360
         Width           =   690
      End
      Begin VB.Label LblNewCli 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   46
         Top             =   315
         Width           =   2760
      End
      Begin VB.Label LblNewPro 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4590
         TabIndex        =   45
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label LblCodEmp 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5850
         TabIndex        =   44
         Top             =   1035
         Width           =   1815
      End
   End
   Begin VB.Frame FraServ 
      Caption         =   "Creacion de Servicios"
      Height          =   2265
      Left            =   2430
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   10770
      Begin VB.CommandButton CmdCancelarServ 
         Caption         =   "Cancelar"
         Height          =   285
         Left            =   8730
         TabIndex        =   98
         Top             =   1485
         Width           =   1770
      End
      Begin VB.CommandButton CmdGuaServ 
         Caption         =   "Guardar"
         Height          =   285
         Left            =   8730
         TabIndex        =   41
         Top             =   1845
         Width           =   1770
      End
      Begin VB.TextBox TxtCli 
         Height          =   285
         Left            =   1035
         MaxLength       =   300
         TabIndex        =   25
         Top             =   585
         Width           =   5040
      End
      Begin VB.Label LblAyudaCliente
         Caption         =   "F9 para busqueda"
         Height          =   285
         Left            =   6165
         TabIndex        =   99
         Top             =   585
         Width           =   1590
      End
      Begin VB.ComboBox CmbServ 
         Height          =   315
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1215
         Width           =   2670
      End
      Begin VB.TextBox TxtCanServ 
         Height          =   285
         Left            =   1035
         MaxLength       =   6
         TabIndex        =   23
         Top             =   1845
         Width           =   1455
      End
      Begin VB.ComboBox CmbProd 
         Height          =   315
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1575
         Width           =   2670
      End
      Begin VB.TextBox TxtBol 
         Height          =   285
         Left            =   1035
         MaxLength       =   18
         TabIndex        =   21
         Top             =   1530
         Width           =   2400
      End
      Begin VB.Label LblCodServ 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8685
         TabIndex        =   40
         Top             =   585
         Width           =   1815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   8100
         TabIndex        =   39
         Top             =   630
         Width           =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   135
         TabIndex        =   38
         Top             =   630
         Width           =   525
      End
      Begin VB.Label LblDis 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8685
         TabIndex        =   37
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Distrito:"
         Height          =   195
         Left            =   8100
         TabIndex        =   36
         Top             =   945
         Width           =   525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   135
         TabIndex        =   35
         Top             =   945
         Width           =   720
      End
      Begin VB.Label LblDir 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1035
         TabIndex        =   34
         Top             =   900
         Width           =   6720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "RUC:"
         Height          =   195
         Left            =   135
         TabIndex        =   33
         Top             =   1260
         Width           =   390
      End
      Begin VB.Label LblRuc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1035
         TabIndex        =   32
         Top             =   1215
         Width           =   2400
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Servicio:"
         Height          =   195
         Left            =   4275
         TabIndex        =   31
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   1890
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   4275
         TabIndex        =   29
         Top             =   1620
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Boleta:"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   1575
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Serie de Correlativo Interno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   270
         Width           =   3375
      End
      Begin VB.Label Codice 
         AutoSize        =   -1  'True
         Caption         =   "0000000001"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3645
         TabIndex        =   26
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.Frame FraimpExcel 
      Height          =   2265
      Left            =   2430
      TabIndex        =   72
      Top             =   360
      Visible         =   0   'False
      Width           =   10770
      Begin VB.ComboBox CmbModal 
         Height          =   315
         ItemData        =   "FrmEtiquetas.frx":47A7
         Left            =   3510
         List            =   "FrmEtiquetas.frx":47B1
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   225
         Width           =   1725
      End
      Begin VB.TextBox Txtdes 
         Height          =   285
         Left            =   6255
         TabIndex        =   86
         Top             =   225
         Width           =   1455
      End
      Begin VB.ListBox LstCodDis 
         Height          =   450
         Left            =   9585
         TabIndex        =   80
         Top             =   90
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.ListBox LstDis 
         Height          =   450
         Left            =   9000
         TabIndex        =   79
         Top             =   90
         Visible         =   0   'False
         Width           =   600
      End
      Begin MSComDlg.CommonDialog CUAD 
         Left            =   90
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView ViewErr 
         Height          =   1140
         Left            =   180
         TabIndex        =   78
         Top             =   945
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   2011
         _Version        =   393217
         Indentation     =   19
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton CmdCargExcel 
         Caption         =   "Cargar datos"
         Height          =   285
         Left            =   9270
         TabIndex        =   77
         Top             =   585
         Width           =   1365
      End
      Begin VB.CommandButton CmdExa 
         Caption         =   "Examinar"
         Height          =   285
         Left            =   7875
         TabIndex        =   76
         Top             =   585
         Width           =   1140
      End
      Begin VB.TextBox txtexaminar 
         Height          =   285
         Left            =   180
         TabIndex        =   75
         Top             =   585
         Width           =   7530
      End
      Begin VB.CommandButton CmdX 
         Caption         =   "X"
         Height          =   285
         Left            =   10260
         TabIndex        =   73
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label28 
         Caption         =   "Modal:"
         Height          =   240
         Left            =   2970
         TabIndex        =   89
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label26 
         Caption         =   "Descripcion:"
         Height          =   240
         Left            =   5355
         TabIndex        =   87
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Importar Informaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   74
         Top             =   225
         Width           =   2040
      End
   End
   Begin VB.Frame FraRe 
      Caption         =   "Re-Imprimir Registros Masivos"
      Height          =   780
      Left            =   2430
      TabIndex        =   81
      Top             =   360
      Visible         =   0   'False
      Width           =   10770
      Begin VB.CommandButton CmdVolverRegen
         Caption         =   "Volver"
         Height          =   315
         Left            =   8730
         TabIndex        =   100
         Top             =   270
         Width           =   1770
      End
      Begin VB.ComboBox CmbModalReImp 
         Height          =   315
         ItemData        =   "FrmEtiquetas.frx":47D8
         Left            =   3600
         List            =   "FrmEtiquetas.frx":47E2
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   270
         Width           =   1725
      End
      Begin VB.TextBox txtRegenerarMasivo 
         Height          =   285
         Left            =   990
         TabIndex        =   82
         Top             =   315
         Width           =   1860
      End
      Begin VB.Label Label29 
         Caption         =   "Modal:"
         Height          =   240
         Left            =   3060
         TabIndex        =   93
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label19 
         Caption         =   "Registro"
         Height          =   240
         Left            =   270
         TabIndex        =   83
         Top             =   315
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5460
      Left            =   2430
      TabIndex        =   0
      Top             =   1170
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   9631
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "FrmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cantidad As Long
Private CodigoProducto As Long
Private PanelActual As Integer
Private Cargando As Boolean
Private Completando As Boolean
Private OmitirCompletar As Boolean
Private Importando As Boolean
Private CargaTerminada As Boolean
Private SeleccionandoCliente As Boolean
Private Guardando As Boolean
Private ConsultaPendiente As Boolean
Private ServicioAnteriorReimpresion As String

Private Sub Form_Load()
    Dim Columna As Integer
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrInt
    Dim Panel As MSComctlLib.Panel
    For Each Panel In StatusBar1.Panels
        Panel.Style = sbrText: Panel.MinWidth = 15
    Next
    StatusBar1.Panels(1).Width = 2700
    StatusBar1.Panels(2).AutoSize = sbrSpring
    StatusBar1.Panels(3).Width = 2400
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    ViewErr.ImageList = Me.ICO
    Grid.Rows = 1: Grid.Cols = 10
    For Columna = 0 To 9: Grid.ColWidth(Columna) = 0: Next
    Grid.TextMatrix(0, 0) = "Codigo": Grid.ColWidth(0) = 1100
    Grid.TextMatrix(0, 1) = "Cliente": Grid.TextMatrix(0, 2) = "CodCliente"
    Grid.TextMatrix(0, 3) = "Producto": Grid.ColWidth(3) = 1000
    Grid.TextMatrix(0, 4) = "Registro"
    Grid.TextMatrix(0, 5) = "Destinatario": Grid.ColWidth(5) = 4500
    Grid.TextMatrix(0, 6) = "Direccion": Grid.ColWidth(6) = 3700
    Grid.TextMatrix(0, 7) = "Distrito": Grid.TextMatrix(0, 8) = "Contacto": Grid.TextMatrix(0, 9) = "Fecha"
    CargarLista CmbServ, "SELECT CodTip,DesTip FROM dbo.IntTipService WHERE Active=1 ORDER BY CodTip"
    CargarLista CmbProd, "SELECT CodPro,DesPro FROM dbo.MDBProduct WHERE Active=1 ORDER BY CodPro"
    CargarLista CmbServPro, "SELECT CodPro,DesPro FROM dbo.MDBProduct WHERE Active=1 ORDER BY CodPro"
    CargarLista CmbNewDis, "SELECT CoDist,NomDis FROM dbo.MDBDistritos WHERE Active=1 ORDER BY NomDis"
    CmbModal.ListIndex = 0: CmbModalNew.ListIndex = 0: CmbModalReImp.ListIndex = 0
    TxtNewDat.MaxLength = 50: TxtNewEmp.MaxLength = 500: TxtNewDir.MaxLength = 500
    txtNewCont.MaxLength = 150: TxtDescrip.MaxLength = 150: Txtdes.MaxLength = 150
    txtexaminar.Locked = True
    CambiarPanel 0
    LimpiarServicio
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub ActualizarEstado()
    Dim ServicioSeleccionado As Boolean, Disponible As Boolean, General As Boolean
    Dim RegistroSeleccionado As Boolean, Registrados As Long
    Dim Mensaje As String
    ServicioSeleccionado = CodigoValido(lblServ.Caption)
    Disponible = (PanelActual = 0 And Not Guardando And Not Importando)
    General = (lblServ.Caption = "105")
    Registrados = Grid.Rows - 1
    If Grid.Row > 0 And Grid.Row < Grid.Rows Then RegistroSeleccionado = CodigoValido(Grid.TextMatrix(Grid.Row, 0))
    BarHer.Enabled = Disponible
    BarHer.Buttons("Crear").Enabled = (Disponible And ServicioSeleccionado And CodigoProducto > 0 And Not ConsultaPendiente And (General Or Registrados < Cantidad))
    BarHer.Buttons("Anular").Enabled = (Disponible And ServicioSeleccionado And RegistroSeleccionado And Not ConsultaPendiente)
    BarHer.Buttons("Finalizar").Enabled = (Disponible And ServicioSeleccionado And Not General And Not ConsultaPendiente And Registrados > 0 And Registrados = Cantidad)
    'La reimpresion consulta documentos por codigo; no necesita un servicio cargado.
    BarHer.Buttons("Regenerar").Enabled = Disponible
    TxtCan.Enabled = (Disponible And ServicioSeleccionado And Not General)
    CmdOkCan.Enabled = (TxtCan.Enabled And CodigoValido(TxtCan.Text) And TxtCan.Text <> CStr(Cantidad))
    CmdCancelarServ.Enabled = (Not Guardando And Not Importando)
    CmdCancelarReg.Enabled = (Not Guardando And Not Importando)
    CmdVolverRegen.Enabled = (Not Guardando And Not Importando)
    Select Case PanelActual
        Case 1: Mensaje = "Ingresando documento"
        Case 2: Mensaje = "Creando servicio"
        Case 3: Mensaje = "Reimpresion por codigo"
        Case 4
            If Importando Then
                Mensaje = "Importando Excel"
            ElseIf CargaTerminada Then
                Mensaje = "Importacion guardada"
            Else
                Mensaje = "Importacion Excel"
            End If
        Case Else
            If Not ServicioSeleccionado Then
                Mensaje = "Seleccione un servicio"
            ElseIf ConsultaPendiente Then
                Mensaje = "Consulta pendiente"
            ElseIf TxtCan.Text <> CStr(Cantidad) Then
                Mensaje = "Cantidad sin confirmar"
            Else
                Mensaje = "Consulta de servicio"
            End If
    End Select
    StatusBar1.Panels(1).Text = Mensaje
    If PanelActual = 3 Then
        StatusBar1.Panels(2).Text = "Ingrese el codigo del documento. Volver para salir."
        StatusBar1.Panels(3).Text = "Consultados: " & CStr(Registrados)
    ElseIf ServicioSeleccionado Then
        StatusBar1.Panels(2).Text = "Servicio " & lblServ.Caption & " - " & LblCli.Caption
        If ConsultaPendiente Then
            StatusBar1.Panels(3).Text = "Requiere consulta"
        ElseIf General Then
            StatusBar1.Panels(3).Text = "Registrados: " & CStr(Registrados)
        Else
            StatusBar1.Panels(3).Text = "Registrados: " & CStr(Registrados) & " de " & CStr(Cantidad)
        End If
    Else
        StatusBar1.Panels(2).Text = "Busque un servicio o agregue uno nuevo"
        StatusBar1.Panels(3).Text = "Sin servicio"
    End If
End Sub

Private Sub Grid_RowColChange()
    If Not Cargando Then ActualizarEstado
End Sub

Private Sub CmdVolverRegen_Click()
    On Error GoTo ErrInt
    CancelarIngreso
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CambiarPanel(ByVal Panel As Integer)
    PanelActual = Panel
    FraNew.Visible = (Panel = 1): FraServ.Visible = (Panel = 2)
    FraRe.Visible = (Panel = 3): FraimpExcel.Visible = (Panel = 4)
    Grid.Enabled = (Panel = 0 Or Panel = 3)
    FraInfServ.Enabled = (Panel = 0): FraInf.Enabled = (Panel = 0)
    FraImp.Enabled = (Panel = 0): ChkImp.Enabled = (Panel = 0): BarHer.Enabled = (Panel = 0)
    Grid.Top = 450: Grid.Height = 6180
    If Panel = 1 Or Panel = 2 Or Panel = 4 Then Grid.Top = 2655: Grid.Height = 3975
    If Panel = 3 Then Grid.Top = 1170: Grid.Height = 5460
    ActualizarEstado
End Sub

Private Sub LimpiarServicio()
    Cargando = True
    LimpiarDocumento
    ConsultaPendiente = False
    TxtCan.FontBold = False: TxtCan.ForeColor = vbBlack
    lblServ.Caption = "": LblBol.Caption = "": LblCli.Caption = "": LblCod.Caption = "": LblProd.Caption = ""
    TxtCan.Text = "0": Cantidad = 0: CodigoProducto = 0: ChkGen.Value = 0
    CmbServPro.Visible = False: CmdOkCan.Enabled = True: CmdNewServ.Enabled = True: CmdBusServ.Enabled = True
    BarHer.Buttons(3).Enabled = True
    Grid.Rows = 1: lblCantidad.Caption = "0": LstReg.Clear: CmdReImp.Enabled = False
    Cargando = False
    ActualizarEstado
End Sub

Public Sub CargarServicio(ByVal Codigo As String)
    Dim Datos As ADODB.Recordset
    If Not CodigoValido(Codigo) Then Err.Raise vbObjectError + 340, , "Codigo de servicio no valido."
    Set Datos = ConsultarDatos("SELECT CodSer,NroBol,NomCli,CodCLi,Tipo,CanSer,Estado FROM dbo.ViewDatosServicios WHERE CodSer=?", Codigo, True)
    If Datos.EOF Then Datos.Close: Err.Raise vbObjectError + 341, , "No existe el servicio."
    If Not CBool(Datos.Fields(6).Value) Then Datos.Close: Err.Raise vbObjectError + 342, , "El servicio esta finalizado. Debe reactivarlo desde la consulta de documentos."
    Cargando = True
    LimpiarDocumento
    lblServ.Caption = CStr(Datos.Fields(0).Value): LblBol.Caption = CStr(Datos.Fields(1).Value)
    LblCli.Caption = TextoDato(Datos.Fields(2).Value): LblCod.Caption = CStr(Datos.Fields(3).Value)
    CodigoProducto = CLng(Datos.Fields(4).Value): Cantidad = CLng(Datos.Fields(5).Value)
    TxtCan.Text = CStr(Cantidad): TxtCan.FontBold = False: TxtCan.ForeColor = vbBlack
    ChkGen.Value = 0
    If CDec(Codigo) = 105 Then ChkGen.Value = 1
    CmbServPro.Visible = (ChkGen.Value = 1): CmbServPro.ListIndex = -1
    CmdOkCan.Enabled = (ChkGen.Value = 0): CmdBusServ.Enabled = (ChkGen.Value = 0): CmdNewServ.Enabled = (ChkGen.Value = 0)
    BarHer.Buttons(3).Enabled = (ChkGen.Value = 0)
    Datos.Close
    Set Datos = ConsultarDatos("SELECT DesPro FROM dbo.MDBProduct WHERE CodPro=?", CStr(CodigoProducto), True)
    LblProd.Caption = TextoDato(Datos.Fields(0).Value): Datos.Close
    If ChkGen.Value = 1 Then LblProd.Caption = "": CodigoProducto = 0
    Cargando = False
    CambiarPanel 0
    RecargarRegistros
End Sub

Public Sub RecargarRegistros()
    Dim Datos As ADODB.Recordset
    Dim Columna As Integer, Fila As Long
    Set Datos = ConsultarDatos("SELECT Codigo,Cliente,CodCliente,Producto,Registro,Empresa,Direccion,Distrito,Contacto,Fecha FROM dbo.ViewRegistAct WHERE Servicio=? ORDER BY Codigo DESC", lblServ.Caption, True)
    Grid.Rows = 1: LstReg.Clear
    Do Until Datos.EOF
        Grid.Rows = Grid.Rows + 1: Fila = Grid.Rows - 1
        For Columna = 0 To 9: Grid.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
        LstReg.AddItem CStr(Datos.Fields(0).Value)
        Datos.MoveNext
    Loop
    lblCantidad.Caption = CStr(Grid.Rows - 1)
    LstReg.Enabled = True: CmdReImp.Enabled = (Grid.Rows > 1)
    Datos.Close
    ActualizarCorrelativo
    ConsultaPendiente = False
    ActualizarEstado
End Sub

Private Sub ActualizarCorrelativo()
    Dim Datos As ADODB.Recordset
    Set Datos = ConsultarDatos("EXEC dbo.Correlativo 'CDO'")
    LblCodice.Caption = CStr(Datos.Fields(0).Value): Datos.Close
End Sub

Private Sub LimpiarDocumento()
    Completando = True
    TxtNewDat.Text = "": TxtNewEmp.Text = "": TxtNewDir.Text = "": txtNewCont.Text = ""
    TxtDescrip.Text = "": LblCodEmp.Caption = "": CmbNewDis.ListIndex = -1
    ChkNewDat.Value = 0
    OmitirCompletar = False: Completando = False
End Sub

Private Sub RestaurarCantidad()
    Cargando = True
    TxtCan.Text = CStr(Cantidad): TxtCan.FontBold = False: TxtCan.ForeColor = vbBlack
    Cargando = False
    ActualizarEstado
End Sub

Private Function PuedeDescartarPendientes() As Boolean
    If Importando Or Guardando Then Exit Function
    PuedeDescartarPendientes = True
    If HayCamposPendientes() Then PuedeDescartarPendientes = (MsgBox("Hay campos o una cantidad sin guardar. Desea descartarlos?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Clever SoftPro") = vbYes)
End Function

Private Sub CancelarIngreso()
    If Not PuedeDescartarPendientes() Then Exit Sub
    If PanelActual = 3 Then
        txtRegenerarMasivo.Text = ""
        CambiarPanel 0
        If Len(ServicioAnteriorReimpresion) > 0 Then CargarServicio ServicioAnteriorReimpresion Else LimpiarServicio
        ServicioAnteriorReimpresion = ""
        Exit Sub
    End If
    LimpiarDocumento
    TxtCli.Text = "": TxtBol.Text = "": TxtCanServ.Text = ""
    CmbServ.ListIndex = -1: CmbProd.ListIndex = -1
    txtexaminar.Text = "": Txtdes.Text = "": CargaTerminada = False
    RestaurarCantidad
    CambiarPanel 0
End Sub

Private Sub CmdCancelarServ_Click()
    CancelarIngreso
End Sub

Private Sub CmdCancelarReg_Click()
    CancelarIngreso
End Sub

Private Sub TxtCli_Change()
    If SeleccionandoCliente Then Exit Sub
    LblCodServ.Caption = "": LblDir.Caption = "": LblDis.Caption = "": LblRuc.Caption = ""
End Sub

Private Sub ActualizarGuardados(ByVal Detalle As String)
    On Error GoTo ErrInt
    RecargarRegistros
    Exit Sub
ErrInt:
    ConsultaPendiente = True
    Grid.Rows = 1: LstReg.Clear: CmdReImp.Enabled = False
    CambiarPanel 0
    MsgBox Detalle & vbCrLf & "Los datos estan guardados, pero no se pudo actualizar la lista: " & Err.Description & vbCrLf & "No repita el ingreso. Vuelva a consultar el servicio para verlos o reimprimirlos.", vbExclamation, "Clever SoftPro"
End Sub

Private Sub PrepararDocumento()
    If CodigoProducto = 0 Or Not CodigoValido(lblServ.Caption) Then Err.Raise vbObjectError + 343, , "Seleccione un servicio y su producto."
    If ConsultaPendiente Then RecargarRegistros
    If ChkGen.Value = 0 And Grid.Rows - 1 >= Cantidad Then Err.Raise vbObjectError + 343, , "Se alcanzo la cantidad de documentos del servicio."
    LimpiarDocumento
    LblNewCli.Caption = LblCli.Caption: LblNewCod.Caption = LblCod.Caption: LblNewPro.Caption = LblProd.Caption
    CambiarPanel 1
End Sub

Public Function AgregarDocumentoActual() As String
    Dim Documento As DatosDocumento
    Dim Faltantes As String, Codigo As String
    Dim Numero As Long, Descripcion As String
    If Guardando Or Importando Then Err.Raise vbObjectError + 344, , "Hay un guardado en curso."
    If Not CodigoValido(lblServ.Caption) Or CodigoProducto = 0 Then Err.Raise vbObjectError + 344, , "Seleccione un servicio y su producto."
    If ChkNewDat.Value = 0 And Len(Trim$(TxtNewDat.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Referencia del cliente, o marque Sin Data del Cliente"
    If Len(Trim$(TxtNewEmp.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Destinatario"
    If Len(Trim$(TxtNewDir.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Direccion"
    If CmbNewDis.ListIndex < 0 Then Faltantes = Faltantes & vbCrLf & "- Distrito"
    If Len(Faltantes) > 0 Then Err.Raise vbObjectError + 344, , "Complete los campos obligatorios:" & Faltantes
    On Error GoTo ErrInt
    Guardando = True: CmdAgreNew.Enabled = False
    Documento.Servicio = lblServ.Caption: Documento.Cliente = LblCod.Caption: Documento.Producto = CodigoProducto
    Documento.Referencia = Trim$(TxtNewDat.Text)
    If ChkNewDat.Value = 1 Then Documento.Referencia = ""
    Documento.Destinatario = Trim$(TxtNewEmp.Text): Documento.Direccion = Trim$(TxtNewDir.Text)
    Documento.Distrito = CmbNewDis.ItemData(CmbNewDis.ListIndex): Documento.Contacto = Trim$(txtNewCont.Text)
    Documento.Descripcion = Trim$(TxtDescrip.Text): Documento.CodigoDestinatario = LblCodEmp.Caption
    Codigo = GuardarRegistro(Documento)
    'Limpiar el ingreso confirmado antes de cualquier consulta posterior.
    LimpiarDocumento
    AgregarDocumentoActual = Codigo
    ActualizarGuardados "Registro guardado: " & Codigo
    If Not ConsultaPendiente And ChkGen.Value = 0 And Grid.Rows - 1 >= Cantidad Then CambiarPanel 0
    Guardando = False: CmdAgreNew.Enabled = True
    ActualizarEstado
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    Guardando = False: CmdAgreNew.Enabled = True
    ActualizarEstado
    Err.Raise Numero, "AgregarDocumentoActual", Descripcion
End Function

Private Sub CmdAgreNew_Click()
    Dim Codigo As String, Codigos As Collection
    On Error GoTo ErrInt
    Codigo = AgregarDocumentoActual()
    If ChkImp.Value = 1 Then
        Set Codigos = New Collection: Codigos.Add Codigo
        ImprimirGuardados Codigos, CmbModalNew.Text
    End If
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CmdBusServ_Click()
    On Error GoTo ErrInt
    If Not PuedeDescartarPendientes() Then Exit Sub
    RestaurarCantidad
    Var.Varflotante = 1
    FrmMenuFlot.Top = Me.Top + 2350: FrmMenuFlot.Left = Me.Left + 260
    FrmMenuFlot.Show vbModal
    If Var.RegistroSeleccionado Then CargarServicio Var.ResultadoRegistro(0)
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CmdNewServ_Click()
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrInt
    If Not PuedeDescartarPendientes() Then Exit Sub
    RestaurarCantidad
    CmdGuaServ.Enabled = True
    TxtCli.Text = "": LblCodServ.Caption = "": LblDir.Caption = "": LblDis.Caption = "": LblRuc.Caption = ""
    TxtBol.Text = "": TxtCanServ.Text = "": CmbServ.ListIndex = -1: CmbProd.ListIndex = -1
    Set Datos = ConsultarDatos("EXEC dbo.Correlativo 'CSV'")
    Codice.Caption = CStr(Datos.Fields(0).Value): Datos.Close
    CambiarPanel 2
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub TxtCli_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrInt
    If KeyCode <> vbKeyF9 Then Exit Sub
    KeyCode = 0
    FrmBusqueda.VariableBusCli = TxtCli.Text
    FrmBusqueda.Show vbModal
    If Var.ResulBus(0) <> "" Then
        SeleccionandoCliente = True: TxtCli.Text = Var.ResulBus(1): LblCodServ.Caption = Var.ResulBus(0)
        LblRuc.Caption = Var.ResulBus(2): LblDir.Caption = Var.ResulBus(3): LblDis.Caption = Var.ResulBus(4)
        SeleccionandoCliente = False
    End If
    Exit Sub
ErrInt:
    SeleccionandoCliente = False
    InformarError
End Sub

Private Sub CmdGuaServ_Click()
    Dim Codigo As String, Faltantes As String, Descripcion As String
    On Error GoTo ErrInt
    If Guardando Or Not CmdGuaServ.Enabled Then Exit Sub
    If Not CodigoValido(LblCodServ.Caption) Then Faltantes = Faltantes & vbCrLf & "- Seleccione el cliente con F9"
    If Not CodigoValido(TxtBol.Text) Then Faltantes = Faltantes & vbCrLf & "- Boleta: entero positivo de hasta 18 digitos"
    If Not CodigoValido(TxtCanServ.Text) Then Faltantes = Faltantes & vbCrLf & "- Cantidad: entero positivo"
    If CmbServ.ListIndex < 0 Then Faltantes = Faltantes & vbCrLf & "- Tipo de servicio"
    If CmbProd.ListIndex < 0 Then Faltantes = Faltantes & vbCrLf & "- Producto"
    If Len(Faltantes) > 0 Then Err.Raise vbObjectError + 345, , "Complete los campos obligatorios:" & Faltantes
    Guardando = True: CmdGuaServ.Enabled = False
    Codigo = CrearServicio(LblCodServ.Caption, CmbServ.ItemData(CmbServ.ListIndex), TxtCanServ.Text, CmbProd.ItemData(CmbProd.ListIndex), TxtBol.Text)
    CambiarPanel 0
    CargarServicio Codigo
    Guardando = False
    ActualizarEstado
    Exit Sub
ErrInt:
    Guardando = False
    If Len(Codigo) > 0 Then
        Descripcion = Err.Description
        LimpiarServicio
        CambiarPanel 0
        MsgBox "El servicio " & Codigo & " esta guardado, pero no se pudo cargar su consulta: " & Descripcion & vbCrLf & "Busquelo por su codigo. No vuelva a crearlo.", vbExclamation, "Clever SoftPro"
    Else
        CmdGuaServ.Enabled = True
        InformarError
    End If
End Sub

Private Sub CmdOkCan_Click()
    On Error GoTo ErrInt
    If Not CodigoValido(lblServ.Caption) Or ChkGen.Value = 1 Then Exit Sub
    If Not CodigoValido(TxtCan.Text) Then Err.Raise vbObjectError + 346, , "La cantidad debe ser un entero positivo."
    CambiarServicio lblServ.Caption, 1, TxtCan.Text
    CargarServicio lblServ.Caption
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub BarHer_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrInt
    ActualizarEstado
    If Not BarHer.Enabled Or Not Button.Enabled Then Exit Sub
    If Not PuedeDescartarPendientes() Then Exit Sub
    RestaurarCantidad
    Select Case Button.Key
        Case "Crear": PrepararDocumento
        Case "Finalizar"
            If MsgBox("Desea finalizar el servicio " & lblServ.Caption & "?", vbQuestion Or vbYesNo, "Finalizar Servicio") = vbYes Then
                CambiarServicio lblServ.Caption, 2
                LimpiarServicio
            End If
        Case "Anular"
            If Grid.Row < 1 Then Exit Sub
            If MsgBox("Desea anular el documento seleccionado?", vbQuestion Or vbYesNo, "Anular Registro") = vbYes Then
                AnularRegistro Grid.TextMatrix(Grid.Row, 0)
                RecargarRegistros
            End If
        Case "Regenerar"
            ServicioAnteriorReimpresion = lblServ.Caption
            LimpiarServicio
            CambiarPanel 3
    End Select
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub ChkGen_Click()
    On Error GoTo ErrInt
    If Cargando Then Exit Sub
    If Not PuedeDescartarPendientes() Then
        Cargando = True
        If lblServ.Caption = "105" Then ChkGen.Value = 1 Else ChkGen.Value = 0
        Cargando = False
        Exit Sub
    End If
    If ChkGen.Value = 1 Then CargarServicio "105" Else LimpiarServicio
    Exit Sub
ErrInt:
    Cargando = False
    ChkGen.Value = 0
    InformarError
End Sub

Private Sub CmbServPro_Click()
    If CmbServPro.ListIndex < 0 Then Exit Sub
    CodigoProducto = CmbServPro.ItemData(CmbServPro.ListIndex): LblProd.Caption = CmbServPro.Text
    ActualizarEstado
End Sub

Private Sub ChkNewDat_Click()
    TxtNewDat.Enabled = (ChkNewDat.Value = 0)
    If ChkNewDat.Value = 1 Then TxtNewDat.Text = "": TxtNewDat.BackColor = &H8000000F Else TxtNewDat.BackColor = &H80000005
End Sub

Private Sub TxtCan_Change()
    If Cargando Then Exit Sub
    TxtCan.FontBold = True: TxtCan.ForeColor = vbRed
    ActualizarEstado
End Sub

Private Sub TxtCan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then RestaurarCantidad: KeyCode = 0
End Sub

Private Sub CmdImport_Click()
    If HayDocumentoPendiente() Then
        If MsgBox("Desea descartar el registro manual pendiente y abrir la importacion?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    End If
    LimpiarDocumento
    CargaTerminada = False: CmdCargExcel.Enabled = True
    txtexaminar.Text = "": Txtdes.Text = ""
    CambiarPanel 4
    ViewErr.Nodes.Clear
End Sub

Private Sub CmdX_Click()
    If Importando Or Guardando Then Exit Sub
    If Not CargaTerminada And Len(txtexaminar.Text & Txtdes.Text) > 0 Then
        If MsgBox("Desea descartar la importacion pendiente?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    End If
    txtexaminar.Text = "": Txtdes.Text = ""
    CambiarPanel 1
End Sub

Private Sub CmdExa_Click()
    On Error GoTo ErrInt
    If Importando Or Guardando Then Exit Sub
    CUAD.FileName = ""
    CUAD.CancelError = False: CUAD.DialogTitle = "Abrir"
    CUAD.Filter = "Archivos de Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
    CUAD.ShowOpen
    If Len(CUAD.FileName) > 0 Then
        txtexaminar.Text = CUAD.FileName
        CargaTerminada = False: CmdCargExcel.Enabled = True
    End If
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CmdCargExcel_Click()
    Dim Plantilla As DatosDocumento, Codigos As Collection
    On Error GoTo ErrInt
    If Importando Or CargaTerminada Then Exit Sub
    If CodigoProducto = 0 Then Err.Raise vbObjectError + 347, , "Seleccione el producto del servicio."
    Importando = True: CmdCargExcel.Enabled = False: Me.MousePointer = vbHourglass
    ActualizarEstado
    Plantilla.Servicio = lblServ.Caption: Plantilla.Cliente = LblCod.Caption: Plantilla.Producto = CodigoProducto
    Plantilla.Descripcion = Txtdes.Text
    ViewErr.Nodes.Clear
    ViewErr.Nodes.Add , , "Inicio", "Validando archivo de Excel..."
    Set Codigos = ImportarDocumentos(txtexaminar.Text, Plantilla)
    'El lote ya esta confirmado, aunque falle la consulta posterior.
    CargaTerminada = True
    ActualizarGuardados CStr(Codigos.Count) & " registros guardados."
    ViewErr.Nodes.Add , , "Fin", CStr(Codigos.Count) & " registros guardados correctamente."
    Importando = False: CargaTerminada = True: CmdCargExcel.Enabled = False: Me.MousePointer = vbDefault
    ActualizarEstado
    If ChkImp.Value = 1 Then ImprimirGuardados Codigos, CmbModal.Text
    Exit Sub
ErrInt:
    Importando = False: CmdCargExcel.Enabled = Not CargaTerminada: Me.MousePointer = vbDefault
    ActualizarEstado
    ViewErr.Nodes.Add , , , "Error: " & Err.Description
    InformarError
End Sub

Private Sub ImprimirGuardados(ByVal Codigos As Collection, ByVal Modalidad As String)
    On Error GoTo ErrInt
    ImprimirRegistros Codigos, Modalidad
    Exit Sub
ErrInt:
    MsgBox "Los registros estan guardados, pero no se pudo imprimir: " & Err.Description & vbCrLf & "Puede volver a imprimirlos desde Re-Imprimir Etiquetas.", vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdReImp_Click()
    Dim Codigos As Collection, Posicion As Long
    Set Codigos = New Collection
    If LstReg.ListIndex >= 0 Then
        Codigos.Add LstReg.List(LstReg.ListIndex)
    Else
        For Posicion = 0 To LstReg.ListCount - 1: Codigos.Add LstReg.List(Posicion): Next
    End If
    If Codigos.Count = 0 Then Exit Sub
    If MsgBox("Desea imprimir " & CStr(Codigos.Count) & " registros en hoja?", vbQuestion Or vbYesNo, "Imprimir Registros") = vbYes Then ImprimirGuardados Codigos, CmbModalReImp.Text
End Sub

Private Sub txtRegenerarMasivo_KeyPress(KeyAscii As Integer)
    Dim Datos As ADODB.Recordset, Columna As Integer, Fila As Long
    On Error GoTo ErrInt
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    If Not CodigoValido(txtRegenerarMasivo.Text) Then Err.Raise vbObjectError + 348, , "Especifique un codigo de registro valido."
    Set Datos = ConsultarDatos("SELECT Codigo,Cliente,CodCliente,Producto,Registro,Empresa,Direccion,Distrito,Contacto,Fecha FROM dbo.ViewRegistAct WHERE Codigo=?", txtRegenerarMasivo.Text, True)
    If Datos.EOF Then Datos.Close: Err.Raise vbObjectError + 349, , "No se encontro el registro."
    Grid.Rows = Grid.Rows + 1: Fila = Grid.Rows - 1
    For Columna = 0 To 9: Grid.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
    Dim Codigos As Collection
    Set Codigos = New Collection: Codigos.Add CStr(Datos.Fields(0).Value)
    Datos.Close
    ImprimirGuardados Codigos, CmbModalReImp.Text
    txtRegenerarMasivo.Text = ""
    ActualizarEstado
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub TxtNewEmp_KeyDown(KeyCode As Integer, Shift As Integer)
    OmitirCompletar = (KeyCode = vbKeyBack Or KeyCode = vbKeyDelete)
End Sub

Private Sub TxtNewEmp_Change()
    Dim Datos As ADODB.Recordset, Texto As String, Posicion As Long, Indice As Long
    On Error GoTo ErrInt
    If Completando Then Exit Sub
    LblCodEmp.Caption = ""
    If OmitirCompletar Or Len(TxtNewEmp.Text) = 0 Then OmitirCompletar = False: Exit Sub
    Texto = Replace(Replace(Replace(TxtNewEmp.Text, "[", "[[]"), "%", "[%]"), "_", "[_]")
    Set Datos = ConsultarDatos("SELECT TOP 1 CodEmp,NomEmp,DirEmp,DisEmp,ConEmp FROM dbo.MDBHisEmp WHERE NomEmp LIKE ? ORDER BY CodEmp DESC", Texto & "%", True)
    If Not Datos.EOF Then
        Completando = True
        Posicion = TxtNewEmp.SelStart
        TxtNewEmp.Text = TextoDato(Datos.Fields(1).Value)
        TxtNewEmp.SelStart = Posicion: TxtNewEmp.SelLength = Len(TxtNewEmp.Text) - Posicion
        LblCodEmp.Caption = CStr(Datos.Fields(0).Value): TxtNewDir.Text = TextoDato(Datos.Fields(2).Value)
        For Indice = 0 To CmbNewDis.ListCount - 1
            If CmbNewDis.ItemData(Indice) = CLng(Datos.Fields(3).Value) Then CmbNewDis.ListIndex = Indice: Exit For
        Next
        txtNewCont.Text = TextoDato(Datos.Fields(4).Value)
        Completando = False
    End If
    Datos.Close
    Exit Sub
ErrInt:
    Completando = False
    InformarError
End Sub

Private Sub TxtNewEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: CmdAgreNew.SetFocus
End Sub

Private Sub TxtNewDat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: TxtNewEmp.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Or Importando Or Guardando Then Exit Sub
    KeyCode = 0
    If PanelActual <> 0 Then
        CancelarIngreso
    ElseIf TxtCan.Text <> CStr(Cantidad) Then
        RestaurarCantidad
    End If
End Sub

Public Function PuedeCerrar() As Boolean
    If Importando Or Guardando Then Exit Function
    PuedeCerrar = True
    If HayCamposPendientes() Then PuedeCerrar = (MsgBox("Desea salir sin guardar los campos pendientes?", vbQuestion Or vbYesNo, "Clever SoftPro") = vbYes)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then Cancel = Not PuedeCerrar()
End Sub

Private Sub InformarError()
    Cargando = False
    MsgBox "No se pudo completar la operacion: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Function HayCamposPendientes() As Boolean
    HayCamposPendientes = (TxtCan.Text <> CStr(Cantidad))
    Select Case PanelActual
        Case 1
            HayCamposPendientes = HayCamposPendientes Or HayDocumentoPendiente()
        Case 2
            HayCamposPendientes = HayCamposPendientes Or (Len(TxtCli.Text & TxtBol.Text & TxtCanServ.Text) > 0 Or CmbServ.ListIndex >= 0 Or CmbProd.ListIndex >= 0)
        Case 4
            HayCamposPendientes = HayCamposPendientes Or HayDocumentoPendiente() Or (Not CargaTerminada And Len(txtexaminar.Text & Txtdes.Text) > 0)
    End Select
End Function

Private Function HayDocumentoPendiente() As Boolean
    HayDocumentoPendiente = (Len(TxtNewDat.Text & TxtNewEmp.Text & TxtNewDir.Text & txtNewCont.Text & TxtDescrip.Text) > 0 Or CmbNewDis.ListIndex >= 0 Or ChkNewDat.Value = 1)
End Function

Private Sub TxtCan_GotFocus()
    TxtCan.SelStart = 0: TxtCan.SelLength = Len(TxtCan.Text)
End Sub
Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 3
End Sub
