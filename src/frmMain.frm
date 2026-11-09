VERSION 5.00
Begin VB.Form frmMain
   BorderStyle     =   0  'None
   Caption         =   "Clever SoftPro X6"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   BeginProperty Font
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5475
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrInicio
      Enabled         =   -1  'True
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label LblVersion
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Versión interna 6.1"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7500
      TabIndex        =   0
      Top             =   195
      Width           =   2265
   End
   Begin VB.Label LblDescripcion
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión de distribución y control de documentos"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00503A2B&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   1605
      Width           =   9300
   End
   Begin VB.Label LblAutorTitulo
      BackStyle       =   0  'Transparent
      Caption         =   "Diseñado y desarrollado por"
      ForeColor       =   &H00505050&
      Height          =   225
      Left            =   360
      TabIndex        =   2
      Top             =   4110
      Width           =   7200
   End
   Begin VB.Label LblAutor
      BackStyle       =   0  'Transparent
      Caption         =   "Juan Pablo Aguilar Campos"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00333333&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   4395
      Width           =   7200
   End
   Begin VB.Label LblCorreo
      BackStyle       =   0  'Transparent
      Caption         =   "jaguilar.rationalsys@outlook.com"
      ForeColor       =   &H00505050&
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   4740
      Width           =   7200
   End
   Begin VB.Label LblFecha
      BackStyle       =   0  'Transparent
      Caption         =   "Lima, Perú · Mayo de 2011"
      ForeColor       =   &H00505050&
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   5100
      Width           =   7200
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Segundos As Integer

Private Sub TmrInicio_Timer()
    Segundos = Segundos + 1
    If Segundos < 5 Then Exit Sub
    TmrInicio.Enabled = False
    MdiConect.IniciarSesion
End Sub
