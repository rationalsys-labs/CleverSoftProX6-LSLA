VERSION 5.00
Begin VB.Form FrmSerch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Registros"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "FrmSerch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "Buscar Registros"
      Height          =   330
      Left            =   6255
      TabIndex        =   1
      Top             =   495
      Width           =   1725
   End
   Begin VB.TextBox txtBus 
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   7845
   End
End
Attribute VB_Name = "FrmSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdbuscar_Click()
    On Error GoTo ErrInt
    FrmMenuFlot.FiltrarRegistros txtBus.Text
    Unload Me
    Exit Sub
ErrInt:
    MsgBox "No se pudo aplicar el filtro: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub txtBus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdbuscar_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: Unload Me
End Sub
