VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmActualizaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizaciones"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   KeyPreview = -1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCerrar
      Caption = "Cerrar"
      Left = 225
      Top = 2760
      Width = 1200
      Height = 330
      TabIndex = 10
   End
   Begin VB.Label LblResultado
      Caption = "Seleccione destino para exportar pendientes del esquema 2."
      Left = 1530
      Top = 2760
      Width = 4300
      Height = 330
   End
   Begin VB.OptionButton OptCar 
      Caption         =   "Cargar Archivo de Actualizacion"
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   1440
      Width           =   2760
   End
   Begin VB.OptionButton OptCre 
      Caption         =   "Generar Archivo de Actualizacion"
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   90
      Value           =   -1  'True
      Width           =   2760
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   90
      TabIndex        =   9
      Top             =   135
      Width           =   5820
      Begin VB.TextBox TxtDir 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   3795
      End
      Begin VB.CommandButton CmdDestino 
         Caption         =   "Destino"
         Height          =   330
         Left            =   4095
         TabIndex        =   2
         Top             =   225
         Width           =   1590
      End
      Begin VB.CommandButton CmdGuardar 
         Caption         =   "Guardar"
         Height          =   465
         Left            =   4095
         TabIndex        =   3
         Top             =   630
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   90
      TabIndex        =   8
      Top             =   1485
      Width           =   5820
      Begin VB.TextBox TxtDirAct 
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   3795
      End
      Begin VB.CommandButton CmdExaminar 
         Caption         =   "Examinar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4095
         TabIndex        =   6
         Top             =   225
         Width           =   1590
      End
      Begin VB.CommandButton CmdCargar 
         Caption         =   "Cargar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   4095
         TabIndex        =   7
         Top             =   630
         Width           =   1590
      End
   End
   Begin MSComDlg.CommonDialog CUAD 
      Left            =   180
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmActualizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Procesando As Boolean

Private Sub Form_Load()
    TxtDir.Locked = True: TxtDirAct.Locked = True
    CambiarModo
End Sub

Private Sub CambiarModo()
    Frame2.Enabled = OptCre.Value And Not Procesando
    Frame1.Enabled = OptCar.Value And Not Procesando
    TxtDirAct.Enabled = (OptCar.Value And Not Procesando): CmdExaminar.Enabled = TxtDirAct.Enabled
    CmdCargar.Enabled = (TxtDirAct.Enabled And Len(TxtDirAct.Text) > 0)
    CmdGuardar.Enabled = (OptCre.Value And Not Procesando And Len(TxtDir.Text) > 0)
    CmdDestino.Enabled = (OptCre.Value And Not Procesando): CmdCerrar.Enabled = Not Procesando
    OptCre.Enabled = Not Procesando: OptCar.Enabled = Not Procesando
    If Procesando Then
        LblResultado.Caption = "Procesando el archivo completo..."
    ElseIf OptCre.Value Then
        LblResultado.Caption = "Seleccione destino para exportar pendientes del esquema 2."
    Else
        LblResultado.Caption = "Seleccione un archivo de actualizacion del esquema 2."
    End If
End Sub

Private Sub OptCre_Click()
    CambiarModo
End Sub

Private Sub OptCar_Click()
    CambiarModo
End Sub

Private Sub CmdDestino_Click()
    On Error GoTo ErrInt
    CUAD.CancelError = True: CUAD.FileName = "": CUAD.DialogTitle = "Guardar"
    CUAD.Filter = "Archivos de actualizacion (*.sql)|*.sql": CUAD.DefaultExt = "sql"
    CUAD.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    CUAD.ShowSave
    TxtDir.Text = CUAD.FileName
    Exit Sub
ErrInt:
    If Err.Number <> cdlCancel Then MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdExaminar_Click()
    On Error GoTo ErrInt
    CUAD.CancelError = True: CUAD.FileName = "": CUAD.DialogTitle = "Abrir"
    CUAD.Filter = "Archivos de actualizacion (*.sql)|*.sql"
    CUAD.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
    CUAD.ShowOpen
    TxtDirAct.Text = CUAD.FileName
    Exit Sub
ErrInt:
    If Err.Number <> cdlCancel Then MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdGuardar_Click()
    Dim Cantidad As Long
    On Error GoTo ErrInt
    If Procesando Or Not OptCre.Value Or Len(TxtDir.Text) = 0 Then Exit Sub
    If Len(Dir$(TxtDir.Text)) > 0 Then
        If MsgBox("El archivo ya existe: " & TxtDir.Text & vbCrLf & "Desea reemplazarlo?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    End If
    Procesando = True: CambiarModo
    Cantidad = ExportarActualizaciones(TxtDir.Text, True)
    Procesando = False: CambiarModo
    LblResultado.Caption = CStr(Cantidad) & " actualizaciones exportadas."
    MsgBox CStr(Cantidad) & " actualizaciones exportadas.", vbInformation, "Clever SoftPro"
    Exit Sub
ErrInt:
    Procesando = False: CambiarModo
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdCargar_Click()
    Dim Cantidad As Long
    On Error GoTo ErrInt
    If Procesando Or Not OptCar.Value Or Len(TxtDirAct.Text) = 0 Then Exit Sub
    If MsgBox("Se aplicara el archivo completo: " & TxtDirAct.Text & vbCrLf & "Si ocurre un error se revertira todo el lote. Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    Procesando = True: CambiarModo
    Cantidad = ImportarActualizaciones(TxtDirAct.Text)
    Procesando = False: CambiarModo
    LblResultado.Caption = CStr(Cantidad) & " actualizaciones cargadas."
    MsgBox CStr(Cantidad) & " actualizaciones cargadas.", vbInformation, "Clever SoftPro"
    Exit Sub
ErrInt:
    Procesando = False: CambiarModo
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Procesando Then Cancel = True
End Sub
Private Sub TxtDir_Change()
    CambiarModo
End Sub
Private Sub TxtDirAct_Change()
    CambiarModo
End Sub
Private Sub CmdCerrar_Click()
    If Not Procesando Then Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Not Procesando Then KeyCode = 0: Unload Me
End Sub
Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 11
End Sub
