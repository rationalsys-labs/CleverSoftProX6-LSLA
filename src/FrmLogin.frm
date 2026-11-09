VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autentificación de Usuario"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6255
   Begin VB.CommandButton CmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   4410
      TabIndex        =   8
      Top             =   2430
      Width           =   1590
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2745
      TabIndex        =   7
      Top             =   2430
      Width           =   1590
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2430
      Width           =   1590
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox CmbUsuario 
      Height          =   315
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1620
      Width           =   3285
   End
   Begin VB.TextBox TxtPasword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2700
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2025
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   0
      Picture         =   "FrmLogin.frx":0000
      ScaleHeight     =   1365
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   0
      Width           =   6270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      Height          =   195
      Left            =   1890
      TabIndex        =   4
      Top             =   1665
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Pasword:"
      Height          =   195
      Left            =   1890
      TabIndex        =   3
      Top             =   2070
      Width           =   660
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Validando As Boolean

Private Sub Form_Load()
    Dim Datos As ADODB.Recordset
    Dim UsuarioWindows As String
    On Error GoTo ErrInt
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    CmdAceptar.Default = True
    CmdCancelar.Cancel = True
    CmbUsuario.Clear
    Set Datos = CargarUsuarios()
    UsuarioWindows = Left$(get_Usuario(), 5)
    Do Until Datos.EOF
        If Not IsNull(Datos.Fields(0).Value) Then
            CmbUsuario.AddItem CStr(Datos.Fields(0).Value)
            If StrComp(CStr(Datos.Fields(0).Value), UsuarioWindows, vbTextCompare) = 0 Then CmbUsuario.ListIndex = CmbUsuario.NewIndex
        End If
        Datos.MoveNext
    Loop
    Datos.Close
    If CmbUsuario.ListCount = 0 Then
        StatusBar1.Panels(2).Text = "No hay usuarios disponibles."
        CmdAceptar.Enabled = False
    End If
    MDIPri.CentrarLogin
    Exit Sub
ErrInt:
    CmdAceptar.Enabled = False
    StatusBar1.Panels(2).Text = "No se pudo conectar con la base de datos."
    MsgBox "No se pudo cargar la lista de usuarios: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Function get_Usuario() As String
    Dim Nombre As String
    Dim Longitud As Long
    Nombre = Space$(256)
    Longitud = Len(Nombre)
    If GetUserName(Nombre, Longitud) <> 0 Then get_Usuario = Left$(Nombre, Longitud - 1)
End Function

Private Sub CmdAceptar_Click()
    If Validando Then Exit Sub
    If Len(CmbUsuario.Text) = 0 Then
        MsgBox "Seleccione el usuario que desea validar.", vbInformation, "Clever SoftPro"
        CmbUsuario.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrInt
    Validando = True
    CmdAceptar.Enabled = False
    StatusBar1.Panels(2).Text = "Validando usuario..."
    If AutenticarUsuario(CmbUsuario.Text, TxtPasword.Text) Then
        VarUsuario = CmbUsuario.Text
        EjecuteDatrav = True
        MDIPri.ActualizarSesion
        TxtPasword.Text = vbNullString
        Unload Me
        Exit Sub
    End If
    StatusBar1.Panels(2).Text = "Usuario o clave incorrectos."
    MsgBox "Los datos del usuario no son correctos. Verifique e intente nuevamente.", vbExclamation, "Clever SoftPro"
    TxtPasword.SetFocus
    TxtPasword.SelStart = 0
    TxtPasword.SelLength = Len(TxtPasword.Text)
Salida:
    Validando = False
    CmdAceptar.Enabled = True
    Exit Sub
ErrInt:
    StatusBar1.Panels(2).Text = "No se pudo completar la validacion."
    MsgBox "No se pudo validar el usuario: " & Err.Description, vbExclamation, "Clever SoftPro"
    Resume Salida
End Sub

Private Sub CmdCancelar_Click()
    MdiConect.CerrarAplicacion
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not EjecuteDatrav And Not CerrandoAplicacion Then
        Cancel = True
        Me.Hide
        MDIPri.ProgramarCierre
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TxtPasword.Text = vbNullString
End Sub

Private Sub CmdHelp_Click()
    MsgBox "Seleccione su usuario e ingrese su clave de Clever. Para problemas de acceso, consulte al administrador.", vbInformation, "Inicio de sesion"
End Sub

Private Sub TxtPasword_KeyPress(KeyAscii As Integer)
    'Se conserva el conjunto de caracteres aceptado por el login original.
    Select Case KeyAscii
        Case 48 To 57, 65 To 90, 97 To 122, 42, 8
        Case 13
            KeyAscii = 0
            CmdAceptar_Click
        Case Else
            KeyAscii = 0
    End Select
End Sub