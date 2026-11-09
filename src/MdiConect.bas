Attribute VB_Name = "MdiConect"
Option Explicit

Public Sub Main()
    On Error GoTo ErrInt
    EjecuteDatrav = False
    CerrandoAplicacion = False
    frmMain.Show
    Exit Sub
ErrInt:
    MsgBox "No se pudo iniciar Clever: " & Err.Description, vbCritical, "Clever SoftPro"
    CerrarAplicacion
End Sub

Public Sub IniciarSesion()
    On Error GoTo ErrInt
    MDIPri.Show
    Unload frmMain
    FrmLogin.Show
    Exit Sub
ErrInt:
    MsgBox "No se pudo abrir el inicio de sesion: " & Err.Description, vbCritical, "Clever SoftPro"
    CerrarAplicacion
End Sub

Public Sub CerrarAplicacion()
    Dim i As Long
    If CerrandoAplicacion Then Exit Sub
    If Not PuedeCerrarFormularios() Then Exit Sub
    CerrandoAplicacion = True
    Unload RepConsol
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i
    VarUsuario = vbNullString
    EjecuteDatrav = False
End Sub

Public Function PuedeCerrarFormularios() As Boolean
    Dim Formulario As Object
    PuedeCerrarFormularios = True
    For Each Formulario In Forms
        If TypeOf Formulario Is FrmManten Or TypeOf Formulario Is FrmEtiquetas Or TypeOf Formulario Is FrmgestServ Or TypeOf Formulario Is FrmReparto Or TypeOf Formulario Is Frmconsolidado Or TypeOf Formulario Is FrmConfirm Or TypeOf Formulario Is FrmConfig Or TypeOf Formulario Is FrmUsuarios Then
            If Not Formulario.PuedeCerrar() Then
                PuedeCerrarFormularios = False
                Exit Function
            End If
        End If
    Next
End Function
