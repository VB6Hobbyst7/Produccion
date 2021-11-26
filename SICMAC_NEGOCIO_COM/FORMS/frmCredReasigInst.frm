VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCredReasigInst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reasignar Institucion"
   ClientHeight    =   4275
   ClientLeft      =   750
   ClientTop       =   2625
   ClientWidth     =   10635
   Icon            =   "frmCredReasigInst.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   405
      Left            =   8295
      TabIndex        =   13
      Top             =   3825
      Width           =   1080
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   9390
      TabIndex        =   12
      Top             =   3810
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Height          =   2880
      Left            =   135
      TabIndex        =   5
      Top             =   825
      Width           =   10410
      Begin VB.CommandButton CmdQuitarT 
         Caption         =   "To&dos <<"
         Height          =   390
         Left            =   4680
         TabIndex        =   10
         Top             =   2085
         Width           =   1035
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "&Quitar"
         Height          =   390
         Left            =   4680
         TabIndex        =   9
         Top             =   1575
         Width           =   1035
      End
      Begin VB.CommandButton CmdAgregarT 
         Caption         =   "&Todos >>"
         Height          =   390
         Left            =   4680
         TabIndex        =   8
         Top             =   900
         Width           =   1035
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   390
         Left            =   4680
         TabIndex        =   7
         Top             =   435
         Width           =   1035
      End
      Begin MSComctlLib.ListView LstCredInst 
         Height          =   2565
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Credito"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView LstCredInstN 
         Height          =   2565
         Left            =   5805
         TabIndex        =   11
         Top             =   195
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Credito"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   10410
      Begin VB.ComboBox CmbInstitucionN 
         Height          =   315
         Left            =   5550
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   4710
      End
      Begin VB.ComboBox CmbInstitucion 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   4425
      End
      Begin VB.Label Label2 
         Caption         =   "Institucion Reasignada : "
         Height          =   195
         Left            =   5550
         TabIndex        =   4
         Top             =   165
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Institucion Asignada : "
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   180
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmCredReasigInst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GrabarNuevaCartera()
Dim i As Integer
'Dim oCredito As COMDCredito.DCOMCredActBD
Dim oCredito As COMDCredito.DCOMCredito
Dim MatCuentas As Variant

    On Error GoTo ErrorGrabarNuevaCartera
'    Set oCredito = New COMDCredito.DCOMCredActBD
    ReDim MatCuentas(LstCredInstN.ListItems.Count)
    
    For i = 1 To LstCredInstN.ListItems.Count
        'Call oCredito.dUpdateColocacConvenio(LstCredInstN.ListItems(i).Text, Trim(Right(CmbInstitucionN.Text, 20)), False)
        MatCuentas(i) = LstCredInstN.ListItems(i).Text
    Next i
    
    Set oCredito = New COMDCredito.DCOMCredito
    Call oCredito.ReasignaInstituciones(MatCuentas, Trim(Right(CmbInstitucionN.Text, 20)))
    'Jame ------'
    Call RegistraHistorialConvenio(MatCuentas)
    'fin jame'
    Set oCredito = Nothing
    Exit Sub

ErrorGrabarNuevaCartera:
        MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CargaInstitucion()
Dim oPersonas As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset

    On Error GoTo ERRORCargaAnalistas
    CmbInstitucion.Clear
    CmbInstitucionN.Clear
    
    Set oPersonas = New COMDPersona.DCOMPersonas
    Set rs = oPersonas.RecuperaPersonasTipo(gPersTipoConvenio)
    Set oPersonas = Nothing
    
    Do While Not rs.EOF
        CmbInstitucion.AddItem PstaNombre(rs!cPersNombre) & Space(250) & rs!cPersCod
        CmbInstitucionN.AddItem PstaNombre(rs!cPersNombre) & Space(250) & rs!cPersCod
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    'Carga Instituciones
    'Call CargaComboPersonasTipo(gPersTipoConvenio, CmbInstitucion)
    'Call CargaComboPersonasTipo(gPersTipoConvenio, CmbInstitucionN)
    
    If CmbInstitucion.ListCount > 0 Then
        CmbInstitucion.ListIndex = 0
    End If
    If CmbInstitucionN.ListCount > 0 Then
        CmbInstitucionN.ListIndex = 0
    End If
    Exit Sub
ERRORCargaAnalistas:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaListaCarteraAnalista()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim L As ListItem

    On Error GoTo ErrorCargaListaCarteraAnalista
    Set oCredito = New COMDCredito.DCOMCredito
    Set R = oCredito.CarteraInstitucion(Right(CmbInstitucion.Text, 20))
    Set oCredito = Nothing
    LstCredInst.ListItems.Clear
    Do While Not R.EOF
        Set L = LstCredInst.ListItems.Add(, , R!cCtaCod)
        L.SubItems(1) = R!cPersNombre
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ErrorCargaListaCarteraAnalista:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmbInstitucion_Click()
    Screen.MousePointer = 11
    Call CargaListaCarteraAnalista
    LstCredInstN.ListItems.Clear
    Screen.MousePointer = 0
End Sub

Private Sub CmbInstitucionN_Click()
    LstCredInstN.ListItems.Clear
End Sub

Private Sub CmdAgregar_Click()
Dim L As ListItem

    On Error GoTo ErrorCmdAgregar_Click
    
    If LstCredInst.ListItems.Count <= 0 Then
        MsgBox "No Existen registros para Reasignar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set L = LstCredInstN.ListItems.Add(, , LstCredInst.SelectedItem.Text)
    L.SubItems(1) = LstCredInst.SelectedItem.SubItems(1)
    Call LstCredInst.ListItems.Remove(LstCredInst.SelectedItem.Index)
    Exit Sub

ErrorCmdAgregar_Click:
        MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CmdAgregarT_Click()
Dim i As Integer
Dim L As ListItem
    On Error GoTo ErrorCmdAgregarT_Click
    Screen.MousePointer = 11
    For i = 1 To LstCredInst.ListItems.Count
        Set L = LstCredInstN.ListItems.Add(, , LstCredInst.ListItems(i).Text)
        L.SubItems(1) = LstCredInst.ListItems(i).SubItems(1)
    Next i
    
    Do While LstCredInst.ListItems.Count > 0
        Call LstCredInst.ListItems.Remove(1)
    Loop
    'LblTotalAna.Caption = LstCredInst.ListItems.Count
    'LblTotalAnaN.Caption = LstCredInstN.ListItems.Count
    Screen.MousePointer = 0
    Exit Sub

ErrorCmdAgregarT_Click:
    Screen.MousePointer = 0
        MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdGrabar_Click()
    If LstCredInstN.ListItems.Count > 0 Then
        If MsgBox("Se va a Actualizar la Institucion : " & Trim(Left(CmbInstitucionN.Text, 40)) & ", Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Call GrabarNuevaCartera
            LstCredInstN.ListItems.Clear
        End If
    Else
        MsgBox "No Existen Creditos a Reasignar", vbInformation, "Aviso"
    End If
End Sub

Private Sub CmdQuitar_Click()
Dim L As ListItem

On Error GoTo ErrorCmdQuitar_Click
    
    If LstCredInstN.ListItems.Count <= 0 Then
        MsgBox "No Existen Registros para Reasignar", vbInformation, "Aviso"
        Exit Sub
    End If
    Set L = LstCredInst.ListItems.Add(, , LstCredInstN.SelectedItem.Text)
    L.SubItems(1) = LstCredInstN.SelectedItem.SubItems(1)
    Call LstCredInstN.ListItems.Remove(LstCredInstN.SelectedItem.Index)
   ' LblTotalAna.Caption = LstCredInst.ListItems.Count
   ' LblTotalAnaN.Caption = LstCredInstN.ListItems.Count
    Exit Sub

ErrorCmdQuitar_Click:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdQuitarT_Click()
Dim L As ListItem
Dim i As Integer
    On Error GoTo ErrorCmdQuitarT_Click
    Screen.MousePointer = 11
    For i = 1 To LstCredInstN.ListItems.Count
        Set L = LstCredInst.ListItems.Add(, , LstCredInstN.ListItems(i).Text)
        L.SubItems(1) = LstCredInstN.ListItems(i).SubItems(1)
        
    Next i
    Do While LstCredInstN.ListItems.Count > 0
        Call LstCredInstN.ListItems.Remove(1)
    Loop
    'LblTotalAna.Caption = LstCredInst.ListItems.Count
    'LblTotalAnaN.Caption = LstCredInstN.ListItems.Count
    Screen.MousePointer = 0
    Exit Sub

ErrorCmdQuitarT_Click:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Call CargaInstitucion
End Sub

Private Sub LstCredInst_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstCredInst.SortKey = ColumnHeader.SubItemIndex
    LstCredInst.SortOrder = lvwAscending
    LstCredInst.Sorted = True
End Sub

Private Sub LstCredInst_DblClick()
    If LstCredInst.ListItems.Count > 0 Then
        CmdAgregar_Click
    End If
End Sub

Private Sub LstCredInstN_DblClick()
    If LstCredInstN.ListItems.Count > 0 Then
        CmdQuitar_Click
    End If
End Sub
'jame **********'
Private Sub RegistraHistorialConvenio(ByVal pMatCuentas)
    Dim oCredito As COMDCredito.DCOMCredito
    Dim i As Integer
    Set oCredito = New COMDCredito.DCOMCredito
 
    For i = 1 To UBound(pMatCuentas)
        Call oCredito.RegistroHistorialConvenio(pMatCuentas(i), gsCodUser, Format(gdFecSis & " " & GetHoraServer, "yyyy/MM/dd hh:mm:ss"), Trim(Right(CmbInstitucion.Text, 20)), Trim(Right(CmbInstitucionN.Text, 20)), 2)
    
    Next
End Sub
'jame fin *******'
