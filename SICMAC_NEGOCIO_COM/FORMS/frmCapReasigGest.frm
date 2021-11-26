VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapReasigGest 
   Caption         =   "Reasignacion de Cuentas de PF y CTS"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   Icon            =   "frmCapReasigGest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8745
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   7425
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cambio de Gestor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   105
      TabIndex        =   2
      Top             =   720
      Width           =   9855
      Begin VB.ComboBox cmbGestorReasig 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox cmbGestorAsig 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   4095
      End
      Begin VB.CheckBox chkCTS 
         Caption         =   "CTS"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkPF 
         Caption         =   "Plazo Fijo"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gestor Reasignado:"
         Height          =   195
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Gestor Asignado:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbAgencias 
      Height          =   315
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   105
      TabIndex        =   0
      Top             =   2160
      Width           =   9855
      Begin VB.CommandButton cmdQTodos 
         Caption         =   "<< T&odos"
         Height          =   375
         Left            =   4425
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   375
         Left            =   4425
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdTodos 
         Caption         =   "&Todos >>"
         Height          =   375
         Left            =   4425
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   4425
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin MSComctlLib.ListView lstCtasGes 
         Height          =   2895
         Left            =   45
         TabIndex        =   17
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5106
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombres"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nº Cuentas"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstCtasGesN 
         Height          =   2895
         Left            =   5565
         TabIndex        =   18
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5106
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nº Cuentas"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5745
      TabIndex        =   22
      Top             =   5520
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   21
      Top             =   5520
      Width           =   510
   End
   Begin VB.Label lblTotalGesN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6465
      TabIndex        =   20
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lblTotalGes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   705
      TabIndex        =   19
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmCapReasigGest.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAgencia 
      AutoSize        =   -1  'True
      Caption         =   "Agencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "frmCapReasigGest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCTS_Click()
    Screen.MousePointer = 11
    Call CargaListaCuentasGestor
    lstCtasGesN.ListItems.Clear
    Screen.MousePointer = 0
End Sub

Private Sub chkPF_Click()
    Screen.MousePointer = 11
    Call CargaListaCuentasGestor
    lstCtasGesN.ListItems.Clear
    Screen.MousePointer = 0
End Sub

Private Sub cmbGestorAsig_Click()
    If cmbAgencias.ListIndex = -1 Then
        MsgBox "Debe seleccionar una agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 11
    Call CargaListaCuentasGestor
    lstCtasGesN.ListItems.Clear
    Screen.MousePointer = 0
End Sub

Private Sub cmbGestorReasig_Change()
    lstCtasGes.ListItems.Clear
End Sub

Private Sub CmdAgregar_Click()
Dim L As ListItem

    On Error GoTo ErrorCmdAgregar_Click
    
    Set L = lstCtasGesN.ListItems.Add(, , lstCtasGes.SelectedItem.Text)
    L.SubItems(1) = lstCtasGes.SelectedItem.SubItems(1)
    L.SubItems(2) = lstCtasGes.SelectedItem.SubItems(2)
    Call lstCtasGes.ListItems.Remove(lstCtasGes.SelectedItem.Index)
    lblTotalGes.Caption = lstCtasGes.ListItems.Count
    lblTotalGesN.Caption = lstCtasGesN.ListItems.Count
    Exit Sub

ErrorCmdAgregar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub
Private Sub cmdGuardar_Click()
    If lstCtasGesN.ListItems.Count > 0 Then
        If MsgBox("Se va a Actualizar la Cartera del Gestor : " & Trim(Left(cmbGestorReasig.Text, 40)) & ", Desea Continuar ? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Call GrabarNuevaCartera
            lstCtasGesN.ListItems.Clear
        End If
    Else
        MsgBox "No Existen Creditos a Reasignar", vbInformation, "Aviso"
    End If
    lblTotalGes.Caption = lstCtasGes.ListItems.Count
    lblTotalGesN.Caption = lstCtasGesN.ListItems.Count
End Sub

Private Sub cmdQTodos_Click()
Dim L As ListItem
Dim i As Integer
    On Error GoTo ErrorCmdQuitarT_Click
    Screen.MousePointer = 11
    For i = 1 To lstCtasGesN.ListItems.Count
        Set L = lstCtasGes.ListItems.Add(, , lstCtasGesN.ListItems(i).Text)
        L.SubItems(1) = lstCtasGesN.ListItems(i).SubItems(1)
        L.SubItems(2) = lstCtasGesN.ListItems(i).SubItems(2)
    Next i
    Do While lstCtasGesN.ListItems.Count > 0
        Call lstCtasGesN.ListItems.Remove(1)
    Loop
    lblTotalGes.Caption = lstCtasGes.ListItems.Count
    lblTotalGesN.Caption = lstCtasGesN.ListItems.Count
    Screen.MousePointer = 0
    Exit Sub

ErrorCmdQuitarT_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdQuitar_Click()
Dim L As ListItem

On Error GoTo ErrorCmdQuitar_Click
    Set L = lstCtasGes.ListItems.Add(, , lstCtasGesN.SelectedItem.Text)
    L.SubItems(1) = lstCtasGesN.SelectedItem.SubItems(1)
    L.SubItems(2) = lstCtasGesN.SelectedItem.SubItems(2)
    Call lstCtasGesN.ListItems.Remove(lstCtasGesN.SelectedItem.Index)
    lblTotalGes.Caption = lstCtasGes.ListItems.Count
    lblTotalGesN.Caption = lstCtasGesN.ListItems.Count
    Exit Sub

ErrorCmdQuitar_Click:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdTodos_Click()
Dim i As Integer
Dim L As ListItem
    On Error GoTo ErrorCmdAgregarT_Click
    Screen.MousePointer = 11
    For i = 1 To lstCtasGes.ListItems.Count
        Set L = lstCtasGesN.ListItems.Add(, , lstCtasGes.ListItems(i).Text)
        L.SubItems(1) = lstCtasGes.ListItems(i).SubItems(1)
        L.SubItems(2) = lstCtasGes.ListItems(i).SubItems(2)
    Next i
    
    Do While lstCtasGes.ListItems.Count > 0
        Call lstCtasGes.ListItems.Remove(1)
    Loop
    
    lblTotalGes.Caption = lstCtasGes.ListItems.Count
    lblTotalGesN.Caption = lstCtasGesN.ListItems.Count
    Screen.MousePointer = 0
    Exit Sub

ErrorCmdAgregarT_Click:
    Screen.MousePointer = 0
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    Call CargarControles
End Sub

Private Sub CargarControles()
Dim rsGes As ADODB.Recordset
Dim rsAge As ADODB.Recordset
Dim i As Integer
Dim oCred As COMDCredito.DCOMCredito
Dim oCapGen As COMDCaptaGenerales.DCOMCaptaGenerales

On Error GoTo ERRORCargarControles
        
    Set oCred = New COMDCredito.DCOMCredito
    Set rsAge = oCred.RecuperaAgencias
    Set oCred = Nothing
       
    cmbAgencias.Clear
    
    Do Until rsAge.EOF
        cmbAgencias.AddItem rsAge!cAgeDescripcion & Space(100) & rsAge!cAgeCod
        rsAge.MoveNext
    Loop
    
    For i = 0 To cmbAgencias.ListCount - 1
        If Right(cmbAgencias.List(i), 2) = gsCodAge Then
            cmbAgencias.ListIndex = i
            Exit For
        End If
    Next i
    
    Set oCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsGes = oCapGen.GetPromotores(gsCodAge)
    Set oCapGen = Nothing
    
    cmbGestorAsig.Clear
    cmbGestorReasig.Clear
    Do While Not rsGes.EOF
        cmbGestorAsig.AddItem PstaNombre(rsGes!cPersNombre) & Space(100) & rsGes!cPersCod
        cmbGestorReasig.AddItem PstaNombre(rsGes!cPersNombre) & Space(100) & rsGes!cPersCod
        rsGes.MoveNext
    Loop
    cmbGestorAsig.AddItem "SIN GESTOR" & Space(100) & "00000000"
    
    If cmbGestorAsig.ListCount > 0 Then
        cmbGestorAsig.ListIndex = 0
        cmbGestorReasig.ListIndex = 0
    End If
    
    Exit Sub
    
ERRORCargarControles:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub CargaListaCuentasGestor()
Dim oCapGes As COMDCaptaGenerales.DCOMCaptaGenerales
Dim R As ADODB.Recordset
Dim L As ListItem
Dim sTipoPF As String
Dim sTipoCTS As String

    On Error GoTo ErrorCargaListaCuentasGestor
    
    If chkPF.value = 1 Then
       sTipoPF = "233"
    End If
    
    If chkCTS.value = 1 Then
       sTipoCTS = "234"
    End If
         
    Set oCapGes = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set R = oCapGes.CarteraGestor(Right(cmbGestorAsig.Text, 13), Right(cmbAgencias.List(cmbAgencias.ListIndex), 2), sTipoPF, sTipoCTS)
    Set oCapGes = Nothing
    lstCtasGes.ListItems.Clear
    Do While Not R.EOF
        If Right(cmbGestorAsig.Text, 8) = "00000000" Then
            Set L = lstCtasGes.ListItems.Add(, , R!NumCtas)
            L.SubItems(1) = R!cPersNombre
            L.SubItems(2) = R!cPersCod
        Else
            Set L = lstCtasGes.ListItems.Add(, , R!cPersCod)
            L.SubItems(1) = R!cPersNombre
            L.SubItems(2) = R!NumCtas
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    lblTotalGes.Caption = lstCtasGes.ListItems.Count
    lblTotalGesN.Caption = lstCtasGesN.ListItems.Count
    Exit Sub

ErrorCargaListaCuentasGestor:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub GrabarNuevaCartera()
Dim i As Integer
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oCredito As COMDCredito.DCOMCredito
Dim MatCuentas As Variant
Dim sTipoPF As String
Dim sTipoCTS As String
Dim sCodAge As String
    
    On Error GoTo ErrorGrabarNuevaCartera
    Set oCredito = New COMDCredito.DCOMCredito
    ReDim MatCuentas(lstCtasGesN.ListItems.Count)
    
    For i = 1 To lstCtasGesN.ListItems.Count
        MatCuentas(i) = lstCtasGesN.ListItems(i).Text
    Next i
    
    If chkPF.value = 1 Then
       sTipoPF = "233"
    End If
    
    If chkCTS.value = 1 Then
       sTipoCTS = "234"
    End If
    
    sCodAge = Right(cmbAgencias.List(cmbAgencias.ListIndex), 2)
    
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCap.ReasignaCarteraGestor(MatCuentas, Right(cmbGestorReasig.Text, 13), Right(cmbGestorAsig.Text, 13), gdFecSis, sTipoPF, sTipoCTS, sCodAge)
    Set oCredito = Nothing
    Exit Sub

ErrorGrabarNuevaCartera:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub
Private Sub lstCtasGes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstCtasGes.SortKey = ColumnHeader.SubItemIndex
    lstCtasGes.SortOrder = lvwAscending
    lstCtasGes.Sorted = True
End Sub

Private Sub lstCtasGes_DblClick()
    If lstCtasGes.ListItems.Count > 0 Then
        CmdAgregar_Click
    End If
End Sub

Private Sub lstCtasGesN_DblClick()
    If lstCtasGesN.ListItems.Count > 0 Then
        CmdQuitar_Click
    End If
End Sub
