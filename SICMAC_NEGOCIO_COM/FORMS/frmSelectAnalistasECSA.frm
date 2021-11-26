VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectAnalistasECSA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Analistas"
   ClientHeight    =   5670
   ClientLeft      =   8265
   ClientTop       =   855
   ClientWidth     =   5340
   Icon            =   "frmSelectAnalistasECSA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   75
      ScaleHeight     =   5295
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   330
      Width           =   735
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      ScaleHeight     =   300
      ScaleWidth      =   5235
      TabIndex        =   5
      Top             =   45
      Width           =   5235
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   870
      TabIndex        =   4
      Top             =   315
      Width           =   4395
      Begin VB.OptionButton OptAnalista 
         Caption         =   "&Ninguno"
         Height          =   210
         Index           =   1
         Left            =   2565
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton OptAnalista 
         Caption         =   "&Todos"
         Height          =   210
         Index           =   0
         Left            =   855
         TabIndex        =   2
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2370
      TabIndex        =   1
      Top             =   5190
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analistas"
      Height          =   4260
      Left            =   900
      TabIndex        =   0
      Top             =   855
      Width           =   4365
      Begin MSComctlLib.ListView lvAnalistas 
         Height          =   3855
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   6800
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmSelectAnalistasECSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SeleccionaInstituciones()
    Call PrepararListView("Instituciones")
    Call CargaInstituciones
    Me.Caption = "Seleccionar Instituciones"
    Me.Frame1.Caption = "Instiruciones"
    Me.Show 1
End Sub

Public Sub SeleccionaAnalistas()
    Call PrepararListView("Analistas")
    Call CargaAnalistas
    Me.Caption = "Seleccionar Analista"
    OptAnalista(1).value = True
    OptAnalista_Click 1
    Me.Show 1
End Sub

Public Sub SeleccionaCampanas()
    Call PrepararListView("Campañas")
    Call CargaCampanas
    Me.Caption = "Seleccionar Campañas"
    Me.Frame1.Caption = "Campañas"
    Me.Show 1
End Sub

Private Sub CargaInstituciones()
Dim R As ADODB.Recordset
Dim ssql As String
Dim oconecta As COMConecta.DCOMConecta
Dim itm As ListItem
    
    On Error GoTo ERRORCargaAnalistas
    ssql = "select P.cPersNombre, PT.cPersCod "
    ssql = ssql & " from PersTpo PT Inner join Persona P ON PT.cPersCod = P.cPersCod And PT.nPersTipo = 1"
    ssql = ssql & " Order By P.cPersNombre"
    Set oconecta = New DConecta
    oconecta.AbreConexion
    Set R = oconecta.CargaRecordSet(ssql)
    oconecta.CierraConexion
    Set oconecta = Nothing
'    LstAnalista.Clear
'    Do While Not R.EOF
'        LstAnalista.AddItem R!cPersNombre & Space(100) & R!cPersCod
'        R.MoveNext
'    Loop
    
    lvAnalistas.ListItems.Clear
    If R.RecordCount > 0 Then R.MoveFirst
    Do While Not R.EOF
        Set itm = lvAnalistas.ListItems.Add(, , Trim(R!cPersNombre))
        itm.Tag = R!cPersCod
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    Exit Sub
    
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaAnalistas()
Dim R As ADODB.Recordset
Dim ssql As String
Dim oconecta As COMConecta.DCOMConecta
Dim itm As ListItem
Dim sAnalistas As String
Dim oGen  As DGeneral

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New DGeneral
    sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
    Set oGen = Nothing
    
    ssql = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod "
    ssql = ssql & " AND nRHEstado = 201 "
    ssql = ssql & " inner join RHCargos RC ON R.cPersCod = RC.cPersCod "
    ssql = ssql & " where  RC.cRHCargoCod in (" & sAnalistas & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) "
    ssql = ssql & " order by P.cPersNombre "
        
    Set oconecta = New DConecta
    oconecta.AbreConexion
    Set R = oconecta.CargaRecordSet(ssql)
    oconecta.CierraConexion
    Set oconecta = Nothing
    
'    LstAnalista.Clear
'    Do While Not R.EOF
'        LstAnalista.AddItem R!cPersNombre & Space(100) & R!cPersCod
'        R.MoveNext
'    Loop
    
    lvAnalistas.ListItems.Clear
    If R.RecordCount > 0 Then R.MoveFirst
    Do While Not R.EOF
    Set itm = lvAnalistas.ListItems.Add(, , Trim(R!cDescripcion))
        itm.Tag = R!IdCampana
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaCampanas()
Dim R As ADODB.Recordset
Dim ssql As String
Dim itm As ListItem
Dim oconecta As COMConecta.DCOMConecta
    
    On Error GoTo ERRORCargaCampanas
    ssql = "Select IdCampana, cDescripcion From Campanas Where bEstado = 1 And IdCampana <> 0"
    ssql = ssql & " Order By 1 "
    'Set oconecta = New DConecta
    Set oconecta = New COMConecta.DCOMConecta
    oconecta.AbreConexion
    Set R = oconecta.CargaRecordSet(ssql)
    oconecta.CierraConexion
    Set oconecta = Nothing
    
'    LstAnalista.Clear
'    Do While Not R.EOF
'        LstAnalista.AddItem R!cDescripcion & Space(100) & R!IdCampana
'        R.MoveNext
'    Loop
    
    lvAnalistas.ListItems.Clear
    If R.RecordCount > 0 Then R.MoveFirst
    Do While Not R.EOF
    Set itm = lvAnalistas.ListItems.Add(, , Trim(R!cDescripcion))
        itm.Tag = R!IdCampana
        R.MoveNext
    Loop
    
    R.Close
    Set R = Nothing
    Exit Sub
ERRORCargaCampanas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdAceptar_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    CentraForm Me
    OptAnalista(1).value = True
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    OptAnalista(1).value = True
End Sub

Private Sub lvAnalistas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub OptAnalista_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
'    If LstAnalista.ListCount <= 0 Then
'        Exit Sub
'    End If
'    For i = 0 To LstAnalista.ListCount - 1
'        LstAnalista.Selected(i) = bCheck
'    Next i
    If lvAnalistas.ListItems.Count <= 0 Then
        Exit Sub
    End If
    For i = 1 To lvAnalistas.ListItems.Count
        lvAnalistas.ListItems(i).Checked = bCheck
    Next i
End Sub

Private Sub PrepararListView(ByVal pcTitle As String)
Dim Clm As ColumnHeader
    lvAnalistas.ListItems.Clear
    lvAnalistas.ColumnHeaders.Clear
    Set Clm = lvAnalistas.ColumnHeaders.Add(, , pcTitle, 4100)
    Clm.Alignment = lvwColumnLeft
    lvAnalistas.View = lvwReport
    lvAnalistas.CheckBoxes = True
    lvAnalistas.Gridlines = True
End Sub

