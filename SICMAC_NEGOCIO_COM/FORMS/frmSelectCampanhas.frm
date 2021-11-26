VERSION 5.00
Begin VB.Form frmSelectCampanhas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Campaña"
   ClientHeight    =   5670
   ClientLeft      =   3930
   ClientTop       =   1410
   ClientWidth     =   5835
   Icon            =   "frmSelectCampanhas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   60
      TabIndex        =   5
      Top             =   315
      Width           =   5685
      Begin VB.OptionButton OptCampanha 
         Caption         =   "&Ninguno"
         Height          =   210
         Index           =   1
         Left            =   2565
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton OptCampanha 
         Caption         =   "&Todos"
         Height          =   210
         Index           =   0
         Left            =   855
         TabIndex        =   3
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
      Left            =   2040
      TabIndex        =   2
      Top             =   5190
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analistas"
      Height          =   4260
      Left            =   60
      TabIndex        =   0
      Top             =   855
      Width           =   5685
      Begin VB.ListBox LstCampanha 
         Height          =   3885
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   285
         Width           =   5505
      End
   End
End
Attribute VB_Name = "frmSelectCampanhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SeleccionaCampanhas()
    Call CargaCampanhas
    Me.Caption = "Seleccionar Campaña"
    OptCampanha(1).value = True
    OptCampanha_Click 1
    Me.Show 1
End Sub

Private Sub CargaCampanhas()
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
    'Set rs = New ADODB.Recordset
    Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        sSql = "SELECT cdescripcion,idcampana from CAMPANAS order by idcampana"
    Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
    Set oConecta = Nothing
        
    LstCampanha.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstCampanha.AddItem rs!cDescripcion & Space(100) & rs!idcampana
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
'peac 20071025
Public Sub SeleccionaGastos()
    Call CargaGastos
    Me.Caption = "Seleccionar Gastos"
    OptCampanha(1).value = True
    OptCampanha_Click 1
    Me.Show 1
End Sub

Private Sub CargaGastos()
Dim rs As ADODB.Recordset
Dim oConecta As COMConecta.DCOMConecta
Dim sSql As String
    
    Set oConecta = New COMConecta.DCOMConecta
        oConecta.AbreConexion
        
'        sSql = "SELECT cdescripcion,idcampana from CAMPANAS order by idcampana"
        
sSql = "select nPrdConceptoCod,cDescripcion from ProductoConcepto where nPrdConceptoCod like '12__'"
sSql = sSql & " and nPrdConceptoCod not in (1200,1201,1202,1203,1204,1205,1206,1207,1208,1209,1210,1211,1212,1213,1214,1215,1216,1217,"
sSql = sSql & " 1221,1223,1224,1231,1232,1299) order by nPrdConceptoCod"
        
    Set rs = oConecta.CargaRecordSet(sSql)
        oConecta.CierraConexion
    Set oConecta = Nothing
        
    LstCampanha.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstCampanha.AddItem rs!cDescripcion & Space(100) & rs!nPrdConceptoCod
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Sub CmdAceptar_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    CentraForm Me
    OptCampanha(1).value = True
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub OptCampanha_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
    If LstCampanha.ListCount <= 0 Then
        Exit Sub
    End If
    For i = 0 To LstCampanha.ListCount - 1
        LstCampanha.Selected(i) = bCheck
    Next i
End Sub

