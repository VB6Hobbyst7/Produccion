VERSION 5.00
Begin VB.Form frmSelectAnalistas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Analistas"
   ClientHeight    =   5670
   ClientLeft      =   3930
   ClientTop       =   1410
   ClientWidth     =   6015
   Icon            =   "frmSelectAnalistas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
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
      ScaleWidth      =   330
      TabIndex        =   7
      Top             =   330
      Width           =   330
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
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
      ScaleWidth      =   5940
      TabIndex        =   6
      Top             =   45
      Width           =   5940
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   390
      TabIndex        =   5
      Top             =   315
      Width           =   5595
      Begin VB.OptionButton OptAnalista 
         Caption         =   "&Ninguno"
         Height          =   210
         Index           =   1
         Left            =   2565
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton OptAnalista 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   5190
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analistas"
      Height          =   4260
      Left            =   420
      TabIndex        =   0
      Top             =   855
      Width           =   5565
      Begin VB.ListBox LstAnalista 
         Height          =   3885
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   285
         Width           =   5385
      End
   End
End
Attribute VB_Name = "frmSelectAnalistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SeleccionaInstituciones()
    Call CargaInstituciones
    Me.Caption = "Seleccionar Instituciones"
    Me.Show 1
End Sub


Public Sub SeleccionaAnalistas()
    Call CargaAnalistas
    Me.Caption = "Seleccionar Analista"
    OptAnalista(1).value = True
    OptAnalista_Click 1
    Me.Show 1
End Sub

Private Sub CargaInstituciones()
Dim rs As New ADODB.Recordset
Dim oGen  As COMDConstSistema.DCOMGeneral
    
    On Error GoTo ERRORCargaAnalistas
    Set oGen = New COMDConstSistema.DCOMGeneral
    Set rs = oGen.CargaInstituciones()
    Set oGen = Nothing
    LstAnalista.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstAnalista.AddItem rs!cPersNombre & Space(100) & rs!cPersCod
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub



Private Sub CargaAnalistas()
Dim rs As ADODB.Recordset
Dim sAnalistas As String
Dim oGen  As COMDConstSistema.DCOMGeneral

    On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New COMDConstSistema.DCOMGeneral
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaAnalistas()
    Set oGen = Nothing
    LstAnalista.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            LstAnalista.AddItem rs!cPersNombre & Space(100) & rs!cPersCod
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdAceptar_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    CentraForm Me
    OptAnalista(1).value = True
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub OptAnalista_Click(Index As Integer)
Dim bCheck As Boolean
Dim i As Integer
    If Index = 0 Then
        bCheck = True
    Else
        bCheck = False
    End If
    If LstAnalista.ListCount <= 0 Then
        Exit Sub
    End If
    For i = 0 To LstAnalista.ListCount - 1
        LstAnalista.Selected(i) = bCheck
    Next i
End Sub
