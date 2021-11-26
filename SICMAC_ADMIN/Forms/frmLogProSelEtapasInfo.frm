VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLogProSelEtapasInfo 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado del Proceso de Seleccion"
   ClientHeight    =   6780
   ClientLeft      =   2775
   ClientTop       =   2025
   ClientWidth     =   6240
   Icon            =   "frmLogProSelEtapasInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   6015
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Doble Click para Ver Comite"
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   10610
      _Version        =   393216
      BackColor       =   -2147483644
      ForeColor       =   -2147483644
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483629
      BackColorSel    =   -2147483633
      ForeColorSel    =   -2147483644
      BackColorUnpopulated=   -2147483644
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483644
      GridColorUnpopulated=   -2147483644
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLinesUnpopulated=   1
      ScrollBars      =   2
      SelectionMode   =   1
      MergeCells      =   2
      BorderStyle     =   0
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Image imgFlecha 
      Height          =   240
      Index           =   1
      Left            =   8160
      Picture         =   "frmLogProSelEtapasInfo.frx":08CA
      Top             =   3240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNivel 
      Height          =   315
      Index           =   0
      Left            =   8160
      Top             =   2820
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgFlecha 
      Height          =   360
      Index           =   0
      Left            =   8100
      Picture         =   "frmLogProSelEtapasInfo.frx":0C0C
      Top             =   2400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgXX 
      Height          =   240
      Left            =   8100
      Picture         =   "frmLogProSelEtapasInfo.frx":130E
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   8100
      Picture         =   "frmLogProSelEtapasInfo.frx":1650
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogProSelEtapasInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gnProSelNro As Integer

Public Sub Inicio(ByVal pnProSelNro As Integer)
    gnProSelNro = pnProSelNro
    Me.Show 1
End Sub

Private Sub cmdComite_Click()
    MSFlex_DblClick
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    CargarEtapas gnProSelNro
End Sub

Private Sub FormaFlex()
    With MSFlex
        .Rows = 1
        .RowHeight(0) = 0
        .TextMatrix(0, 0) = "Etapa":            .ColWidth(0) = 4775:            .CellAlignment = 4
        .TextMatrix(0, 1) = "Estado":           .ColWidth(1) = 735:             .CellPictureAlignment = 4
        .TextMatrix(0, 2) = "Codigo":           .ColWidth(2) = 0:
        .ForeColor = &H80000008:                .BackColor = &H80000005
    End With
End Sub

Private Sub CargarEtapas(ByVal pnProSelNro As Integer)
On Error GoTo CargarEtapasErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, _
        bColor As Boolean, nColor As Long, i As Integer, bFlecha As Boolean, dFechas As String
    Set oCon = New DConecta
    bColor = True
    If oCon.AbreConexion Then
       sSQL = "select e.*, t.cDescripcion cEtapa " & _
                " From LogProSelEtapa e  " & _
                "     inner join LogEtapa t on t.nEstado = 1 and e.nEtapaCod = t.nEtapaCod " & _
                " Where e.nEstado>0 and e.nProSelNro = " & pnProSelNro & "  "
        Set Rs = oCon.CargaRecordSet(sSQL)
        FormaFlex
        With MSFlex
        If Not Rs.EOF Then
            Do While Not Rs.EOF
                
                If bFlecha Then
                    i = i + 1
                    InsRow MSFlex, i
                    .Col = 0
                    .row = i
                    .CellBackColor = &H8000000A
                    .RowHeight(i) = 300
                    .CellPictureAlignment = 4
                    Set .CellPicture = imgFlecha(1)
                    .Col = 1
                    .CellBackColor = &H8000000A
                End If
                
                If bColor Then
                    nColor = &HEAFFFF
                Else
                    nColor = &HECFFEF
                End If
                
                If Rs!dFechaInicio <> "" Then
                    dFechas = "DEL " & Rs!dFechaInicio & " AL " & Rs!dFechaTermino
                Else
                    dFechas = ""
                End If
                
                i = i + 1
                InsRow MSFlex, i
                .RowHeight(i) = 735:
                .TextMatrix(i, 0) = Rs!cEtapa & vbCrLf & dFechas
                
                .Col = 0: .row = i
                .CellBackColor = nColor
                .CellAlignment = 4
                .Col = 1
                .CellPictureAlignment = 4
                If Rs!nEstado = 2 Then ' Or rs!dFechaTermino < gdFecSis Then
                    Set .CellPicture = imgOK
                Else
                    Set .CellPicture = imgXX
                End If
                .TextMatrix(i, 2) = Rs!nEtapaCod
                
                bColor = Not bColor
                bFlecha = True
                Rs.MoveNext
            Loop
            
'            i = i + 1
'            InsRow MSFlex, i
'            .Col = 0
'            .Row = i
'            .CellBackColor = &H8000000A
'            .RowHeight(i) = 300
'            .CellPictureAlignment = 4
'            Set .CellPicture = imgFlecha(1)
'            .Col = 1
'            .CellBackColor = &H8000000A
'
'            i = i + 1
'            InsRow MSFlex, i
'            .RowHeight(i) = 735:
'            .TextMatrix(i, 0) = "yo" & vbCrLf & dFechas
        End If
        End With
        oCon.CierraConexion
    End If
'    .RowHeight(1) = 735:        .TextMatrix(1, 0) = "Venta de Bases"
'    .Col = 0: .Row = 1
'    .CellBackColor = &HEAFFFF
'    .Col = 1
'    .CellPictureAlignment = 4
'    Set .CellPicture = imgXX
'
'    .Col = 0
'    .Row = 2
'    .RowHeight(2) = 300
'    .CellPictureAlignment = 4
'    Set .CellPicture = imgFlecha(1)
'
'    .CellAlignment = 4
'    .RowHeight(3) = 735:        .TextMatrix(3, 0) = "Consultas"
'    .Col = 0: .Row = 3
'    .CellBackColor = &HECFFEF
'    .Col = 1
'    .CellPictureAlignment = 4
'    Set .CellPicture = imgOK
    Exit Sub
CargarEtapasErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Function ComiteEtapa(ByVal pnProSelNro As Integer, ByVal pnEtapaCod As Integer) As String
On Error GoTo comiteEtapaErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select x.cCargo, e.nEstado, bSuplente, cPersNombre=replace(p.cPersNombre,'/',' '), e.dFechaInicio, e.dFechaTermino, t.cDescripcion cEtapa " & _
                " from LogProSelComite c " & _
                " inner join LogProSelEtapa e on c.nProSelNro = e.nProSelNro " & _
                " inner join LogEtapa t on e.nEtapaCod = t.nEtapaCod and t.nEstado = 1 " & _
                " inner join persona p on c.cPersCod = p.cPersCod " & _
                " inner join LogProSelEtapaComite y on c.nProSelNro = y.nProSelNro and p.cPersCod= y.cPersCod and e.nEtapaCod = Y.nEtapaCod " & _
                " inner join (select nConsValor as nCargo, cConsDescripcion as cCargo from Constante where nConsCod= 9085 and nConsCod<>nConsValor) x on c.nCargo = x.nCargo " & _
                " Where e.nEstado>0 and c.nProselNro = " & pnProSelNro & " and t.nEtapaCod=" & pnEtapaCod
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            i = i + 1
            ComiteEtapa = ComiteEtapa & Rs!cCargo & IIf(Rs!bSuplente, " SUPLENTE", " TITULAR") & ": " & Rs!cPersNombre & vbCrLf
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Set Rs = Nothing
    Exit Function
comiteEtapaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelEtapasInfo = Nothing
End Sub

Private Sub MSFlex_DblClick()
    Dim sComite As String, sEtapa As String, nPos As Integer, nSpace As Integer
    
    If Val(MSFlex.TextMatrix(MSFlex.row, 2)) = 0 Then Exit Sub
    
    nPos = InStr(1, MSFlex.TextMatrix(MSFlex.row, 0), vbCrLf)
    sEtapa = Mid(MSFlex.TextMatrix(MSFlex.row, 0), 1, nPos)
    sComite = Space(23) & sEtapa & Space(23) & vbCrLf
    sComite = sComite & ComiteEtapa(gnProSelNro, Val(MSFlex.TextMatrix(MSFlex.row, 2)))
    MsgBox sComite, vbInformation, "Comite Encargado"
End Sub
