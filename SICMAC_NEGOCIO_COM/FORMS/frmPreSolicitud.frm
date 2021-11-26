VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreSolicitud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pre Solicitudes de Créditos"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmPreSolicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPresolicitudes 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   6855
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
      Height          =   360
      Left            =   5640
      TabIndex        =   2
      Top             =   4800
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      TabIndex        =   4
      Top             =   4800
      Width           =   1140
   End
   Begin MSComctlLib.ListView LstPresolicitud 
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   873
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
      Enabled         =   0   'False
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   531
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NOMBRE"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "OPCIÓN"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPreSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RGrid As ADODB.Recordset
'MARG ERS027-2017***
Dim RGrid2 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsFiltro As ADODB.Recordset
Dim MatPreSolicitud() As TPreSolicitud
Private Type TPreSolicitud
    cCodPresolicitud As String
    cPersNombre As String
    nPresolicitudId As Integer
End Type
'END MARG************

Dim cUserAnalista As String
Dim bAmpliado As Integer
Dim nPresolicitudId As Integer

'Public Function Inicio(ByVal pcUserAnalista As String, ByVal pbAmpliado As Integer) As Integer ' COMENTADO POR PTI1 ERS027-2017
Public Function Inicio(ByVal pcUserAnalista As String, ByVal pbAmpliado As Integer, Optional ByVal sPercod As String = "") As Integer ' AGREGADO POR PTI1 ERS027-2017
    Me.txtBuscar.Text = sPercod 'ADD PTI1
    
    nPresolicitudId = -1
    Screen.MousePointer = 0
    cUserAnalista = pcUserAnalista
    bAmpliado = pbAmpliado
    'Call CargarGrid 'COMENT BY MARG ERS027-2017
    'Call cargarGrid2 'ADD BY MARG ERS027-2017
    Call cargarGrid3
    
    'ADD PTI1 ERS027-2017
    If sPercod <> "" Then
        Call BuscarNombre
    End If 'END ANDD PTI1
    
    Me.Show 1
    Inicio = nPresolicitudId
    
   
End Function
'COMENTADO POR PTI1 24-01-2019
'Private Sub CargarGrid()
'    Dim oHojaRuta As COMDCredito.DCOMhojaRuta
'    Set oHojaRuta = New COMDCredito.DCOMhojaRuta
'    On Error GoTo ERRORCargaGrid
'    Set RGrid = oHojaRuta.ObtenerPreSolicitudes(cUserAnalista, bAmpliado)
''    Set DGSolicitudes.DataSource = RGrid 'comentado por pti1 24-01-2019
''    DGSolicitudes.Refresh 'comentado por pti1
'    Set fgPresolicitudes.DataSource = RGrid 'add pti1 24-01-2019
'    fgPresolicitudes.Refresh 'add pti1 24-01-2019
'    Set oHojaRuta = Nothing
'    Exit Sub
'ERRORCargaGrid:
'    MsgBox Err.Description, vbCritical, "Aviso"
'End Sub
'FIN COMENTADO POR PTI1

Private Sub cmdAceptar_Click()
'MARG ERS027-2017***
 Dim row As Long
row = Me.fgPresolicitudes.row
'END MARG************

    'If RGrid Is Nothing Then 'COMMENTED BY MARG ERS027-2017
    If RGrid2 Is Nothing Then 'ADD BY MARG ERS027-2017
            nPresolicitudId = -1
            Unload Me
            Exit Sub
    End If
'    If RGrid.RecordCount > 0 Then 'COMMENTED BY MARG ERS027-2017
    If RGrid2.RecordCount > 0 Then 'ADD BY MARG ERS027-2017
         'nPresolicitudId = RGrid.Fields(0) 'COMMENTED BY MARG ERS027-2017
         nPresolicitudId = fgPresolicitudes.TextMatrix(row, 3)
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
 nPresolicitudId = -1
 Unload Me
End Sub


Private Sub fgPresolicitudes_KeyPress(KeyAscii As Integer)
    Dim Col As Long
    Col = fgPresolicitudes.Col
    
     If KeyAscii = 13 And Col = 2 Then
        fgPresolicitudes_DblClick
    ElseIf KeyAscii = 13 And Col <> 2 Then
        EnfocaControl CmdAceptar
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
End Sub

'MARG ERS027-2017********************************************************
Private Sub fgPresolicitudes_DblClick()
  Dim row As Long
    Dim Col As Long
    row = fgPresolicitudes.row
    Col = fgPresolicitudes.Col
    If Col = 2 Then
          If RGrid2 Is Nothing Then
            nPresolicitudId = -1
            Unload Me
            Exit Sub
        End If
        If RGrid2.RecordCount > 0 Then
             nPresolicitudId = MatPreSolicitud(row).nPresolicitudId
             frmPreSolicitudRechazo.Inicio (nPresolicitudId)
             Call cargarGrid3
             Me.txtBuscar.Text = ""
             EnfocaControl Me.txtBuscar
        End If
    End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call BuscarNombre
    End If
End Sub
'PTI1 24-01-2019 ACTA 144-2018
'Private Sub fePresolicitud_DblClick()
'    Dim row As Long
'    Dim Col As Long
'    row = fePresolicitud.row
'    Col = fePresolicitud.Col
'    If Col = 3 Then
'          If RGrid2 Is Nothing Then
'            nPresolicitudId = -1
'            Unload Me
'            Exit Sub
'        End If
'        If RGrid2.RecordCount > 0 Then
'             nPresolicitudId = MatPreSolicitud(row).nPresolicitudId
'             frmPreSolicitudRechazo.Inicio (nPresolicitudId)
'             Call cargarGrid2
'             Me.txtBuscar.Text = ""
'
'        End If
'    End If
'End Sub
'COMENTADO POR PTI1 ACTA 144-2018 24-01-2019
'Private Sub cargarGrid2()
'    Dim oHojaRuta As COMDCredito.DCOMhojaRuta
'    Dim i As Integer
'    Dim ix As Integer
'
'    Set oHojaRuta = New COMDCredito.DCOMhojaRuta
'On Error GoTo ERRORCargaGrid
'   Set RGrid2 = Nothing
'    Set RGrid2 = oHojaRuta.ObtenerPreSolicitudes(cUserAnalista, bAmpliado)
'    Set rsFiltro = RGrid2.Clone
'
'    ReDim Preserve MatPreSolicitud(0)
'    Do While Not RGrid2.EOF
'        ReDim Preserve MatPreSolicitud(UBound(MatPreSolicitud) + 1)
'        MatPreSolicitud(UBound(MatPreSolicitud)).cCodPresolicitud = RGrid2!cCodPresolicitud
'        MatPreSolicitud(UBound(MatPreSolicitud)).cPersNombre = RGrid2!cPersNombre
'        MatPreSolicitud(UBound(MatPreSolicitud)).nPresolicitudId = RGrid2!nPresolicitudId
'        RGrid2.MoveNext
'    Loop
'
''    fePresolicitud.Clear
''    FormateaFlex fePresolicitud
'    LimpiaFlex Me.fePresolicitud
'    'si es coordinador o jefe de agencia, ocultamos la opcion de rechazo
'    If Not oHojaRuta.puedeRechazarPresolicitud(cUserAnalista) Then
'            fePresolicitud.ColWidth(2) = fePresolicitud.ColWidth(2) + fePresolicitud.ColWidth(3)
'            fePresolicitud.ColWidth(3) = 0
'    End If
'
'    With fePresolicitud
'    For i = 1 To UBound(MatPreSolicitud)
'        .AdicionaFila
'        .TextMatrix(i, 1) = MatPreSolicitud(i).cCodPresolicitud
'        .TextMatrix(i, 2) = MatPreSolicitud(i).cPersNombre
'        .TextMatrix(i, 3) = "Rechazar"
'        .TextMatrix(i, 4) = MatPreSolicitud(i).nPresolicitudId
'        .Col = 3: .CellForeColor = vbRed
'        .row = i
'        .CellForeColor
'        '.Font.Underline = True
'    Next
'    End With
'
'    Set oHojaRuta = Nothing
'    Exit Sub
'ERRORCargaGrid:
'    MsgBox Err.Description, vbCritical, "Aviso"
'FIN COMENTADO PTI1
Private Sub cargarGrid3()
    Dim oHojaRuta As COMDCredito.DCOMhojaRuta
    Dim i As Integer
    Dim ix As Integer

    Set oHojaRuta = New COMDCredito.DCOMhojaRuta
On Error GoTo ERRORCargaGrid
   Set RGrid2 = Nothing
    Set RGrid2 = oHojaRuta.ObtenerPreSolicitudes(cUserAnalista, bAmpliado)
    Set rsFiltro = RGrid2.Clone

    ReDim Preserve MatPreSolicitud(0)
    Do While Not RGrid2.EOF
        ReDim Preserve MatPreSolicitud(UBound(MatPreSolicitud) + 1)
        MatPreSolicitud(UBound(MatPreSolicitud)).cCodPresolicitud = RGrid2!cCodPresolicitud
        MatPreSolicitud(UBound(MatPreSolicitud)).cPersNombre = RGrid2!cPersNombre
        MatPreSolicitud(UBound(MatPreSolicitud)).nPresolicitudId = RGrid2!nPresolicitudId
        RGrid2.MoveNext
    Loop

'    fePresolicitud.Clear
'    FormateaFlex fePresolicitud
    'LimpiaFlex Me.fgPresolicitudes
    fgPresolicitudes.Clear
    With fgPresolicitudes
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "NOMBRE"
        .TextMatrix(0, 2) = "OPCION"
        .TextMatrix(0, 3) = "Id"
        .ColWidth(0) = 1700
        .ColWidth(1) = 4200
        .ColWidth(2) = 1700
        .ColWidth(3) = 0

        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter

        'si es coordinador o jefe de agencia, ocultamos la opcion de rechazo
        If Not oHojaRuta.puedeRechazarPresolicitud(cUserAnalista) Then
            .ColWidth(1) = .ColWidth(1) + .ColWidth(2)
            .ColWidth(2) = 0
        End If
    For i = 1 To UBound(MatPreSolicitud)
        .rows = i + 1
        .row = i
        .TextMatrix(i, 0) = MatPreSolicitud(i).cCodPresolicitud
        .TextMatrix(i, 1) = MatPreSolicitud(i).cPersNombre
        .TextMatrix(i, 2) = "Rechazar"
        .TextMatrix(i, 3) = MatPreSolicitud(i).nPresolicitudId
        .Col = 2
        .row = i
        .CellFontUnderline = True
    Next
    End With

    fgPresolicitudes.row = 1
    fgPresolicitudes.Col = 1
    fgPresolicitudes.RowSel = 1
'    EnfocaControl Me.txtBuscar

    Set oHojaRuta = Nothing
    Exit Sub
ERRORCargaGrid:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub


Private Sub txtBuscar_Change()
    If txtBuscar.Text = "" Then
        Call cargarGrid3
    End If
End Sub

Private Sub BuscarNombre()
    Dim oHojaRuta As COMDCredito.DCOMhojaRuta
    Dim i As Integer
    
    Set oHojaRuta = New COMDCredito.DCOMhojaRuta
    If txtBuscar.Text <> "" Then
            rsFiltro.Filter = "cPersNombre like '*" + txtBuscar.Text + "*'"
    
            If Not rsFiltro.EOF And Not rsFiltro.BOF Then
               fgPresolicitudes.Clear
                fgPresolicitudes.Enabled = True
                
                ReDim Preserve MatPreSolicitud(0)
                Do While Not rsFiltro.EOF
                    ReDim Preserve MatPreSolicitud(UBound(MatPreSolicitud) + 1)
                    MatPreSolicitud(UBound(MatPreSolicitud)).cCodPresolicitud = rsFiltro!cCodPresolicitud
                    MatPreSolicitud(UBound(MatPreSolicitud)).cPersNombre = rsFiltro!cPersNombre
                    MatPreSolicitud(UBound(MatPreSolicitud)).nPresolicitudId = rsFiltro!nPresolicitudId
                    rsFiltro.MoveNext
                Loop
                 With fgPresolicitudes
                    .TextMatrix(0, 0) = "Código"
                    .TextMatrix(0, 1) = "NOMBRE"
                    .TextMatrix(0, 2) = "OPCION"
                    .TextMatrix(0, 3) = "Id"
                    .ColWidth(0) = 1700
                    .ColWidth(1) = 4200
                    .ColWidth(2) = 1700
                    .ColWidth(3) = 0
                    
                    .ColAlignment(0) = flexAlignCenterCenter
                    .ColAlignment(2) = flexAlignCenterCenter
                    
                    'si es coordinador o jefe de agencia, ocultamos la opcion de rechazo
                    If Not oHojaRuta.puedeRechazarPresolicitud(cUserAnalista) Then
                        .ColWidth(1) = .ColWidth(1) + .ColWidth(2)
                        .ColWidth(2) = 0
                    End If
                    For i = 1 To UBound(MatPreSolicitud)
                        .rows = i + 1
                        .row = i
                        .TextMatrix(i, 0) = MatPreSolicitud(i).cCodPresolicitud
                        .TextMatrix(i, 1) = MatPreSolicitud(i).cPersNombre
                        .TextMatrix(i, 2) = "Rechazar"
                        .TextMatrix(i, 3) = MatPreSolicitud(i).nPresolicitudId
                        .Col = 2
                        .row = i
                        .CellFontUnderline = True
                    Next
                End With
                fgPresolicitudes.row = 1
                fgPresolicitudes.Col = 1
                fgPresolicitudes.RowSel = 1
                EnfocaControl fgPresolicitudes
            End If
    Else
        Call cargarGrid3
    End If
End Sub


'END MARG ***

