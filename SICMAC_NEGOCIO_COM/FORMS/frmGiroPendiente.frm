VERSION 5.00
Begin VB.Form frmGiroPendiente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BUSCAR GIROS"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13035
   Icon            =   "frmGiroPendiente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4920
      Width           =   1100
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Giros Pendientes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12795
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   5940
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtnumdoc 
         Height          =   375
         Left            =   3800
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin SICMACT.FlexEdit grdGiro1 
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   6588
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cuenta-Apertura-Remitente-Monto-Destinatario-Agencia Dest.-ValPinPad"
         EncabezadosAnchos=   "400-2000-1200-3900-1100-3900-1600-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-1-1-1-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.ComboBox cboTipoDoi 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo DOI:"
         Height          =   255
         Left            =   140
         TabIndex        =   7
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese DOI:"
         Height          =   255
         Left            =   2750
         TabIndex        =   6
         Top             =   330
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmGiroPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ValidaDatosDocumentos As Boolean
Public nDocAnt As String
Dim nFila As Long

Dim oFrmPadre As Form 'RECO20140415 ERS008-2014

Private Sub CargaGirosPendientes(ByVal pDocNum As String)
Dim rsGiro As Recordset
Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios

Dim i As Integer
 
For i = 1 To nFila
      grdGiro1.EliminaFila (i)
Next i
nFila = 0
'grdGiro1.EncabezadosNombres
Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsGiro = clsGiro.GetGiroPendientexDoc(gsCodAge, pDocNum)
    If Not (rsGiro.EOF And rsGiro.BOF) Then
    
        Do While Not rsGiro.EOF
            If grdGiro1.TextMatrix(1, 1) <> "" Then grdGiro1.rows = grdGiro1.rows + 1
            nFila = grdGiro1.rows - 1
            grdGiro1.TextMatrix(nFila, 0) = nFila
            grdGiro1.TextMatrix(nFila, 1) = rsGiro("cCtaCod")
            grdGiro1.TextMatrix(nFila, 2) = Format$(rsGiro("dPrdEstado"), "dd/mm/yyyy")
            grdGiro1.TextMatrix(nFila, 3) = PstaNombre(rsGiro("cRemitente"), False)
            grdGiro1.TextMatrix(nFila, 4) = Format$(rsGiro("nSaldo"), "#,##0.00")
            grdGiro1.TextMatrix(nFila, 5) = rsGiro("cDestinatario")
            grdGiro1.TextMatrix(nFila, 6) = Trim(rsGiro("cAgencia"))
            grdGiro1.TextMatrix(nFila, 7) = Trim(rsGiro("nClavePinPad"))
            rsGiro.MoveNext
        Loop
    Else
        MsgBox "No hay giros pendientes registrados.", vbInformation, "Aviso"
        'cmdAceptar.Enabled = False RECO20170716
        'fraDatos.Enabled = False RECO20140716
        Call LimpiarFormulario 'RECO20140716
    End If
    
End Sub


'Comentado x MADM 20101011
'Private Sub CargaGirosPendientes()
'Dim rsGiro As Recordset
'Dim clsGiro As COMNCaptaServicios.NCOMCaptaServicios
'Dim nFila As Long
'Set clsGiro = New COMNCaptaServicios.NCOMCaptaServicios
'Set rsGiro = clsGiro.GetGiroPendiente(gsCodAge)
'If Not (rsGiro.EOF And rsGiro.BOF) Then
'    Do While Not rsGiro.EOF
'        If grdGiro1.TextMatrix(1, 1) <> "" Then grdGiro1.Rows = grdGiro1.Rows + 1
'        nFila = grdGiro1.Rows - 1
'        grdGiro1.TextMatrix(nFila, 0) = nFila
'        grdGiro1.TextMatrix(nFila, 1) = rsGiro("cCtaCod")
'        grdGiro1.TextMatrix(nFila, 2) = Format$(rsGiro("dPrdEstado"), "dd/mm/yyyy")
'        grdGiro1.TextMatrix(nFila, 3) = PstaNombre(rsGiro("cRemitente"), False)
'        grdGiro1.TextMatrix(nFila, 4) = Format$(rsGiro("nSaldo"), "#,##0.00")
'        grdGiro1.TextMatrix(nFila, 5) = rsGiro("cDestinatario")
'        grdGiro1.TextMatrix(nFila, 6) = Trim(rsGiro("cAgencia"))
'        rsGiro.MoveNext
'    Loop
'Else
'    MsgBox "No hay giros pendientes registrados.", vbInformation, "Aviso"
'    cmdAceptar.Enabled = False
'    fraDatos.Enabled = False
'End If
'End Sub

'Comentado x MADM 20100524
'Private Sub MergeGrid()
'Dim i As Integer
'For i = 0 To grdGiro.Cols - 1
'    grdGiro.MergeCol(i) = True
'Next i
'grdGiro.MergeCells = flexMergeFree
'grdGiro.BandExpandable(0) = True
'End Sub

'Private Sub SetupGrid()
'With grdGiro
'    .Clear
'    .Cols = 7
'    .Rows = 2
'    .TextMatrix(0, 0) = "#"
'    .TextMatrix(0, 1) = "Cuenta"
'    .TextMatrix(0, 2) = "Apertura"
'    .TextMatrix(0, 3) = "Remitente"
'    .TextMatrix(0, 4) = "Monto"
'    .TextMatrix(0, 5) = "Destinatario"
'    .TextMatrix(0, 6) = "Agencia Dest."
'    .ColWidth(0) = 350
'    .ColWidth(1) = 2000
'    .ColWidth(2) = 1200
'    .ColWidth(3) = 3900
'    .ColWidth(4) = 1100
'    .ColWidth(5) = 3900
'    .ColWidth(6) = 1600
'    .ColAlignment(0) = 4
'    .ColAlignment(1) = 4
'    .ColAlignment(2) = 4
'    .ColAlignment(3) = 1
'    .ColAlignment(4) = 7
'    .ColAlignment(5) = 1
'    .ColAlignment(6) = 4
'    .ColAlignmentFixed(0) = 4
'    .ColAlignmentFixed(1) = 4
'    .ColAlignmentFixed(2) = 4
'    .ColAlignmentFixed(3) = 4
'    .ColAlignmentFixed(4) = 4
'    .ColAlignmentFixed(5) = 4
'    .ColAlignmentFixed(6) = 4
'End With
'End Sub

Private Sub CmdAceptar_Click()
Dim nFila As Long
nFila = grdGiro1.row
With oFrmPadre 'RECO20140415 ERS008-2014
'    .txtCuenta.Age = Mid(grdGiro.TextMatrix(nFila, 1), 4, 2)
'    .txtCuenta.cuenta = Right(grdGiro.TextMatrix(nFila, 1), 10)
    .txtCuenta.Age = Mid(grdGiro1.TextMatrix(nFila, 1), 4, 2)
    .txtCuenta.Cuenta = Right(grdGiro1.TextMatrix(nFila, 1), 10)
    .lnValPinPad = grdGiro1.TextMatrix(nFila, 7)
    
End With
Unload Me
End Sub

Private Sub CmdBuscar_Click()
    Dim j As Integer
    '20180611 NRLO INC1804170006
    Dim sDoi As String
    Dim cTipo As String
    sDoi = txtnumdoc.Text
    cTipo = CInt(Trim(Right(cboTipoDoi.Text, 2)))
    'END 20180611 NRLO INC1804170006
    
    If Len(txtnumdoc.Text) = 0 Then
            MsgBox "Falta Ingresar el Número de Documento", vbInformation, "Aviso"
            txtnumdoc.SetFocus
            ValidaDatosDocumentos = False
            Call LimpiarFormulario
            Exit Sub
    End If
    
    'If Len(txtnumdoc.Text) <> gnNroDigitosDNI Then  20180611 NRLO INC1804170006
    '      MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
    '      txtnumdoc.SetFocus
    '      ValidaDatosDocumentos = False
    '      Call LimpiarFormulario
    '      Exit Sub
    'End If END 20180611 NRLO INC1804170006
    
    '20180611 NRLO INC1804170006
    If cTipo = 1 And Len(sDoi) <> gnNroDigitosDNI Then
        MsgBox "DOI No es de " & gnNroDigitosDNI & " dígitos", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        Exit Sub
    ElseIf cTipo = 2 And Len(sDoi) <> 11 Then
        MsgBox "DOI No es de 11 digitos", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        Exit Sub
    ElseIf cTipo = 4 And (Len(sDoi) > 12 Or Len(sDoi) < 5) Then
        MsgBox "DOI debe tener entre 5 y 12 dígitos", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        Exit Sub
    End If
    'END 20180611 NRLO INC1804170006
    
    If cTipo = 1 Then
        For j = 1 To Len(Trim(txtnumdoc.Text))
            If (Mid(txtnumdoc.Text, j, 1) < "0" Or Mid(txtnumdoc.Text, j, 1) > "9") Then
               MsgBox "Uno de los Digitos del DNI no es un Numero", vbInformation, "Aviso"
               txtnumdoc.SetFocus
               ValidaDatosDocumentos = False
               Call LimpiarFormulario
               Exit Sub
            End If
        Next j
    End If
    'nDocAnt = Me.txtnumdoc
    cmdAceptar.Enabled = True 'RECO20140717
    CargaGirosPendientes Me.txtnumdoc
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
ValidaDatosDocumentos = True
'nVeces = 0
'SetupGrid
'CargaGirosPendientes 'COMENTADO X MADM 20101011
'MergeGrid
Call CargarDOIs '20180611 NRLO FIN INC1804170006
End Sub
Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
    Dim cTipo As String
    cTipo = CInt(Trim(Right(cboTipoDoi.Text, 2)))
    
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If

    If cTipo = 1 Then
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
    If cTipo = 4 Then
        KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    End If
End Sub
'20180611 NRLO FIN INC1804170006
'RECO20140415 ERS008-2014********************************************
Public Sub Inicio(ByVal oFrmcontrol As Form)
    Set oFrmPadre = oFrmcontrol
    Me.Show 1
End Sub
'RECO FIN ***********************************************************
'RECO20140716*******************************************************
Public Sub LimpiarFormulario()
    txtnumdoc.Text = ""
    grdGiro1.Clear
    FormateaFlex grdGiro1
    txtnumdoc.Enabled = True
    cmdAceptar.Enabled = False
End Sub
'RECO FIN***********************************************************

Private Sub CargarDOIs() '20180611 NRLO INC1804170006
    Dim lsDOIs As ADODB.Recordset
    Dim lsDOIFinal As ADODB.Recordset
    Dim oConstante As COMDConstantes.DCOMConstantes
    Set oConstante = New COMDConstantes.DCOMConstantes
    Set lsDOIs = oConstante.RecuperaConstantes(1003)
    
    Set lsDOIFinal = New ADODB.Recordset
    lsDOIFinal.Fields.Append "nConsValor", adInteger
    lsDOIFinal.Fields.Append "cConsDescripcion", adBSTR, 10, adFldUpdatable
    lsDOIFinal.Open
    
    Do While Not lsDOIs.EOF
        If (lsDOIs!nConsValor = 1 Or lsDOIs!nConsValor = 4) Then
            lsDOIFinal.AddNew
            lsDOIFinal.Fields!cConsDescripcion = lsDOIs!cConsDescripcion
            lsDOIFinal.Fields!nConsValor = lsDOIs!nConsValor
            lsDOIFinal.Update
            lsDOIFinal.MoveFirst
        End If
        lsDOIs.MoveNext
    Loop
    lsDOIs.Close

    cboTipoDoi.Clear
    Call Llenar_Combo_con_Recordset(lsDOIFinal, Me.cboTipoDoi)
    cboTipoDoi.ListIndex = 0
    
    txtnumdoc.MaxLength = 8
End Sub '20180611 NRLO FIN INC1804170006

Private Sub cboTipoDoi_Click()
    '20180611 NRLO INC1804170006
    Dim cTipo As String
    cTipo = CInt(Trim(Right(cboTipoDoi.Text, 2)))
    
    If cTipo = 1 Then
        txtnumdoc.Text = "" 'Left(txtnumdoc.Text, 8)
        txtnumdoc.MaxLength = 8
    ElseIf cTipo = 2 Then
        txtnumdoc.Text = Left(txtnumdoc.Text, 11)
        txtnumdoc.MaxLength = 11
    ElseIf cTipo = 4 Then
        txtnumdoc.Text = "" 'Left(txtnumdoc.Text, 12)
        txtnumdoc.MaxLength = 12
    Else
        txtnumdoc.MaxLength = 20
    End If '20180611 NRLO FIN INC1804170006
End Sub
