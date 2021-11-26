VERSION 5.00
Begin VB.Form frmCajeroOpeEgreRef 
   Caption         =   "Egresos por Ope.Tramite de Refinanción y Otras Ope.Cred."
   ClientHeight    =   5715
   ClientLeft      =   4050
   ClientTop       =   2670
   ClientWidth     =   6960
   Icon            =   "frmCajeroOpeEgreRef.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6960
   Begin VB.TextBox txtGlosa 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   60
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3270
      Width           =   6780
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   5340
      TabIndex        =   11
      Top             =   5235
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   3900
      TabIndex        =   10
      Top             =   5235
      Width           =   1455
   End
   Begin VB.OptionButton OptMon 
      Caption         =   "Dolares"
      Height          =   315
      Index           =   1
      Left            =   5670
      TabIndex        =   1
      Top             =   60
      Width           =   900
   End
   Begin VB.OptionButton OptMon 
      Caption         =   "Soles"
      Height          =   315
      Index           =   0
      Left            =   4755
      TabIndex        =   0
      Top             =   60
      Value           =   -1  'True
      Width           =   765
   End
   Begin SICMACT.FlexEdit fgOpe 
      Height          =   2100
      Left            =   45
      TabIndex        =   4
      Top             =   915
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3704
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Fecha-Referencia-Importe-nMovNro-nMoneda"
      EncabezadosAnchos=   "500-1500-3000-1200-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-R-L-L"
      FormatosEdit    =   "0-0-0-2-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin SICMACT.TxtBuscar txtCodCli 
      Height          =   345
      Left            =   900
      TabIndex        =   2
      Top             =   105
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   3
      sTitulo         =   ""
      TipoBusPers     =   1
   End
   Begin VB.Label lblMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5025
      TabIndex        =   6
      Top             =   4005
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monto :"
      Height          =   195
      Left            =   4335
      TabIndex        =   15
      Top             =   4050
      Width           =   540
   End
   Begin VB.Label lblITF 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5010
      TabIndex        =   7
      Top             =   4380
      Width           =   1755
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "ITF :"
      Height          =   195
      Left            =   4320
      TabIndex        =   14
      Top             =   4410
      Width           =   330
   End
   Begin VB.Label lblGlosa 
      Caption         =   "Glosa"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   75
      TabIndex        =   13
      Top             =   3060
      Width           =   915
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5175
      TabIndex        =   8
      Top             =   4785
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3945
      TabIndex        =   9
      Top             =   4740
      Width           =   2865
   End
   Begin VB.Label lblNomCli 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   900
      TabIndex        =   3
      Top             =   450
      Width           =   5790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      Height          =   195
      Left            =   195
      TabIndex        =   12
      Top             =   150
      Width           =   570
   End
End
Attribute VB_Name = "frmCajeroOpeEgreRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bOpeAfecta As Boolean
Dim lsOpeCod As Long
Dim lsCaption As String
Dim nRedondeoITF As Double
Public Sub Inicia(ByVal pnOpeCod As Long, ByVal psOpeDesc As String)
Dim oITF As New COMDConstSistema.FCOMITF

lsOpeCod = pnOpeCod
lsCaption = psOpeDesc
bOpeAfecta = oITF.VerifOpeVariasAfectaITF(Str(lsOpeCod))
Set oITF = Nothing
Me.Caption = lsCaption
Me.Show 1
End Sub
Private Sub cmdCancelar_Click()
txtCodCli.Enabled = True
txtCodCli = ""
lblnomcli = ""
lblTotal = "0.00"
lblMonto = "0.00"
lblITF = "0.00"
fgOpe.Clear
fgOpe.Rows = 2
fgOpe.FormaCabecera
optMon(0).Enabled = True
optMon(1).Enabled = True
nRedondeoITF = 1
End Sub

Private Sub CmdGrabar_Click()
Dim CodOpe As String
Dim lnMonto As Currency
Dim Moneda As String
Dim lsMov As String
Dim lsMovITF As String
Dim lsDocumento As String
Dim lnMovNro As Long
Dim lnMovNroITF As Long
Dim lbBan As Boolean
    
Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim clsCont As COMNContabilidad.NCOMContFunciones
Dim oCapMov As COMDCaptaGenerales.DCOMCaptaMovimiento

Dim ofun As New COMDConstSistema.DCOMGeneral
Dim oITF As New COMDConstSistema.FCOMITF
Dim loMov As COMDMov.DCOMMov

Dim i As Integer
Dim pnMovNroRef As Long

Set oCapMov = New COMDCaptaGenerales.DCOMCaptaMovimiento

Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
Set clsCont = New COMNContabilidad.NCOMContFunciones
Set loMov = New COMDMov.DCOMMov

'On Error GoTo Error

lnMonto = CCur(lblTotal.Caption)
lsMov = ofun.FechaHora(gdFecSis)
Set ofun = Nothing
lsDocumento = ""
If Len(lblnomcli.Caption) = 0 Then
    MsgBox "Selecciona a un Cliente por favor", vbInformation, "Aviso"
    txtCodCli.SetFocus
    Exit Sub
ElseIf Len(Trim(txtGlosa.Text)) = 0 Then
    MsgBox "Ingrese la glosa o comentario correspondiente", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    lnMovNro = clsCapMov.OtrasOperaciones(lsMov, lsOpeCod, lnMonto, lsDocumento, Me.txtGlosa.Text, IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), txtCodCli.Text)
    
    'agregamos los mov de referencia de las operaciones
    For i = 1 To fgOpe.Rows - 1
        pnMovNroRef = fgOpe.TextMatrix(i, 4)
        oCapMov.AgregaMovRef lnMovNro, pnMovNroRef
    Next
    If oITF.gbITFAplica And CCur(Me.lblITF.Caption) > 0 Then
        lsMovITF = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser, lsMov)
        lnMovNroITF = clsCapMov.OtrasOperaciones(lsMovITF, gITFCobroEfectivo, Me.lblITF.Caption, lsDocumento, Me.txtGlosa.Text, IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), Me.txtCodCli.Text)
        '*** BRGO 20110906 ***************************
           Call loMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.lblITF) + nRedondeoITF, CCur(Me.lblITF))
           Set loMov = Nothing
        '*** BRGO
    End If
    
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim oBolITF As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim lsBoleta As String
    Dim lsBoletaITF As String
    Dim nFicSal As Integer
    
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
       lsBoleta = oBol.ImprimeBoleta("OTRAS OPERACIONES", Left(lsCaption, 15), "", Str(lnMonto), lblnomcli.Caption, "________" & IIf(Me.optMon(0).value = True, gMonedaNacional, gMonedaExtranjera), lsDocumento, 0, "0", IIf(Len(lsDocumento) = 0, "", "Nro Documento"), 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False)
    Set oBol = Nothing
    
    Set oBolITF = New COMNCaptaGenerales.NCOMCaptaMovimiento
        If oITF.gbITFAplica And CCur(Me.lblITF.Caption) > 0 Then
            lsBoletaITF = oBolITF.fgITFImprimeBoleta(lblnomcli.Caption, CCur(Me.lblITF.Caption), Me.Caption, lnMovNroITF, sLpt, , , , , , , False, , , , 0, 0, , "")
        End If
    Set oBolITF = Nothing
    Do
       
        If Trim(lsBoleta) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta
                Print #nFicSal, ""
            Close #nFicSal
        End If
          
        If Trim(lsBoletaITF) <> "" Then
          nFicSal = FreeFile
          Open sLpt For Output As nFicSal
              Print #nFicSal, lsBoletaITF
              Print #nFicSal, ""
          Close #nFicSal
        End If
        
    Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
    
    cmdCancelar_Click
End If
Set clsCapMov = Nothing
Set clsCont = Nothing
Set oCapMov = Nothing
Set oITF = Nothing
End Sub

Private Sub Form_Load()
CentraForm Me
Call ImpreSensa
End Sub

Private Sub optMon_Click(Index As Integer)
Select Case Index
    Case 0
       Me.lblTotal.BackColor = &H80000005
    Case 1
        Me.lblTotal.BackColor = &H80FF80
End Select
End Sub

Private Sub txtCodCli_EmiteDatos()
    If txtCodCli <> "" Then
        CargaDatos
    End If
End Sub
Function CargaDatos()
Dim oCaj As COMNCajaGeneral.NCOMCajero
Dim lnMoneda As COMDConstantes.Moneda
Dim i As Long
Dim rs As ADODB.Recordset
Dim oITF As New COMDConstSistema.FCOMITF
Dim lnTotal As Currency
lnMoneda = IIf(optMon(0).value, COMDConstantes.gMonedaNacional, COMDConstantes.gMonedaExtranjera)
Set oCaj = New COMNCajaGeneral.NCOMCajero

Set rs = oCaj.GetCajeroOpeIngxRef(Me.txtCodCli.Text, lnMoneda)
If Not rs.EOF And Not rs.BOF Then
    Set fgOpe.Recordset = rs
    lblnomcli = txtCodCli.psDescripcion
    lnTotal = 0
    For i = 1 To fgOpe.Rows - 1
        lnTotal = lnTotal + CCur(fgOpe.TextMatrix(i, 3))
    Next
    lblMonto = Format(lnTotal, "#0.00")
    If oITF.gbITFAplica And bOpeAfecta Then
        lblITF.Caption = Format(oITF.fgITFCalculaImpuesto(lnTotal), "#,##0.00")
        '*** BRGO 20110908 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.lblITF.Caption))
        If nRedondeoITF > 0 Then
           Me.lblITF.Caption = Format(CCur(Me.lblITF.Caption) - nRedondeoITF, "#,##0.00")
        End If
        '*** END BRGO
    End If
    lblTotal.Caption = Format(CCur(lblITF.Caption) + lnTotal, "#,##0.00")
    
    txtCodCli.Enabled = False
    optMon(0).Enabled = False
    optMon(1).Enabled = False
    txtGlosa.SetFocus
Else
    MsgBox "Cliente Seleccionado no Posee ingresos para refinancion y/o Otras Ope.Credito", vbInformation, "aviso"
    txtCodCli = ""
    Exit Function
End If
Set oCaj = Nothing
Set oITF = Nothing
End Function

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub
