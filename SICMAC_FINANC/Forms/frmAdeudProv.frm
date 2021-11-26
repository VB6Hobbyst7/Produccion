VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAdeudProv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provision de Adeudados"
   ClientHeight    =   4620
   ClientLeft      =   585
   ClientTop       =   2040
   ClientWidth     =   10335
   Icon            =   "frmAdeudProv.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin Sicmact.FlexEdit fgInt 
      Height          =   2625
      Left            =   90
      TabIndex        =   9
      Top             =   600
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   4630
      Cols0           =   18
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmAdeudProv.frx":030A
      EncabezadosAnchos=   "350-500-2000-1800-1200-1000-1000-0-0-0-700-1200-0-0-0-0-1100-0"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-10-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-R-R-R-L-L-L-C-R-L-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-2-2-2-0-0-0-0-2-0-0-0-0-0-0"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.TextBox txtTasaVac 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3390
      TabIndex        =   1
      Top             =   90
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   435
      Left            =   4005
      TabIndex        =   7
      Top             =   4020
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   767
      _Version        =   393217
      TextRTF         =   $"frmAdeudProv.frx":03AE
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8520
      TabIndex        =   2
      Top             =   105
      Width           =   1680
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   105
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8760
      TabIndex        =   5
      Top             =   4065
      Width           =   1440
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   7335
      TabIndex        =   4
      Top             =   4065
      Width           =   1440
   End
   Begin VB.TextBox txtGlosa 
      Height          =   630
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3285
      Width           =   10155
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2970
      Top             =   3930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   42
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdeudProv.frx":0430
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdeudProv.frx":0C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tasa VAC :"
      Height          =   195
      Left            =   2490
      TabIndex        =   8
      Top             =   165
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha : "
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
      Left            =   240
      TabIndex        =   6
      Top             =   165
      Width           =   720
   End
End
Attribute VB_Name = "frmAdeudProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbTrans  As Boolean
Dim lnTasaVac As Double
Dim lMN As Boolean
Dim oPrg As clsProgressBar

Private Function ValidaDatos() As Boolean
Dim i As Integer
Dim lbSel As Boolean

ValidaDatos = False
    If fgInt.TextMatrix(1, 0) = "" Then
        MsgBox "Lista se encuentra vacia para realizar las operacion", vbInformation, "Aviso"
        fgInt.SetFocus
        Exit Function
    End If
    lbSel = False
    For i = 1 To Me.fgInt.Rows - 1
        If fgInt.TextMatrix(i, 1) = "." Then
            lbSel = True
            Exit For
        End If
    Next
    If lbSel = False Then
        MsgBox "Seleccione algun pagaré para generar la Operación", vbInformation, "Aviso"
        fgInt.SetFocus
        Exit Function
    End If
    If Len(Trim(txtGlosa)) = 0 Then
        MsgBox "Por favor Ingrese la Glosa para la Operación", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Function
    End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
    Dim lsMovNro   As String
    Dim lnMovNro   As Long
    Dim lnImporte  As Currency
    Dim lsPersCod  As String
    Dim lsIFTpo    As String
    Dim lsCtaIFCod As String
    Dim i As Integer, J As Integer
    Dim lsMsgErr   As String
    Dim aMovs() As String
    Dim nMov    As Integer
    
    On Error GoTo ErrorAceptar
    If Not ValidaDatos Then
        Exit Sub
    End If
    Dim Item As Integer
    Dim lsCtaDebe As String, lsCtaHaber As String
    
    If MsgBox(" ¿ Desea Grabar la Operación ? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        Exit Sub
    End If
    i = 0
    
    Dim oMov As New DMov
    Dim oOpe As New DOperacion
    Dim oContImp As New NContImprimir
  
    oMov.BeginTrans
    lbTrans = True
    nMov = 0
    Do While True
        If fgInt.TextMatrix(i, 1) = "." Then
            Item = 0
            
            lsPersCod = Mid(fgInt.TextMatrix(i, 14), 4, 13)
            lsIFTpo = Mid(fgInt.TextMatrix(i, 14), 1, 2)
            lsCtaIFCod = Mid(fgInt.TextMatrix(i, 14), 18, 10)
            
            lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , fgInt.TextMatrix(i, 14), ObjEntidadesFinancieras, True)
            If lsCtaDebe = "" Then
                Err.Raise 50001, "Provision Adeudados", "No se definió correctamente Cuenta del Debe de la Operación"
            End If
            
            lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , fgInt.TextMatrix(i, 14), ObjEntidadesFinancieras, True)
            If lsCtaHaber = "" Then
                Err.Raise 50001, "Provision Adeudados", "No se definió correctamente Cuenta del Haber de la Operación"
            End If
            lsMovNro = oMov.GeneraMovNro(txtFecha, Right(gsCodAge, 2), gsCodUser)
            lnImporte = CCur(fgInt.TextMatrix(i, 11))
            
            oMov.InsertaMov lsMovNro, gsOpeCod, txtGlosa, gMovEstContabMovContable, gMovFlagVigente
            lnMovNro = oMov.GetnMovNro(lsMovNro)
            oMov.InsertaMovCont lnMovNro, lnImporte, 0, ""

            Item = Item + 1
            oMov.InsertaMovCta lnMovNro, Item, lsCtaDebe, lnImporte
            oMov.InsertaMovObj lnMovNro, Item, 1, ObjEntidadesFinancieras
            oMov.InsertaMovObjIF lnMovNro, Item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
            
            Item = Item + 1
            oMov.InsertaMovCta lnMovNro, Item, lsCtaHaber, lnImporte * -1
            oMov.InsertaMovObj lnMovNro, Item, 1, ObjEntidadesFinancieras
            oMov.InsertaMovObjIF lnMovNro, Item, 1, lsPersCod, lsIFTpo, lsCtaIFCod
           
            If Val(txtTasaVac) > 0 And Mid(gsOpeCod, 3, 1) = "1" Then
                oMov.InsertaMovTpoCambio lnMovNro, CCur(txtTasaVac)
            End If
            
            oMov.ActualizaCtaIF lsPersCod, lsIFTpo, lsCtaIFCod, , , , Format(txtFecha, gsFormatoFecha), , , , lnImporte, lsMovNro, True
            oMov.ActualizaAdeudadosProvision lsPersCod, lsIFTpo, lsCtaIFCod, Format(txtFecha, gsFormatoFecha), lnImporte, fgInt.TextMatrix(i, 7), "", lsMovNro, ""
           
           If Mid(gsOpeCod, 3, 1) = "2" Then
              oMov.GeneraMovME lnMovNro, lsMovNro
           End If
            oMov.ActualizaSaldoMovimiento lsMovNro, "+"
            nMov = nMov + 1
            ReDim Preserve aMovs(nMov)
            aMovs(nMov) = lsMovNro
            fgInt.EliminaFila i
       Else
            i = i + 1
       End If
        If i >= Me.fgInt.Rows Or fgInt.TextMatrix(1, 0) = "" Then
             Exit Do
        End If
    Loop
    oMov.CommitTrans
    lbTrans = False
    rtf.Text = ""
    For i = 1 To nMov
         rtf.Text = rtf.Text & oContImp.ImprimeAsientoContable(aMovs(i), gnLinPage, gnColPage, "PROVISION DE PAGARES DE ADEUDADOS", , "19", , gsNomCmac) & oImpresora.gPrnSaltoPagina
    Next
    EnviaPrevio rtf.Text, "Asiento de Provision de Adeudados - Caja General", 66
    If MsgBox("Desea realizar otra operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Unload Me
        Exit Sub
    End If
    
    Set oMov = Nothing
    Set oOpe = Nothing
    Set oContImp = Nothing
    
    txtGlosa = ""
    fgInt.SetFocus
    Exit Sub
ErrorAceptar:
    lsMsgErr = TextErr(Err.Description)
    If lbTrans Then
        oMov.RollbackTrans
    End If
    MsgBox lsMsgErr, vbInformation, "Aviso"
  
End Sub

Private Sub cmdProcesar_Click()
    On Error GoTo ErrorAceptar
    If ValFecha(txtFecha) = False Then Exit Sub
    Me.MousePointer = 11
    Me.CmdProcesar.Enabled = False
    CargaDatos txtFecha
    Me.CmdProcesar.Enabled = True
    Me.MousePointer = 0
    Exit Sub
ErrorAceptar:
    Me.MousePointer = 0
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgInt_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If fgInt.TextMatrix(pnRow, pnCol) = "." Then
    If Val(fgInt.TextMatrix(pnRow, 11)) <= 0 Then
        MsgBox "Monto no válido", vbInformation, "Aviso"
        Exit Sub
    End If
End If
End Sub

Private Sub fgInt_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim lnPeriodo As Long
    Dim lnTasaInt As Currency
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency
    Dim lnDias    As Long
    Dim oAdeud As New DCaja_Adeudados
    If Val(fgInt.TextMatrix(pnRow, 10)) > 0 Then
       
        If Val(Left(fgInt.TextMatrix(pnRow, 14), 2)) = gTpoIFFuenteFinanciamiento Then
            lnMontoTotal = CCur(fgInt.TextMatrix(pnRow, 15)) - CCur(fgInt.TextMatrix(pnRow, 17))
        Else
            lnMontoTotal = CCur(fgInt.TextMatrix(pnRow, 15)) + CCur(fgInt.TextMatrix(pnRow, 5))
        End If
        lnTasaInt = CCur(fgInt.TextMatrix(pnRow, 9))
        lnPeriodo = Val(fgInt.TextMatrix(pnRow, 8))
        lnDias = fgInt.TextMatrix(pnRow, 10)
        lnInteres = oAdeud.CalculaInteres(lnDias, lnPeriodo, lnTasaInt, lnMontoTotal)
        fgInt.TextMatrix(pnRow, 12) = Format(lnInteres, "#0.00")
        If fgInt.TextMatrix(pnRow, 13) = "2" And Mid(fgInt.TextMatrix(pnRow, 14), 20, 1) = "1" Then
            fgInt.TextMatrix(pnRow, 11) = Format(lnInteres * lnTasaVac, "#,#0.00")
        Else
            fgInt.TextMatrix(pnRow, 11) = Format(lnInteres, "#,#0.00")
        End If
    Else
        Cancel = False
    End If
    Set oAdeud = Nothing
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Caption = gsOpeDesc
    Me.txtFecha = gdFecSis
    
    Dim oAdeud As New DCaja_Adeudados
    lnTasaVac = oAdeud.CargaIndiceVAC(gdFecSis)
    txtTasaVac = lnTasaVac
End Sub

Private Sub CargaDatos(ldFecha As Date)
    Dim rs As ADODB.Recordset
    Dim N As Integer
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency
    Dim lnTotal As Integer, i As Integer
    
    Set oPrg = New clsProgressBar
    
    lnTasaVac = txtTasaVac
    If lnTasaVac = 0 Then
        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada" & Chr(13) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            txtTasaVac.SetFocus
            Exit Sub
        End If
    End If
    Dim oAdeud As New DCaja_Adeudados
    Dim oIF As New NCajaAdeudados
    
    Set rs = oAdeud.GetAdeudadosProvision(gsOpeCod, txtFecha, Mid(gsOpeCod, 3, 1))
    lnTotal = rs.RecordCount
    i = 0
    fgInt.Clear
    fgInt.Rows = 2
    fgInt.FormaCabecera
    oPrg.ShowForm Me
    oPrg.Max = rs.RecordCount
    oPrg.CaptionSyle = eCap_CaptionPercent
    Do While Not rs.EOF
        oPrg.Progress rs.Bookmark, "Provisión de Adeudados", "Procesando...", , vbBlue
        i = i + 1
        fgInt.AdicionaFila , , True
        N = fgInt.Row
        fgInt.TextMatrix(N, 2) = Trim(rs!cPersNombre)
        fgInt.TextMatrix(N, 3) = Trim(rs!cCtaIFDesc)
        
        fgInt.TextMatrix(N, 5) = Trim(rs!nInteresPagado) ' Interes acumulado pagado por cuota
        fgInt.TextMatrix(N, 6) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
        fgInt.TextMatrix(N, 7) = Trim(rs!nNroCuota)    ' numero de cuota pendiente
        fgInt.TextMatrix(N, 8) = Trim(rs!nPeriodo)
        fgInt.TextMatrix(N, 9) = Trim(rs!nInteres)
        fgInt.TextMatrix(N, 10) = Trim(rs!nDiasUltPAgo + IIf(gsCodCMAC = "102", 1, 0))
        If Val(Left(rs!cIFTpo, 2)) = gTpoIFFuenteFinanciamiento Then
            lnMontoTotal = rs!nSaldoCap - rs!nSaldoConcesion
        Else
            lnMontoTotal = rs!nSaldoCap + rs!nInteresPagado
        End If
        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
            fgInt.TextMatrix(N, 4) = Format(rs!nSaldoCap * lnTasaVac, "#,#0.00")  'Saldo * la tasa vac
        Else
            If Not gsCodCMAC = "102" Then
                fgInt.TextMatrix(N, 4) = Format(rs!nSaldoCap, "#,#0.00")
            Else
                fgInt.TextMatrix(N, 4) = Format(lnMontoTotal, "#,#0.00")
            End If
        End If
        lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo + IIf(gsCodCMAC = "102", 1, 0), rs!nPeriodo, rs!nInteres, lnMontoTotal)
        fgInt.TextMatrix(N, 12) = Format(lnInteres, "#0.00")
        If lnInteres > 0 Then
           fgInt.TextMatrix(N, 1) = "1"
        End If
        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
            If lnTasaVac > 0 Then
                lnInteres = Format(lnInteres * lnTasaVac, "#0.00")
            Else
                lnInteres = Format(lnInteres, "#0.00")
            End If
        End If
        fgInt.TextMatrix(N, 11) = Format(lnInteres, "#,#0.00")  'interes al cambio en soles si el pago es en dolares
        fgInt.TextMatrix(N, 13) = Trim(rs!cMonedaPago)
        fgInt.TextMatrix(N, 14) = rs!cIFTpo & "." & Trim(rs!cPersCod & "." & rs!cCtaIFCod)
        fgInt.TextMatrix(N, 15) = rs!nSaldoCap
        fgInt.TextMatrix(N, 16) = rs!dVencimiento
        fgInt.TextMatrix(N, 17) = rs!nSaldoConcesion
        rs.MoveNext
    Loop
    oPrg.CloseForm Me
    Set oPrg = Nothing
    RSClose rs
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    Dim oAdeud As New DCaja_Adeudados
    If KeyAscii = 13 Then
        lnTasaVac = oAdeud.CargaIndiceVAC(txtFecha)
        txtTasaVac = Format(lnTasaVac, "#,###.00####")
        txtTasaVac.SetFocus
    End If
    Set oAdeud = Nothing
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtTasaVac_GotFocus()
    fEnfoque txtTasaVac
End Sub

Private Sub txtTasaVac_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTasaVac, KeyAscii, 15, 8)
    If KeyAscii = 13 Then
        Me.CmdProcesar.SetFocus
    End If
End Sub

Private Sub txtTasaVac_LostFocus()
    If txtTasaVac = "" Then txtTasaVac = 0
End Sub
