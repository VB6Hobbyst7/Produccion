VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPigDetalleRemate 
   Caption         =   "Detalle de Venta"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   3120
      TabIndex        =   26
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1485
      Left            =   5760
      TabIndex        =   16
      Top             =   5955
      Width           =   3135
      Begin SICMACT.EditMoney txtSubTotal 
         Height          =   255
         Left            =   1515
         TabIndex        =   17
         Top             =   165
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtComMart 
         Height          =   255
         Left            =   1515
         TabIndex        =   18
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIGV 
         Height          =   240
         Left            =   1515
         TabIndex        =   19
         Top             =   735
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtTotal 
         Height          =   255
         Left            =   1530
         TabIndex        =   20
         Top             =   1140
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         BackColor       =   12648447
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Com.Martillero"
         Height          =   180
         Left            =   105
         TabIndex        =   24
         Top             =   495
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "IGV"
         Height          =   180
         Left            =   135
         TabIndex        =   23
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "TOTAL"
         Height          =   165
         Left            =   135
         TabIndex        =   22
         Top             =   1185
         Width           =   1140
      End
      Begin VB.Label Label7 
         Caption         =   "Sub.Total"
         Height          =   195
         Left            =   105
         TabIndex        =   21
         Top             =   210
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1485
         X2              =   3060
         Y1              =   1065
         Y2              =   1050
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   4560
      TabIndex        =   15
      Top             =   6960
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   8805
      Begin VB.TextBox txtDocId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   4830
         TabIndex        =   5
         Top             =   1020
         Width           =   1185
      End
      Begin VB.TextBox txtnombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1335
         TabIndex        =   4
         Top             =   1380
         Width           =   6975
      End
      Begin VB.TextBox txtPersCod 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1335
         TabIndex        =   3
         Top             =   990
         Width           =   1980
      End
      Begin VB.TextBox txtDocJur 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   7005
         TabIndex        =   2
         Top             =   990
         Width           =   1290
      End
      Begin VB.TextBox txtDirPers 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   1335
         TabIndex        =   1
         Top             =   1380
         Visible         =   0   'False
         Width           =   6975
      End
      Begin MSMask.MaskEdBox txtNumDoc 
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16711680
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSerie 
         Height          =   300
         Left            =   1350
         TabIndex        =   7
         Top             =   225
         Width           =   510
         _ExtentX        =   900
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   16711680
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblSituacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   4800
         TabIndex        =   27
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label lbalias 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label Label6 
         Caption         =   "Comprador"
         Height          =   225
         Left            =   225
         TabIndex        =   13
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Doc ID"
         Height          =   180
         Left            =   4110
         TabIndex        =   12
         Top             =   1065
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   225
         Left            =   225
         TabIndex        =   11
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   1425
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "Doc Jur"
         Height          =   180
         Left            =   6225
         TabIndex        =   9
         Top             =   1050
         Width           =   750
      End
      Begin VB.Label Label11 
         Caption         =   "Número :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   270
         Width           =   645
      End
   End
   Begin SICMACT.FlexEdit fefacturadet 
      Height          =   4050
      Left            =   90
      TabIndex        =   25
      Top             =   1920
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7144
      Cols0           =   14
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Contrato-Pieza-Tipo-Material-Observacion-PesoNeto-Importe-ComMart-Tasacion-Estado-NumRemate-TipoProceso-IGV"
      EncabezadosAnchos=   "400-1800-500-1100-1100-2000-800-1000-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-L-R-R-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmPigDetalleRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'* frmPigDetalleVenta - Venta de Joyas en Remate
'* EAFA - 15/10/2002
'****************************************************
Dim lnMovNroAnt As Long
Dim lmJoyas(19, 6) As String
Dim lmestJoya As Integer
Dim nFilas As Integer

Public Sub Inicio(ByVal lnTipDoc As Integer, ByVal lnNroDoc As String, ByVal lsFecVta As String, ByVal lsEstDoc As String)
GetTipCambio (lsFecVta)
lblSituacion.Caption = ""
'lblTipoCambio.Caption = Format(gnTipCambioC, "##0.000 ")
If lsEstDoc = "ANULADA" Then
   lmestJoya = gPigSituacionPendFacturar
   cmdAnular.Enabled = False
   lblSituacion.Caption = "ANULADA"
Else
   lmestJoya = gPigSituacionFacturado
End If
MuestraDetalleDocumentos lnTipDoc, lnNroDoc
Me.Show 1

End Sub

Private Sub cmdAnular_Click()
Dim lrPigContrato As NPigContrato
Dim loContFunct As NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String


 If MsgBox(" Esta Ud seguro de Extornar dicha Operación ? ", vbQuestion + vbYesNo + vbDefaultButton2, " Aviso ") = vbNo Then
     Exit Sub
 Else
     Set loContFunct = New NContFunciones
     lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
     Set loContFunct = Nothing
     
     lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
     Set lrPigContrato = New NPigContrato
            Call lrPigContrato.nAnularVentaJoyasRemate(lnMovNroAnt, lsFechaHoraGrab, lmJoyas, lsMovNro, nFilas)
     Set lrPigContrato = Nothing
     lblSituacion.Caption = "ANULADA"
     cmdAnular.Enabled = False
     fefacturadet.Clear
    fefacturadet.Rows = 2
    fefacturadet.FormaCabecera
     Limpiar
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub MuestraDatosCliente(ByVal pscliente As String)
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As dPigContrato
Dim lstTmpCliente As ListItem

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New dPigContrato
      Set lrDatos = lrPigContrato.dObtieneDatosClienteVentaJoyas(pscliente)
Set lrPigContrato = Nothing

If lrDatos.EOF And lrDatos.BOF Then
    Exit Sub
Else
'     lstCliente.ListItems.Clear
'     Set lstTmpCliente = lstCliente.ListItems.Add(, , Format(lrDatos!cPersCod, "##0"))
'            lstTmpCliente.SubItems(1) = PstaNombre(Trim(lrDatos!cPersNombre))
'            If lrDatos!nPersPersoneria = gPersonaNat Then
'                lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrDatos!NroDNI), "", lrDatos!NroDNI))
'            Else
'                lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrDatos!NroRUC), "", lrDatos!NroRUC))
'            End If
'            lstTmpCliente.SubItems(3) = Format(lrDatos!cPersDireccDomicilio, "##0")
'            lstTmpCliente.SubItems(4) = Format(lrDatos!cPersTelefono)
     txtPersCod = Format(lrDatos!cPersCod, "##0")
     If lrDatos!nPersPersoneria = gPersonaNat Then
        txtDocId = Trim(IIf(IsNull(lrDatos!NroDNI), "", lrDatos!NroDNI))
     Else
        txtDocJur = Trim(IIf(IsNull(lrDatos!NroRUC), "", lrDatos!NroRUC))
     End If
     txtnombre = PstaNombre(Trim(lrDatos!cPersNombre))
     Set lrDatos = Nothing
End If
End Sub

Private Sub MuestraDetalleDocumento(ByVal pnTipDoc As Integer, ByVal psNroDoc As String)
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As dPigContrato
Dim lstTmpJoyas As ListItem
Dim fnVarSubTotal As Currency
Dim fnVarIGV As Currency
Dim fnVarTotalS As Currency
Dim fnVarTotalD As Currency
Dim f As Integer

fnVarSubTotal = 0: fnVarIGV = 0: fnVarTotalS = 0: fnVarTotalD = 0: f = 0

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New dPigContrato
      Set lrDatos = lrPigContrato.dObtieneDetDocumentoVentasJoyas(pnTipDoc, Mid(psNroDoc, 1, 4) & Mid(psNroDoc, 6, 8))
Set lrPigContrato = Nothing

If lrDatos.EOF And lrDatos.BOF Then
    Exit Sub
Else
     Select Case lrDatos!nCodTipo             '****** Tipo de Documento
'                 Case gPigTipoBoleta
'                           OptDocumento(0).value = True
'                 Case gPigTipoFactura
'                           OptDocumento(1).value = True
'                 Case gPigTipoBoleta
'                           OptDocumento(2).value = True
                           
     End Select
     Select Case lrDatos!nMotivo                '****** Tipo de Venta (Motivo)
'                 Case gPigTipoVentaATerceros
'                           optTipoVenta(0).value = True
'                 Case gPigTipoVentaPorResponsabilidad
'                           optTipoVenta(1).value = True
'                 Case gPigTipoVentaAlTitular
'                           optTipoVenta(2).value = True
     End Select
     If lrDatos!nFlag = gMovFlagExtornado Then
         lblSituacion.Caption = "ANULADA"
         cmdAnular.Enabled = False
     End If
     mskSerDocumento.Text = Mid(lrDatos!cDocumento, 1, 4)          '**** Serie de Boleta/Factura
     mskNroDocumento.Text = Mid(lrDatos!cDocumento, 6, 8)         '**** Número de Boleta/Factura
     lnMovNroAnt = lrDatos!nNroMov
     pscliente = lrDatos!cPersCod
     lstJoyas.ListItems.Clear
     Do While Not lrDatos.EOF
            Set lstTmpJoyas = lstJoyas.ListItems.Add(, , Format(lrDatos!nItemDoc, "##0"))
                   lstTmpJoyas.SubItems(1) = Trim(lrDatos!cCtaCod)
                   lstTmpJoyas.SubItems(2) = Format(lrDatos!nItemPieza, "##0")
                   lstTmpJoyas.SubItems(3) = Trim(lrDatos!Tipo) & " DE " & Trim(lrDatos!Material) & " DE " & Format(lrDatos!pNeto, "#,##0.00") & " grs. " & Trim(lrDatos!Observacion) & " " & Trim(lrDatos!ObservAdic)
                   lstTmpJoyas.SubItems(4) = Format(lrDatos!subVenta, "###,##0.00")
                   lstTmpJoyas.SubItems(5) = Format(lrDatos!IGV, "##,##0.00")
                   lstTmpJoyas.SubItems(6) = lrDatos!nRema
                   
                   fnVarSubTotal = fnVarSubTotal + Val(lrDatos!subVenta)
                   fnVarIGV = fnVarIGV + Val(lrDatos!IGV)
                   
                   lmJoyas(f, 0) = Trim(lrDatos!cCtaCod)           '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 1) = Trim(lrDatos!nItemPieza)        '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 2) = Trim(lrDatos!nRema)               '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 3) = Trim(lrDatos!subVenta)          '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 4) = Trim(lrDatos!IGV)          '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   f = f + 1
            lrDatos.MoveNext
     Loop
     nFilas = f - 1
     Set lrDatos = Nothing
                   
     fnVarTotalS = Round(fnVarSubTotal + fnVarIGV, 2)
     fnVarTotalD = Round(fnVarTotalS / Val(lblTipoCambio.Caption), 2)
     lblSubTotal.Caption = Format(fnVarSubTotal, "###,##0.00 ")
     lblIGV.Caption = Format(Round(fnVarIGV, 2), "###,##0.00 ")
     lblTotalS.Caption = Format(fnVarTotalS, "###,##0.00 ")
     lblTotalD.Caption = Format(fnVarTotalD, "###,##0.00 ")
     
     MuestraDatosCliente (pscliente)
End If
End Sub

Private Sub MuestraDetalleDocumentos(ByVal pnTipDoc As Integer, ByVal psNroDoc As String)
Dim nDatosPiezas As Recordset
Dim psNombre As String
Dim i As Integer
Dim lnTVenta As Currency, lnTComMart As Currency, lnTIGV As Currency
Dim oParam As DPigFunciones
Dim lsSerPoliza As String
Dim lsNumPoliza As String


Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As dPigContrato
Dim f As Integer

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New dPigContrato
      Set lrDatos = lrPigContrato.dObtieneDetDocumentoVentasJoyasRemate(pnTipDoc, Mid(psNroDoc, 1, 4) & Mid(psNroDoc, 6, 8), lmestJoya)
Set lrPigContrato = Nothing

    Set oParam = New DPigFunciones
    
    lsSerPoliza = Mid(psNroDoc, 1, 4)
    lsNumPoliza = Mid(psNroDoc, 6, 8)
    txtSerie = lsSerPoliza
    txtNumDoc = lsNumPoliza
    
  lnPorcComMart = oParam.GetParamValor(gPigParamPorcComisMartillero)
    
  '  Set oParam = Nothing
    
'    lbalias.Caption = FrmPigClienteRemate.feclienteremate.TextMatrix(FrmPigClienteRemate.feclienteremate.Row, 1)
'    psNombre = lbalias.Caption
    
'    Set nDatosPiezas = FrmPigClienteRemate.fepiezasrem.GetRsNew
'    lnTVenta = 0: lnTComMart = 0: lnTIGV = 0
    lnMovNroAnt = lrDatos!nNroMov
     pscliente = lrDatos!cPersCod
       lbalias = Trim(IIf(IsNull(lrDatos!Comprador), "", lrDatos!Comprador))
   
    fefacturadet.Clear
    fefacturadet.Rows = 2
    fefacturadet.FormaCabecera
    
    lrDatos.MoveFirst
    Do While Not (lrDatos.EOF)
         fefacturadet.AdicionaFila
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 1) = lrDatos!cCtaCod
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 2) = lrDatos!nPieza
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 3) = lrDatos!Tipo
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 4) = lrDatos!Material
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 5) = lrDatos!Observacion
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 6) = lrDatos!pNeto
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 7) = Format(lrDatos!subVenta, "#####,###.00")
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 8) = (CCur(lrDatos!subVenta) * lnPorcComMart / 100)
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 13) = (CCur(lrDatos!subVenta) * lnPorcComMart / 100) * 0.18
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 9) = 0
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 10) = lrDatos!Estado
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 11) = lrDatos!nRema
         fefacturadet.TextMatrix(fefacturadet.Rows - 1, 12) = lrDatos!nTipoProceso
         lnTVenta = lnTVenta + lrDatos!subVenta
         lnTComMart = lnTComMart + (CCur(lrDatos!subVenta) * lnPorcComMart / 100)
         lnTIGV = lnTIGV + ((CCur(lrDatos!subVenta) * lnPorcComMart / 100) * 0.18)
         
                   lmJoyas(f, 0) = Trim(lrDatos!cCtaCod)           '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 1) = Trim(lrDatos!nPieza)        '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 2) = Trim(lrDatos!nRema)               '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 3) = Trim(lrDatos!subVenta)          '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 4) = Trim(lrDatos!IGV)
                   lmJoyas(f, 5) = Trim(lrDatos!nTipoProceso)
                  f = f + 1
         lrDatos.MoveNext
     Loop
     nFilas = f - 1
     Set lrDatos = Nothing
     
     txtSubTotal = Format(lnTVenta, "#####,###.00")
     txtComMart = Format(lnTComMart, "#####,###.00")
     txtIGV = Format(lnTIGV, "######,###.00")
     txtTotal.Text = Format((CCur(txtSubTotal.Text) + CCur(txtComMart.Text) + CCur(txtIGV.Text)), "#####,###.00")
    MuestraDatosCliente (pscliente)
End Sub

Private Sub Limpiar()

    txtDirPers = ""
    txtComMart = ""
    txtDocId = ""
    txtDocJur = ""
    txtIGV = ""
    txtnombre = ""
    txtNumDoc = ""
    txtPersCod = ""
    txtSerie = ""
    txtSubTotal = ""
    txtTotal = ""
    lblSituacion.Caption = ""
    fefacturadet.Clear
    fefacturadet.FormaCabecera

End Sub

