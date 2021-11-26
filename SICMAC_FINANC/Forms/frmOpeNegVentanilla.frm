VERSION 5.00
Begin VB.Form frmOpeNegVentanilla 
   Caption         =   "Operaciones en Ventanilla: Pendientes"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   Icon            =   "frmOpeNegVentanilla.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAcumulaAge 
      Caption         =   "Acumular datos de Agencias"
      Height          =   315
      Left            =   1830
      TabIndex        =   18
      Top             =   4620
      Width           =   2385
   End
   Begin VB.CheckBox chkAjuste 
      Caption         =   "Ajustar diferencia"
      Height          =   315
      Left            =   90
      TabIndex        =   17
      Top             =   4620
      Width           =   1845
   End
   Begin Sicmact.ProgressBarra oPrg 
      Height          =   405
      Left            =   60
      TabIndex        =   16
      Top             =   4500
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
   End
   Begin VB.Frame fraAge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   780
      Width           =   8910
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   345
         Left            =   1065
         TabIndex        =   13
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   -2147483641
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   750
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2445
         TabIndex        =   14
         Top             =   210
         Width           =   6315
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6060
      TabIndex        =   9
      Top             =   4590
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7470
      TabIndex        =   8
      Top             =   4590
      Width           =   1380
   End
   Begin VB.Frame fraDatosPrinc 
      Caption         =   "Datos Generales"
      Height          =   675
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   8925
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2430
         TabIndex        =   7
         Top             =   255
         Width           =   3780
      End
      Begin VB.Label lblCaptionDocNro 
         AutoSize        =   -1  'True
         Caption         =   "A Rendir N° :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   285
         Width           =   930
      End
      Begin VB.Label lblNroDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1095
         TabIndex        =   5
         Top             =   255
         Width           =   1320
      End
      Begin VB.Label lblArendirMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   7005
         TabIndex        =   4
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   6315
         TabIndex        =   3
         Top             =   315
         Width           =   615
      End
   End
   Begin VB.Frame fraVentanilla 
      Caption         =   "Operaciones Pendientes "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   3000
      Left            =   60
      TabIndex        =   0
      Top             =   1470
      Width           =   8910
      Begin Sicmact.FlexEdit fgPago 
         Height          =   2175
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   3836
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Ord-#1-Ok-Fecha-Persona-Importe-Regulariza-cPersCod-cAgeCod-cCtaPendiente-Descripcion-Nro.Doc.-cCodOpe"
         EncabezadosAnchos=   "0-0-350-1200-3200-1200-1200-0-0-0-3000-1500-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-6-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-L-R-R-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-2-0-0-0-0-0-0"
         TextArray0      =   "Ord"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   5700
         TabIndex        =   11
         Top             =   2610
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6810
         TabIndex        =   10
         Top             =   2565
         Width           =   1965
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   5340
         Top             =   2550
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmOpeNegVentanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vbOk As Boolean
Dim vsAgeCod As String
Dim rsAux As ADODB.Recordset
Dim lnMoneda As Moneda
Dim lnMonto As Currency
Dim lnArendirFase As ARendirFases
Dim lnDiferencia As Currency
Dim lsMotivo     As String
Dim lsCaptionDocNro As String
Dim lnParamDif   As Currency
Dim lsAgeCod     As String
Dim lsCtaContCod As String
Dim lsClaseCta   As String
Dim lbOpeNegocio As Boolean

Public Sub Inicio(ByVal psNroDoc As String, PnMoneda As Moneda, ByVal pnMontoARendir As Currency, _
       Optional ByVal psPersCod As String, Optional ByVal psPersNombre As String, Optional psDocCaption As String, _
       Optional psAgeCod As String = "", Optional psCtaContCod As String = "", Optional psClaseCta As String = "", Optional pbOpeNegocio As Boolean = True)

lnMoneda = PnMoneda
lnMonto = pnMontoARendir
lsAgeCod = psAgeCod
lsCtaContCod = psCtaContCod
lsClaseCta = IIf(psClaseCta = "A", "H", psClaseCta)
lbOpeNegocio = pbOpeNegocio
If psPersCod = "" Then
    fraDatosPrinc.Visible = False
    fraAge.Top = fraAge.Top - fraDatosPrinc.Height
    fraVentanilla.Top = fraVentanilla.Top - fraDatosPrinc.Height
    cmdAceptar.Top = cmdAceptar.Top - fraDatosPrinc.Height
    cmdCancelar.Top = cmdCancelar.Top - fraDatosPrinc.Height
    Me.chkAjuste.Top = Me.chkAjuste.Top - fraDatosPrinc.Height
    Me.Height = Me.Height - fraDatosPrinc.Height
Else
    lblPersNombre = psPersNombre
    lblPersNombre.Tag = psPersCod
    lblNroDoc = psNroDoc
    lsCaptionDocNro = psDocCaption
End If
lblArendirMonto = Format(Abs(pnMontoARendir), gsFormatoNumeroView)

If Val(pnMontoARendir) < 0 Then
   lblArendirMonto.ForeColor = &HFF&
Else
   lblArendirMonto.ForeColor = &HC00000
End If
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
lnDiferencia = 0
    If CCur(lblTotal) <> CCur(lblArendirMonto) Then
        If Abs(CCur(lblTotal) - CCur(lblArendirMonto)) < lnParamDif Then
            If MsgBox("Monto Seleccionado no cubre el Monto a Regularizar " & Chr(13) & "Existe una diferencia que puede ser Ajustada." & Chr(13) & "Desea continuar???", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                fgPago.SetFocus
                Exit Sub
            End If
        Else
            If chkAjuste.value = vbChecked Then
               MsgBox "Diferencia de " & Format(Abs(CCur(lblTotal) - CCur(lblArendirMonto)), gsFormatoNumeroView) & " se ajustará!!! ", vbInformation, "¡Aviso!"
            Else
               MsgBox "Monto Seleccionado no cubre el Monto a Regularizar", vbInformation, "Aviso"
               fgPago.SetFocus
               Exit Sub
            End If
        End If
    End If
lnDiferencia = CCur(lblArendirMonto) - CCur(Abs(lblTotal))
Set rsAux = fgPago.GetRsNew
vbOk = True
vsAgeCod = txtAgeCod
Unload Me
DoEvents
End Sub
Private Sub cmdCancelar_Click()
vbOk = False
vsAgeCod = ""
Unload Me
End Sub

Public Property Get lsAgeCodRef() As String
    lsAgeCodRef = vsAgeCod
End Property
Public Property Let lsAgeCodRef(ByVal vNewValue As String)
    vsAgeCod = vNewValue
End Property

Public Property Get lbOk() As Variant
    lbOk = vbOk
End Property
Public Property Let lbOk(ByVal vNewValue As Variant)
    vbOk = vNewValue
End Property

Public Property Get rsPago() As ADODB.Recordset
    Set rsPago = rsAux
End Property
Public Property Let rsPago(ByVal vNewValue As ADODB.Recordset)
    Set rsAux = vNewValue
End Property

Private Sub fgPago_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim lnTotal As Currency
Dim N As Integer
lnTotal = 0
For N = 1 To fgPago.Rows - 1
    If fgPago.TextMatrix(N, 2) = "." Then
        lnTotal = lnTotal + nVal(fgPago.TextMatrix(N, 6))
        lsMotivo = fgPago.TextMatrix(N, 10)
    End If
Next
lblTotal = Format(lnTotal, gsFormatoNumeroView)
End Sub

Private Sub Form_Activate()
CargaLista Right(lsAgeCod, 2)

End Sub

Private Sub Form_Load()
Dim oGen As DGeneral
Dim oAreas As DActualizaDatosArea
Set oGen = New DGeneral
Set oAreas = New DActualizaDatosArea
lnParamDif = oGen.GetParametro(4000, 1001)
Set oGen = Nothing
CentraForm Me
Me.Caption = gsOpeDesc
If lbOpeNegocio Then
   If lsAgeCod = "" Then
       txtAgeCod.rs = oAreas.GetAgencias
       txtAgeCod = Right(gsCodAge, 2)
       lblAgeDesc.Caption = gsNomAge
       lsAgeCod = txtAgeCod
   Else
       txtAgeCod = lsAgeCod
       lblAgeDesc.Caption = oAreas.GetNombreAgencia(txtAgeCod)
       txtAgeCod.Enabled = False
   End If
Else
   fraAge.Visible = False
   chkAcumulaAge.Visible = False
End If

If lsCaptionDocNro <> "" Then lblCaptionDocNro = lsCaptionDocNro
CargaLista Right(lsAgeCod, 2)
lblTotal = Format(fgPago.SumaRow(2), gsFormatoNumeroView)

End Sub

Public Property Get vnDiferencia() As Currency
vnDiferencia = lnDiferencia
End Property

Public Property Let vnDiferencia(ByVal vNewValue As Currency)
lnDiferencia = vNewValue
End Property

Public Property Get vsMotivo() As String
vsMotivo = lsMotivo
End Property

Public Property Let vsMotivo(ByVal vNewValue As String)
lsMotivo = vNewValue
End Property

Private Sub CargaLista(psCodAge As String)
Dim lvItem As ListItem
Dim N  As Integer
Dim rs As ADODB.Recordset
Dim oNeg As New NNegOpePendientes
Dim oAna As New NAnalisisCtas
If lsCtaContCod <> "" Then
   If lbOpeNegocio Then
    'Por ahora no Interesa el Tipo de Cuenta
      If gbBitCentral Then
          Set rs = oNeg.CargaOpeVentanillaPendCtaContCentral(lsCtaContCod, gsCodCMAC & psCodAge, Mid(gsOpeCod, 3, 1), Mid(lblPersNombre.Tag, 4, 10))
      Else
          Set rs = oNeg.CargaOpeVentanillaPendCtaCont(lsCtaContCod, lsClaseCta, gsCodCMAC & psCodAge, lnMonto, Mid(gsOpeCod, 3, 1), Mid(lblPersNombre.Tag, 4, 10))
      End If
   Else
      Set rs = oAna.GetOpePendientesMov(gbBitCentral, gdFecSis, Mid(gsOpeCod, 3, 1), lsCtaContCod, lsClaseCta)
   End If
Else
    If gbBitCentral Then
        'Set rs = oNeg.CargaIngresoVentanillaPendiente(gOtrOpeIngresosoCajaGeneral, psCodAge, Mid(gsOpeCod, 3, 1), lblPersNombre.Tag)
        Set rs = oNeg.CargaIngresoVentanillaPendiente("300421", psCodAge, Mid(gsOpeCod, 3, 1), lblPersNombre.Tag)
    Else
        Set rs = oNeg.CargaIngresoVentanillaPendiente(gCapIngresoRegulaCG, gsCodCMAC & psCodAge, Mid(gsOpeCod, 3, 1), Mid(lblPersNombre.Tag, 4), True)
    End If
End If
If Not chkAcumulaAge.value = vbChecked Then
   fgPago.Clear
   fgPago.Rows = 2
   fgPago.FormaCabecera
End If
If Not rs.EOF Then
    oPrg.Max = rs.RecordCount
End If
Do While Not rs.EOF
    oPrg.Progress rs.Bookmark
   fgPago.AdicionaFila
   N = fgPago.Row
   fgPago.TextMatrix(N, 1) = N
   If lbOpeNegocio Then
      fgPago.TextMatrix(N, 3) = Format(rs!DFECTRAN, gsFormatoFechaView)
      fgPago.TextMatrix(N, 4) = rs!cNomPers
      fgPago.TextMatrix(N, 5) = Format(rs!nMonTran, gsFormatoNumeroView)
      fgPago.TextMatrix(N, 6) = Format(rs!nMonTran, gsFormatoNumeroView)
      
      If gbBitCentral Then
          fgPago.TextMatrix(N, 7) = rs!cPersCodIF
          fgPago.TextMatrix(N, 8) = psCodAge
          fgPago.TextMatrix(N, 11) = rs!nMovNro
          fgPago.TextMatrix(N, 12) = rs!cOpeCod
      Else
          fgPago.TextMatrix(N, 7) = gsCodCMAC & rs!cCodPers
          fgPago.TextMatrix(N, 8) = gsCodCMAC & psCodAge
          fgPago.TextMatrix(N, 9) = rs!cCtaContCod
          fgPago.TextMatrix(N, 11) = rs!nnumTran
          fgPago.TextMatrix(N, 12) = rs!cCodOpe
      End If
      fgPago.TextMatrix(N, 10) = rs!cGlosa
   Else
      fgPago.TextMatrix(0, 4) = "Doc.Girado"
      fgPago.TextMatrix(N, 3) = Format(GetFechaMov(rs!cMovNro, True), gsFormatoFechaView)
      fgPago.TextMatrix(N, 4) = rs!cDocAbrev & "-" & rs!cDocNro
      fgPago.TextMatrix(N, 5) = Format(rs!nMovImporte, gsFormatoNumeroView)
      fgPago.TextMatrix(N, 6) = Format(rs!nMovImporte, gsFormatoNumeroView)
      
      fgPago.TextMatrix(N, 7) = rs!cPersCod
      fgPago.TextMatrix(N, 8) = "" 'Agencia
      fgPago.TextMatrix(N, 9) = lsCtaContCod
      fgPago.TextMatrix(N, 11) = rs!nMovNro
      fgPago.TextMatrix(N, 12) = "" 'rs!cOpeCod
      fgPago.TextMatrix(N, 10) = rs!cMovDesc
   End If
   rs.MoveNext
Loop
RSClose rs
Set oAna = Nothing
Set oNeg = Nothing

End Sub

Private Sub txtAgeCod_EmiteDatos()
    lblAgeDesc.Caption = txtAgeCod.psDescripcion
    If lblAgeDesc <> "" Then
        ProgressShow oPrg, Me
        CargaLista txtAgeCod
        ProgressClose oPrg, Me
        fgPago.SetFocus
    End If
End Sub
