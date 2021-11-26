VERSION 5.00
Begin VB.Form FrmPigRetasacionMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retasacion Manual de Lotes - Piezas"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "FrmPigRetasacionMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   7110
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Piezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   75
      TabIndex        =   5
      Top             =   4140
      Width           =   9570
      Begin SICMACT.FlexEdit feDetalle 
         Height          =   2580
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   4551
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Tipo-SubTipo-Material-Estado-PBruto-PNeto-Tasac-TasAdic-Observacion-item-ObsAdicion"
         EncabezadosAnchos=   "400-1030-1030-1030-1030-700-700-700-800-2200-500-0"
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
         ColumnasAEditar =   "X-1-2-3-4-5-6-X-8-9-X-11"
         ListaControles  =   "0-3-3-3-3-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R-R-R-R-L-C-L"
         FormatosEdit    =   "0-0-0-0-0-2-2-2-2-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contratos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3480
      Left            =   90
      TabIndex        =   3
      Top             =   600
      Width           =   9525
      Begin SICMACT.FlexEdit feContrato 
         Height          =   3105
         Left            =   75
         TabIndex        =   4
         Top             =   285
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5477
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Nro Contrato-Piezas-Estado-Ubicacion-T. Tasacion-Dias Atraso-nEstado"
         EncabezadosAnchos=   "400-1800-650-1600-1600-1600-1000-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-3"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cbSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   7110
      Width           =   1245
   End
   Begin VB.TextBox txtnremate 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1470
      TabIndex        =   1
      Top             =   225
      Width           =   675
   End
   Begin VB.Label lblFecIniRem 
      Caption         =   "Label2"
      Height          =   210
      Left            =   8235
      TabIndex        =   9
      Top             =   1290
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "CONTRATOS PARA RETASACION MANUAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2580
      TabIndex        =   7
      Top             =   180
      Width           =   6330
   End
   Begin VB.Label Label1 
      Caption         =   "Nro.Remate"
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
      Height          =   225
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "FrmPigRetasacionMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTipoJoya As Recordset
Dim rsSTipoJoya As Recordset
Dim rsMaterial As Recordset
Dim rsEstadoJoya As Recordset
Dim lnEstado As Integer
Dim lnTipoTasacion As Integer

Private Sub cbsalir_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
Dim oContFunc As NContFunciones
Dim oGrabarMod As NPigContrato
Dim oImprime As nPigImpre
Dim oPrevio As Previo.clsPrevio
Dim rsJoyas As Recordset

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lsLote As String

    If MsgBox(" Grabar Cambios del contrato? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            cmdGrabar.Enabled = False
    
        Set oContFunc = New NContFunciones
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oContFunc = Nothing
        lsCuenta = Me.feContrato.TextMatrix(feContrato.Row, 1)
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set rsJoyas = feDetalle.GetRsNew
        
        Set oGrabarMod = New NPigContrato
        Call oGrabarMod.nRetasacionManual(lsCuenta, lsFechaHoraGrab, lnEstado, rsJoyas, lsMovNro, gsCodPersUser)
        Set oGrabarMod = Nothing
        Set rsJoyas = Nothing
                    
        feDetalle.Clear
        feDetalle.FormaCabecera
            
    End If

End Sub

Private Sub fecontrato_DblClick()
    Dim lrDatos As dPigContrato
    Dim lrDatosContrato As ADODB.Recordset
    Dim lsCtaCod As String
    Dim i As Integer
    
    lsCtaCod = Me.feContrato.TextMatrix(feContrato.Row, 1)
    lnEstado = CInt(feContrato.TextMatrix(feContrato.Row, 7))
    lnTipoTasacion = feContrato.TextMatrix(feContrato.Row, 5)
    Set lrDatos = New dPigContrato
    Set lrDatosContrato = lrDatos.dObtieneDetalleJoyas(lsCtaCod)
    Set lrDatos = Nothing
    i = 1
    If Not (lrDatosContrato.EOF) Then
      feDetalle.Clear
      feDetalle.Rows = 2
      feDetalle.FormaCabecera
      Do While Not lrDatosContrato.EOF
        feDetalle.AdicionaFila
        i = lrDatosContrato!Item
        feDetalle.TextMatrix(i, 0) = lrDatosContrato!Item
        feDetalle.TextMatrix(i, 1) = lrDatosContrato!Tipo & Space(75) & lrDatosContrato!nTipo
        If lrDatosContrato!nSubTipo <> 0 Or Not IsNull(lrDatosContrato!nSubTipo) Then
            feDetalle.TextMatrix(i, 2) = lrDatosContrato!SubTipo & Space(75) & lrDatosContrato!nSubTipo
        End If
        feDetalle.TextMatrix(i, 3) = lrDatosContrato!Material & Space(75) & lrDatosContrato!nMaterial
        feDetalle.TextMatrix(i, 4) = lrDatosContrato!Estado & Space(75) & lrDatosContrato!nConservacion
        feDetalle.TextMatrix(i, 5) = Format(lrDatosContrato!PBruto, "###,###.00")
        feDetalle.TextMatrix(i, 6) = Format(lrDatosContrato!pNeto, "###,###.00")
        feDetalle.TextMatrix(i, 7) = Format(lrDatosContrato!Tasacion, "###,###.00")
        feDetalle.TextMatrix(i, 8) = Format(IIf(IsNull(lrDatosContrato!TasAdicion), 0, lrDatosContrato!TasAdicion), "###,###.00")
        feDetalle.TextMatrix(i, 9) = IIf(IsNull(lrDatosContrato!Observacion), "", lrDatosContrato!Observacion)
        feDetalle.TextMatrix(i, 10) = lrDatosContrato!Item
        feDetalle.TextMatrix(i, 11) = IIf(IsNull(lrDatosContrato!ObsAdicion), "", lrDatosContrato!ObsAdicion)
        lrDatosContrato.MoveNext
      Loop
    End If
    
   Set lrDatosContrato = Nothing
   Me.cmdGrabar.Enabled = True
End Sub

Private Sub feDetalle_Click()
Dim oConst As DConstante

Set oConst = New DConstante

Select Case feDetalle.Col
Case 1
    Set rsTipoJoya = oConst.RecuperaConstantes(gColocPigTipoJoya, , "C.cConsDescripcion")
    feDetalle.CargaCombo rsTipoJoya
    Set rsTipoJoya = Nothing
Case 2
    Set rsSTipoJoya = oConst.RecuperaConstantes(gColocPigSubTipoJoya, , "C.cConsDescripcion")
    feDetalle.CargaCombo rsSTipoJoya
    Set rsSTipoJoya = Nothing
Case 3
    Set rsMaterial = oConst.RecuperaConstantes(gColocPigMaterial, , "C.cConsDescripcion")
    feDetalle.CargaCombo rsMaterial
    Set rsMaterial = Nothing
Case 4
    Set rsEstadoJoya = oConst.RecuperaConstantes(gColocPigEstConservaJoya, , "C.cConsDescripcion")
    feDetalle.CargaCombo rsEstadoJoya
    Set rsEstadoJoya = Nothing
End Select

Set oConst = Nothing

End Sub

Private Sub FEDetalle_OnCellChange(pnRow As Long, pnCol As Long)
Dim oPigCalculos As NPigCalculos
Dim oPigFunciones As DPigFunciones
Dim lnPrestamo As Currency
Dim lnTasacionT As Currency

    feDetalle.TextMatrix(feDetalle.Row, 11) = feDetalle.TextMatrix(feDetalle.Row, 0)
    
    If feDetalle.Col = 6 Then     'Peso Neto

        If feDetalle.TextMatrix(feDetalle.Row, 6) <> "" Then
            If CCur(feDetalle.TextMatrix(feDetalle.Row, 6)) < 0 Then
                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
                feDetalle.TextMatrix(feDetalle.Row, 6) = 0
            Else
                If CCur(feDetalle.TextMatrix(feDetalle.Row, 6)) > CCur(feDetalle.TextMatrix(feDetalle.Row, 5)) Then
                    MsgBox "Peso Neto no puede ser mayor que Peso Bruto", vbInformation, "Aviso"
                    feDetalle.TextMatrix(feDetalle.Row, 6) = 0
                End If
            End If
        End If
        
    End If

End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtnremate_KeyPress(KeyAscii As Integer)
 Dim lrDatos As dPigContrato
 Dim lrDatosContrato As ADODB.Recordset
 Dim i As Integer
 Dim oParam As DPigFunciones
 Dim lnDiasVencido As Integer
 
 Set oParam = New DPigFunciones
    lnDiasVencido = oParam.GetParamValor(gPigParamDiasAtrasoSelecRemate)
 Set oParam = Nothing

  If KeyAscii = 13 Then 'se presione enter y el boton nuevo en falso
      Set lrDatos = New dPigContrato
      Set lrDatosContrato = lrDatos.dObtieneContratosRetas(gPigTipoTasacRetasac, lnDiasVencido, Me.lblFecIniRem)
      Set lrDatos = Nothing
      i = 1
     If Not (lrDatosContrato.EOF) Then
         Do While Not lrDatosContrato.EOF ' TAMBIEN SE PUEDE DE HACER DE ESTA MANERA
            Me.feContrato.AdicionaFila
            Me.feContrato.TextMatrix(i, 1) = lrDatosContrato!cCtaCod
            Me.feContrato.TextMatrix(i, 2) = lrDatosContrato!npiezas
            Me.feContrato.TextMatrix(i, 3) = lrDatosContrato!Estado
            Me.feContrato.TextMatrix(i, 4) = lrDatosContrato!nUbicaLote
            Me.feContrato.TextMatrix(i, 5) = lrDatosContrato!nTipoTasacion
            Me.feContrato.TextMatrix(i, 6) = lrDatosContrato!nDiasAtraso
            Me.feContrato.TextMatrix(i, 7) = lrDatosContrato!nPrdEstado
            lrDatosContrato.MoveNext
            i = i + 1
        Loop
        Set lrDatosContrato = Nothing
     Else
         MsgBox "No existe este Numero de Remate"
     End If
 End If
End Sub
