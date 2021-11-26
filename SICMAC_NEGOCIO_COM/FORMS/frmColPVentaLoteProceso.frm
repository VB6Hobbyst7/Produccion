VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPVentaLoteProceso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Venta en Lote de Prendas Adjudicadas"
   ClientHeight    =   6645
   ClientLeft      =   2115
   ClientTop       =   2265
   ClientWidth     =   8790
   Icon            =   "frmColPVentaLoteProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar Venta"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Frame fraCuentas 
      Caption         =   " Créditos Adjudicados "
      Height          =   3735
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   8535
      Begin VB.OptionButton OptGasto 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   3360
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton OptGasto 
         Caption         =   "Ninguno"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   29
         Top             =   3360
         Width           =   990
      End
      Begin SICMACT.FlexEdit FECredito 
         Height          =   2985
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   8160
         _extentx        =   13547
         _extenty        =   5265
         cols0           =   12
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "--Credito-Cliente-Cod. Proc. Adj.-Saldo-Tasacion-K14-K16-K18-K21-ValReg"
         encabezadosanchos=   "300-300-1800-4000-1500-0-0-0-0-0-0-0"
         font            =   "frmColPVentaLoteProceso.frx":030A
         fontfixed       =   "frmColPVentaLoteProceso.frx":0336
         columnasaeditar =   "X-1-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-4-0-0-0-0-0-0-0-0-0-0"
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         encabezadosalineacion=   "C-C-L-L-R-C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
         selectionmode   =   1
         lbeditarflex    =   -1  'True
         lbflexduplicados=   0   'False
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483635
      End
      Begin VB.Label lblTotalS 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Total de Cuentas Seleccionadas :"
         Height          =   195
         Left            =   4680
         TabIndex        =   28
         Top             =   3360
         Width           =   2745
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Total :"
         Height          =   195
         Left            =   2760
         TabIndex        =   27
         Top             =   3360
         Width           =   795
      End
   End
   Begin VB.Frame fraAdjudicacion 
      Caption         =   "Adjudicaciones Disponibles"
      Height          =   1670
      Index           =   3
      Left            =   6360
      TabIndex        =   25
      Top             =   0
      Width           =   2325
      Begin VB.ListBox lstAdjudicacion 
         Height          =   1185
         ItemData        =   "frmColPVentaLoteProceso.frx":0364
         Left            =   105
         List            =   "frmColPVentaLoteProceso.frx":0366
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   240
         Width           =   2070
      End
   End
   Begin VB.CommandButton cmdPlaVenta 
      Caption         =   "Planilla de Venta en Lote"
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   2730
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos del Proceso"
      Height          =   1665
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6195
      Begin VB.Frame Frame4 
         Caption         =   "Precios del Oro "
         Height          =   600
         Left            =   150
         TabIndex        =   19
         Top             =   960
         Width           =   5970
         Begin VB.TextBox txtPreOro14 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   720
            MaxLength       =   6
            TabIndex        =   4
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro16 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2100
            MaxLength       =   6
            TabIndex        =   5
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro18 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3600
            MaxLength       =   6
            TabIndex        =   6
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro21 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5160
            MaxLength       =   6
            TabIndex        =   7
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label5 
            Caption         =   "21 Kl. :"
            Height          =   225
            Index           =   4
            Left            =   4560
            TabIndex        =   23
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "18 Kl. :"
            Height          =   225
            Index           =   3
            Left            =   3000
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "16 Kl. :"
            Height          =   225
            Index           =   2
            Left            =   1545
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "14 Kl. :"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtNumVenta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   210
         Width           =   1185
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         TabIndex        =   3
         Top             =   570
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtFecVenta 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   570
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHorVenta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   570
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   17
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   615
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Número :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   3960
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmColPVentaLoteProceso.frx":0368
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmColPVentaLoteProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'** Venta en Lote de Prendas Adjudicados
'Archivo:  frmColPVentaLoteProceso.frm
'DAOR   :  31/01/2008.
'Resumen:  Muestra la venta Venta actual y permite:
'   - Iniciar Venta
'   - Registrar las Ventas en Lote
'   - Finalizar Venta
'   - Generar la planilla de ventas en lote.
'********************************************************************************

Option Explicit

Dim fnTasaIGV As Double
Dim fsProcesoCadaAgencia As String

Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim pAgeRemSub As String * 2
Dim pBandEOASub As Boolean
Dim pVerCodAnt As Boolean
Dim RegVenta As New ADODB.Recordset
Dim RegCredPrend As New ADODB.Recordset


Dim RegJoyas As New ADODB.Recordset
Dim RegProcesar As New ADODB.Recordset
Dim MuestraImpresion As Boolean
Dim vRTFImp As String
Dim vCont As Double

Dim fsListaAgencias As String
Dim fnNumRegTotal As Integer
Dim fnNumRegTotalSel As Integer




Private Sub cmdAplicar_Click()
Dim loColPRecGar As COMNColoCPig.NCOMColPRecGar
Dim lrs As ADODB.Recordset
Dim i As Integer
    
    
    
    If lstAdjudicacion.ListCount <= 0 Then
        MsgBox "No existen adjudicaciones para realizar el proceso de venta en lote", vbInformation, "Aviso"
        Exit Sub
    End If
    
    fsListaAgencias = getAgencias
    
    If fsListaAgencias = "" Then
        MsgBox "No ha seleccionado ninguna adjudicación", vbInformation, "Aviso"
        Exit Sub
    End If
    
    LimpiaFlex FECredito
        
    Set loColPRecGar = New COMNColoCPig.NCOMColPRecGar
    Set lrs = loColPRecGar.ObtieneCreditosParaVentaEnLote(fsListaAgencias)
    Set loColPRecGar = Nothing
    
    
    Set FECredito.Recordset = lrs
    For i = 1 To FECredito.Rows - 1
        FECredito.TextMatrix(i, 0) = i
        FECredito.TextMatrix(i, 1) = 1
    Next i

    If FECredito.Rows <= 1 Then
        FECredito.Enabled = False
    Else
        FECredito.Enabled = True
    End If
    fnNumRegTotal = lrs.RecordCount
    fnNumRegTotalSel = lrs.RecordCount
    lblTotal.Caption = "Nro. Total : " & fnNumRegTotal
    lblTotalS.Caption = "Nro. Total de Cuentas Seleccionadas : " & fnNumRegTotalSel
    Set lrs = Nothing
End Sub

Private Function getAgencias() As String
Dim i As Integer
    For i = 0 To lstAdjudicacion.ListCount - 1
        If lstAdjudicacion.Selected(i) Then
            'getAgencias = getAgencias & "'" & Left(lstAdjudicacion.List(i), 2) & "',"
            getAgencias = getAgencias & Right("0000" & lstAdjudicacion.ItemData(i), 4) & ","
        End If
    Next i
    If Len(getAgencias) > 1 Then
        getAgencias = Left(getAgencias, Len(getAgencias) - 1)
    End If
End Function


Private Sub cmdPlaVenta_Click()
Dim loColPRecGar As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsprevio

On Error GoTo ControlError
    lsCadImprimir = ""
    Set loColPRecGar = New COMNColoCPig.NCOMColPRecGar
    lsCadImprimir = loColPRecGar.nImprimeListadoVentaEnLote(Format(txtFecVenta.Text, "mm/dd/yyyy"), _
        Right(fsProcesoCadaAgencia, 2), txtNumVenta.Text, gdFecSis, _
        CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), _
        gsNomCmac, gsNomAge, gsCodUser, lsmensaje, gImpresora)
                
    If Trim(lsmensaje) <> "" Then
         MsgBox lsmensaje, vbInformation, "Aviso"
         Exit Sub
    End If
    Set loColPRecGar = Nothing
          
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Planila de Venta en Lote (Adjudicados)", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Inicializa el formulario actual y carga los datos del último Venta
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Call CargaParametros
    Call CargaDatosUltimaVenta
End Sub


Private Sub CargaDatosUltimaVenta()
Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrDatosVen As ADODB.Recordset
Dim lrAdj As ADODB.Recordset
Dim lsUltVenta As String
Dim lsmensaje As String

    txtPreOro14 = Format(0, "#0.00")
    txtPreOro16 = Format(0, "#0.00")
    txtPreOro18 = Format(0, "#0.00")
    txtPreOro21 = Format(0, "#0.00")
        
    Set loDatos = New COMNColoCPig.NCOMColPRecGar
    
    lsUltVenta = loDatos.nObtieneNroUltimoProceso("V", fsProcesoCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set lrDatosVen = loDatos.nObtieneDatosProcesoRGCredPig("V", lsUltVenta, fsProcesoCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
        
    If lrDatosVen!nRGEstado = geColPRecGarEstNoIniciado And Format(lrDatosVen!dProceso, "dd/mm/yyyy") <> Format(gdFecSis, "dd/mm/yyyy") Then
        Set lrDatosVen = Nothing
        Set lrDatosVen = loDatos.nObtieneDatosProcesoRGCredPig("V", FillNum(Trim(Str(Val(lsUltVenta) - 1)), 4, "0"), fsProcesoCadaAgencia, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If lrDatosVen Is Nothing Then Exit Sub
    
    If lrDatosVen!nRGEstado = geColPRecGarEstTerminado Then
        cmdAplicar.Enabled = False
        cmdProcesar.Enabled = False
        fraCuentas.Enabled = False
        cmdPlaVenta.Enabled = True
    Else
        Set lrAdj = loDatos.ObtieneAdjudicacionesParaVentaEnLote(fsProcesoCadaAgencia)
        Set loDatos = Nothing
        lstAdjudicacion.Clear
        While Not lrAdj.EOF
                lstAdjudicacion.AddItem lrAdj!cNroProceso & "  -  " & CStr(lrAdj!dProceso)
                lstAdjudicacion.ItemData(lstAdjudicacion.NewIndex) = lrAdj!cNroProceso
                lrAdj.MoveNext
        Wend
    End If
    
    
    txtNumVenta = lrDatosVen!cNroProceso
    txtFecVenta = Format(lrDatosVen!dProceso, "dd/mm/yyyy")
    txtHorVenta = Format(lrDatosVen!dProceso, "hh:mm")
    txtPreOro14 = Format(lrDatosVen!nPrecioK14, "#0.00")
    txtPreOro16 = Format(lrDatosVen!nPrecioK16, "#0.00")
    txtPreOro18 = Format(lrDatosVen!nPrecioK18, "#0.00")
    txtPreOro21 = Format(lrDatosVen!nPrecioK21, "#0.00")
        
    If lrDatosVen!nRGEstado = gColPRecGarEstNoIniciado Then
        txtEstado = "NO INICIADO"
        If Val(txtPreOro14) = 0 And Val(txtPreOro16) = 0 And Val(txtPreOro18) = 0 And Val(txtPreOro21) = 0 Then
            
            MsgBox " Regrese a Preparacion de Venta a ingresar precio del Oro ", vbInformation, " Aviso "
        ElseIf Format(lrDatosVen!dProceso, "dd/mm/yyyy") <> Format(gdFecSis, "dd/mm/yyyy") Then
            
            MsgBox " No es la fecha para el inicio de la venta en lote ", vbInformation, " Aviso "
        End If
    ElseIf lrDatosVen!nRGEstado = gColPRecGarEstIniciado Then
        txtEstado = "INICIADO"
        
    ElseIf lrDatosVen!nRGEstado = gColPRecGarEstTerminado Then
        txtEstado = "FINALIZADO"
    End If
        

    
    Set lrDatosVen = Nothing
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub


'Valida el campo txtFecVenta
Private Sub txtFecVenta_GotFocus()
    fEnfoque txtFecVenta
End Sub
Private Sub txtFecVenta_LostFocus()
If Not ValFecha(txtFecVenta) Then
    txtFecVenta.SetFocus
End If
End Sub

'Valida el campo txtHorVenta
Private Sub txtHorVenta_GotFocus()
    fEnfoque txtHorVenta
End Sub


' Procedimiento que genera la planilla de venta del Venta
Public Sub ImprimePlanVenSub()
'    Dim vIndice As Integer  'contador de Item
'    Dim vLineas As Integer
'    Dim vPage As Integer
'    Dim vTelefono As String
'    Dim vSumTasaci As Currency
'    Dim vSumPresta As Currency
'    Dim vSumDeuda As Currency
'    Dim vSumBase As Currency
'    Dim vSumVenta As Currency
'    Dim vVerCodAnt As String
'    MousePointer = 11
'    MuestraImpresion = True
'    vRTFImp = ""
'    sSql = " SELECT cp.cCodCta, cp.dfecvenc, cp.nValTasac, " & _
'        " cp.nPrestamo, cp.nTasaIntVenc, cp.nImpuesto, " & _
'        " cp.nCostCusto, ds.ndeuda, ds.nprebasevta, ds.npreventa " & _
'        " FROM CredPrenda CP JOIN DetSubas DS ON cp.ccodcta = ds.ccodcta " & _
'        " WHERE ds.cestado IN ('V','P') AND ds.cNroSubas = '" & Trim(txtNumVenta) & "'" & _
'        " ORDER BY ds.ccodcta"
'    RegCredPrend.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        MsgBox " No existen ningún contrato en la Venta ", vbInformation, " Aviso "
'        MuestraImpresion = False
'        MousePointer = 0
'        Exit Sub
'    Else
'        vPage = 1
'        prgList.Min = 0: vCont = 0
'        prgList.Max = RegCredPrend.RecordCount
'        If optImpresion(0).value = True Then
'            If prgList.Max > pPrevioMax Then
'                RegCredPrend.Close
'                Set RegCredPrend = Nothing
'                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                MuestraImpresion = False
'                MousePointer = 0
'                Exit Sub
'            End If
'            Cabecera "PlanVeSu", vPage
'        Else
'            ImpreBegin True, pHojaFiMax
'            vRTFImp = ""
'            Cabecera "PlanVeSu", vPage
'            Print #ArcSal, ImpreCarEsp(vRTFImp);
'            vRTFImp = ""
'        End If
'        prgList.Visible = True
'        vIndice = 1:        vLineas = 7
'        vSumTasaci = 0:        vSumPresta = 0
'        vSumDeuda = 0:        vSumBase = 0
'        vSumVenta = 0
'        With RegCredPrend
'            Do While Not .EOF
'                'Para Ver el Código Antiguo
'                If pVerCodAnt Then vVerCodAnt = ContratoAntiguo(!cCodCta)
'                If optImpresion(0).value = True Then
'                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & Format(!dFecVenc, "dd/mm/yyyy") & ImpreFormat(Round(!nvaltasac, 2), 10) & ImpreFormat(Round(!nPrestamo, 2), 10) & _
'                        ImpreFormat(Round(!ndeuda, 2), 10) & ImpreFormat(Round(!nprebasevta, 2), 10) & ImpreFormat(Round(!npreventa, 2), 10) & Chr(10)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        vRTFImp = vRTFImp & Space(7) & ImpreFormat(vVerCodAnt, 13, 0) & Chr(10)
'                        vLineas = vLineas + 1
'                    End If
'                Else
'                    Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & Format(!dFecVenc, "dd/mm/yyyy") & ImpreFormat(Round(!nvaltasac, 2), 10) & ImpreFormat(Round(!nPrestamo, 2), 10) & _
'                        ImpreFormat(Round(!ndeuda, 2), 10) & ImpreFormat(Round(!nprebasevta, 2), 10) & ImpreFormat(Round(!npreventa, 2), 10)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        Print #ArcSal, Space(7) & ImpreFormat(vVerCodAnt, 13, 0)
'                        vLineas = vLineas + 1
'                    End If
'                End If
'                vLineas = vLineas + 1
'                vSumTasaci = vSumTasaci + Round(!nvaltasac, 2)
'                vSumPresta = vSumPresta + Round(!nPrestamo, 2)
'                vSumDeuda = vSumDeuda + Round(!ndeuda, 2)
'                vSumBase = vSumBase + Round(!nprebasevta, 2)
'                vSumVenta = vSumVenta + Round(!npreventa, 2)
'                vIndice = vIndice + 1
'                If vLineas >= 55 Then
'                    vPage = vPage + 1
'                    If optImpresion(0).value = True Then
'                        vRTFImp = vRTFImp & Chr(12)
'                        Cabecera "PlanVeSu", vPage
'                    Else
'                        If vPage Mod 5 = 0 Then
'                            ImpreEnd
'                            ImpreBegin True, pHojaFiMax
'                        Else
'                            ImpreNewPage
'                        End If
'                        vRTFImp = ""
'                        Cabecera "PlanVeSu", vPage
'                        Print #ArcSal, ImpreCarEsp(vRTFImp);
'                        vRTFImp = ""
'                    End If
'                    vLineas = 7
'                End If
'                vCont = vCont + 1
'                prgList.value = vCont
'                .MoveNext
'            Loop
'            If optImpresion(0).value = True Then
'                vRTFImp = vRTFImp & Chr(10)
'                vRTFImp = vRTFImp & Space(12) & "RESUMEN " & ImpreFormat(vSumTasaci, 20) & ImpreFormat(vSumPresta, 10) & _
'                    ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumBase, 10) & ImpreFormat(vSumVenta, 10) & Chr(10)
'                rtfImp.Text = vRTFImp
'            Else
'                Print #ArcSal, ""
'                Print #ArcSal, Space(12) & "RESUMEN " & ImpreFormat(vSumTasaci, 20) & ImpreFormat(vSumPresta, 10) & _
'                    ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumBase, 10) & ImpreFormat(vSumVenta, 10)
'                ImpreEnd
'            End If
'        End With
'        prgList.Visible = False
'        prgList.value = 0
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'    End If
'    MousePointer = 0
End Sub

'Cabecera de las Impresiones
Private Sub Cabecera(ByVal vOpt As String, ByVal vPagina As Integer, Optional ByVal pPagCorta As Boolean = True)
    Dim vTitulo As String
    Dim vSubTit As String
    Dim vArea As String * 30
    Dim vNroLineas As Integer
    Select Case vOpt
        Case "PlanVeSu"
            vTitulo = "PLANILLA DE PRENDAS DE LA Venta N°: " & Format(txtNumVenta, "@@@@") & " DEL " & txtFecVenta
    End Select
    vSubTit = "(Después de la Venta)"
    vArea = "Crédito Pignoraticio"
    vNroLineas = IIf(pPagCorta = True, 105, 162)
    'Centra Título
    vTitulo = String(Round((vNroLineas - Len(Trim(vTitulo))) / 2), " ") & vTitulo
    'Centra SubTítulo
    vSubTit = String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ") & vSubTit & String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ")

    vRTFImp = vRTFImp & Chr(10)
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(gsNomAge, 25, 0) & Space(vNroLineas - 40) & "Página: " & Format(vPagina, "@@@@") & Chr(10)
    vRTFImp = vRTFImp & Space(1) & vTitulo & Chr(10)
    vRTFImp = vRTFImp & Space(1) & vArea & vSubTit & Space(7) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & Chr(10)
    vRTFImp = vRTFImp & String(vNroLineas, "-") & Chr(10)
    Select Case vOpt
        Case "PlanVeSu"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO     FECHA        TASACION     CAPITAL        DEUDA        BASE        VENTA" & Chr(10)
            vRTFImp = vRTFImp & Space(1) & "                    VENCIMI.                   PRESTADO                    VENTA        NETA" & Chr(10)
    End Select
    vRTFImp = vRTFImp & String(vNroLineas, "-") & Chr(10)
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

    pPrevioMax = 5000
    pLineasMax = 56
    pHojaFiMax = 66

    Set loParam = New COMDColocPig.DCOMColPCalculos
        fnTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    Set loParam = Nothing
          
    Set loConstSis = New COMDConstSistema.NCOMConstSistema
        lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)
        If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
            fsProcesoCadaAgencia = gsCodCMAC & gsCodAge
        Else
            fsProcesoCadaAgencia = gsCodCMAC & "00"
        End If
    Set loConstSis = Nothing
End Sub

Private Sub OptGasto_Click(Index As Integer)
Dim i As Integer
    If FECredito.Rows <= 1 Then
        Exit Sub
    End If
    If FECredito.TextMatrix(1, 2) = "" Then
        Exit Sub
    End If
    If Index = 0 Then
        For i = 1 To FECredito.Rows - 1
            FECredito.TextMatrix(i, 1) = "1"
        Next i
    Else
        For i = 1 To FECredito.Rows - 1
            FECredito.TextMatrix(i, 1) = "0"
        Next i
    End If
End Sub

Private Sub CmdProcesar_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loColPRecGar As COMNColoCPig.NCOMColPRecGar
Dim oNCred As COMNCredito.NCOMCredito
Dim i As Integer
Dim lMatCuentas() As Variant
Dim lnTamMat As Long, lnPosI As Long
Dim lsFecVenta As String, lsMovNro As String, lsFechaHoraGrab As String, lsmensaje As String
    
    If Trim(FECredito.TextMatrix(1, 2)) = "" Then
        MsgBox "No existen creditos para venta en lote", vbInformation, "Aviso"
        Exit Sub
    End If
    
    For i = 1 To FECredito.Rows - 1
        If FECredito.TextMatrix(i, 1) = "." Then
            lnTamMat = lnTamMat + 1
        End If
    Next i
    
    If lnTamMat = 0 Then
        MsgBox "No a seleccionado ningún crédito, es necesario que por lo menos un crédito se encuentre seleccionado", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lblTotalS.Caption = "Nro. Total de Cuentas Seleccionadas : " & lnTamMat
    
    If MsgBox("Se va a procesar la venta en lote  de las siguientes adjudicaciones : " & fsListaAgencias & _
            " a los " & CStr(lnTamMat) & " créditos adjudicados seleccionados, tener en cuenta que una vez procesado no podrá revertirse, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
     
    'ReDim lMatCuentas(lnTamMat, 7)
    ReDim lMatCuentas(lnTamMat, 8) 'PEAC 20090918
    
    lnPosI = 0
    For i = 1 To FECredito.Rows - 1
        If FECredito.TextMatrix(i, 1) = "." Then
            lMatCuentas(lnPosI, 0) = FECredito.TextMatrix(i, 2) 'cuenta
            lMatCuentas(lnPosI, 1) = FECredito.TextMatrix(i, 5) 'saldo
            lMatCuentas(lnPosI, 2) = FECredito.TextMatrix(i, 6) 'tasacion
            lMatCuentas(lnPosI, 3) = FECredito.TextMatrix(i, 7) 'k14
            lMatCuentas(lnPosI, 4) = FECredito.TextMatrix(i, 8) 'k16
            lMatCuentas(lnPosI, 5) = FECredito.TextMatrix(i, 9) 'k18
            lMatCuentas(lnPosI, 6) = FECredito.TextMatrix(i, 10) 'k21
            lMatCuentas(lnPosI, 7) = FECredito.TextMatrix(i, 11) 'Valor Registro - PEAC 20090918
            lnPosI = lnPosI + 1
        End If
    Next i
     
    lsFecVenta = Format$(txtFecVenta.Text, "mm/dd/yyyy")
    
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
        
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set loColPRecGar = New COMNColoCPig.NCOMColPRecGar
    Call loColPRecGar.VentaEnLote(lsMovNro, lsFechaHoraGrab, lMatCuentas, Trim(txtNumVenta.Text), lsFecVenta, _
            Val(txtPreOro14.Text), Val(txtPreOro16.Text), Val(txtPreOro18.Text), Val(txtPreOro21.Text), fsProcesoCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
         MsgBox lsmensaje, vbInformation, "Aviso"
         Exit Sub
    End If
    Set loColPRecGar = Nothing
    
    txtEstado.Text = "TERMINADO"
    cmdAplicar.Enabled = False
    cmdProcesar.Enabled = False
    fraCuentas.Enabled = False
    cmdPlaVenta.Enabled = True
    
    MsgBox "El proceso de venta en lote a finalizado", vbInformation, "Mensaje"
    
    Set oNCred = Nothing
End Sub

