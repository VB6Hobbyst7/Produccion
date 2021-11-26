VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5F774E03-DB36-4DFC-AAC4-D35DC9379F2F}#1.1#0"; "VertMenu.ocx"
Begin VB.Form frmCajaIntDeveng 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   285
   ClientTop       =   1935
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcajaGenIntDeven.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar-->"
      Height          =   345
      Left            =   9840
      TabIndex        =   29
      Top             =   5610
      Width           =   1215
   End
   Begin VB.Frame fraTransferencia 
      Caption         =   "Transferencia a :"
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
      Height          =   945
      Left            =   1425
      TabIndex        =   14
      Top             =   5025
      Visible         =   0   'False
      Width           =   7005
      Begin VB.CheckBox chkCapitaliza 
         Caption         =   "Capitalizar &Intereses"
         Height          =   210
         Left            =   3510
         TabIndex        =   27
         Top             =   0
         Width           =   1875
      End
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   315
         Left            =   915
         TabIndex        =   8
         Top             =   225
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkCarta 
         Caption         =   "&No Emitir Carta"
         Height          =   210
         Left            =   5460
         TabIndex        =   26
         Top             =   0
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta N° :"
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblDescIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3510
         TabIndex        =   16
         Top             =   225
         Width           =   3285
      End
      Begin VB.Label lblDesCtaIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   915
         TabIndex        =   15
         Top             =   585
         Width           =   5880
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9810
      TabIndex        =   10
      Top             =   5190
      Width           =   1245
   End
   Begin VertMenu.VerticalMenu VMInteres 
      Height          =   5790
      Left            =   150
      TabIndex        =   12
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   10213
      MenuCaption1    =   "Opciones..."
      MenuItemsMax1   =   5
      MenuItemIcon11  =   "frmcajaGenIntDeven.frx":030A
      MenuItemCaption11=   "Efectivo"
      MenuItemIcon12  =   "frmcajaGenIntDeven.frx":0624
      MenuItemCaption12=   "Ingreso Cheque"
      MenuItemIcon13  =   "frmcajaGenIntDeven.frx":093E
      MenuItemCaption13=   "Giro Cheque"
      MenuItemIcon14  =   "frmcajaGenIntDeven.frx":0C58
      MenuItemCaption14=   "Transferencia"
      MenuItemIcon15  =   "frmcajaGenIntDeven.frx":0F72
      MenuItemCaption15=   "Nota Abono"
   End
   Begin VB.Frame fradatosGen 
      Caption         =   "Datos Generales"
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
      Height          =   3510
      Left            =   1425
      TabIndex        =   13
      Top             =   75
      Width           =   9780
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   870
         TabIndex        =   0
         Top             =   210
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame fraInteres 
         Height          =   570
         Left            =   3435
         TabIndex        =   23
         Top             =   2865
         Width           =   6120
         Begin VB.TextBox txtCalculado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   4440
            TabIndex        =   5
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1590
         End
         Begin VB.TextBox txtInteres 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   1065
            TabIndex        =   4
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1830
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Calculado :"
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
            Height          =   195
            Left            =   3360
            TabIndex        =   25
            Top             =   225
            Width           =   975
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   3150
            Top             =   165
            Width           =   2895
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Interes :"
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
            Height          =   195
            Left            =   255
            TabIndex        =   24
            Top             =   225
            Width           =   720
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   150
            Top             =   165
            Width           =   2760
         End
      End
      Begin VB.Frame fraCapital 
         Height          =   570
         Left            =   60
         TabIndex        =   21
         Top             =   2865
         Width           =   3300
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   1485
            TabIndex        =   3
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1680
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capital :"
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
            Height          =   195
            Left            =   330
            TabIndex        =   22
            Top             =   240
            Width           =   720
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   150
            Top             =   165
            Width           =   3045
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   345
         Left            =   8205
         TabIndex        =   1
         Top             =   180
         Width           =   1410
      End
      Begin Sicmact.FlexEdit fgIF 
         Height          =   2265
         Left            =   45
         TabIndex        =   2
         Top             =   585
         Width           =   9645
         _extentx        =   17013
         _extenty        =   3995
         cols0           =   12
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "N°-N° Cuenta-Institución Financiera-Capital-Tasa %-Interés-Fecha Int-Dias-Calculado-cPersCod-cIftpo-cCtaIFCod"
         encabezadosanchos=   "350-2000-1800-1200-700-900-800-500-1000-0-0-0"
         font            =   "frmcajaGenIntDeven.frx":128C
         font            =   "frmcajaGenIntDeven.frx":12B4
         font            =   "frmcajaGenIntDeven.frx":12DC
         font            =   "frmcajaGenIntDeven.frx":1304
         font            =   "frmcajaGenIntDeven.frx":132C
         fontfixed       =   "frmcajaGenIntDeven.frx":1354
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X-8-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-R-R-R-L-R-R-L-L-C"
         formatosedit    =   "0-0-0-2-2-2-0-3-2-0-0-0"
         textarray0      =   "N°"
         lbeditarflex    =   -1
         lbformatocol    =   -1
         lbpuntero       =   -1
         lbordenacol     =   -1
         colwidth0       =   345
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interes al:"
         Height          =   210
         Left            =   105
         TabIndex        =   18
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Frame FraConcepto 
      Height          =   1395
      Left            =   1425
      TabIndex        =   19
      Top             =   3540
      Width           =   9780
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   7755
         TabIndex        =   7
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   750
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   195
         Width           =   9315
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL :"
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
         Height          =   195
         Left            =   6600
         TabIndex        =   20
         Top             =   1065
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         FillColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   6420
         Top             =   990
         Width           =   3045
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   8550
      TabIndex        =   11
      Top             =   5190
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptaInt 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   8550
      TabIndex        =   9
      Top             =   5190
      Width           =   1275
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   255
      Left            =   8820
      OleObjectBlob   =   "frmcajaGenIntDeven.frx":137A
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmCajaIntDeveng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpcion As Long
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim lnTipoCtaIf As CGTipoCtaIF
Dim lbCancelaCuenta As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim objPista As COMManejador.Pista 'ARLO20170217

Function Valida() As Boolean
Valida = True
If fraCapital.Visible Then
    If Val(txtCapital) = 0 Then
        MsgBox "Monto Capital no Calculado", vbInformation, "Aviso"
        Valida = False
        fgIF.SetFocus
        Exit Function
    End If
End If

If fraInteres.Visible Then
    If Val(txtInteres) = 0 Then
        If MsgBox("Monto interes no válido o es cero. Desea proseguir??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            fgIF.SetFocus
            Valida = False
            Exit Function
        End If
    End If
    If Val(txtCalculado) = 0 Then
        If MsgBox("Monto interes Calculado no válido o es cero. Desea proseguir??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
           fgIF.SetFocus
           Valida = False
           Exit Function
        End If
    End If
End If
If fraTransferencia.Visible Then
    If Len(Trim(txtBuscaEntidad)) = 0 Then
        MsgBox "Cuenta de Entidad financiera no válida", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no válida", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Valida = False
    Exit Function
End If
If Val(txtImporte) = 0 Then
    MsgBox "Monto de Operación no válida", vbInformation, "Aviso"
    txtImporte.SetFocus
    Valida = False
    Exit Function
End If

End Function

Private Sub chkCapitaliza_Click()
If chkCapitaliza.value = vbChecked Then
    Me.txtBuscaEntidad.Text = Me.fgIF.TextMatrix(fgIF.Row, 10) & "." & Me.fgIF.TextMatrix(fgIF.Row, 9) & "." & Me.fgIF.TextMatrix(fgIF.Row, 11)
    Me.lblDescIfTransf = Me.fgIF.TextMatrix(fgIF.Row, 2)
    Me.lblDesCtaIfTransf = Me.fgIF.TextMatrix(fgIF.Row, 1)
Else
    Me.txtBuscaEntidad.Text = ""
    Me.lblDescIfTransf = ""
    Me.lblDesCtaIfTransf = ""
End If
End Sub

Private Sub cmdAceptaInt_Click()
Dim oCont As NContFunciones
Dim oCaja As nCajaGeneral
Dim lsMovNro As String
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsPersCod As String
Dim lsCtaIFCod As String
Dim lnDiasTrans As Integer
Dim lnTipoIf As CGTipoIF

Set oCont = New NContFunciones
Set oCaja = New nCajaGeneral

lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , fgIF.TextMatrix(fgIF.Row, 10) & "." & fgIF.TextMatrix(fgIF.Row, 9) & "." & fgIF.TextMatrix(fgIF.Row, 11), ObjEntidadesFinancieras)
lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , fgIF.TextMatrix(fgIF.Row, 10) & "." & fgIF.TextMatrix(fgIF.Row, 9) & "." & fgIF.TextMatrix(fgIF.Row, 11), ObjEntidadesFinancieras)
If Val(txtCalculado) <= 0 Then
    MsgBox "Interes Calculado no válido", vbInformation, "Aviso"
    fgIF.SetFocus
    Exit Sub
End If
If Val(txtImporte) <= 0 Then
    MsgBox "Importe de Operación no válido", vbInformation, "Aviso"
    fgIF.SetFocus
    Exit Sub
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Movimiento no Ingresada", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If

If MsgBox("Desea Registrar intereses devengados de Cuenta Seleccionada??", vbYesNo + vbQuestion, "AViso") = vbYes Then
    lsPersCod = fgIF.TextMatrix(fgIF.Row, 9)
    lnTipoIf = fgIF.TextMatrix(fgIF.Row, 10)
    lsCtaIFCod = fgIF.TextMatrix(fgIF.Row, 11)
    lnDiasTrans = fgIF.TextMatrix(fgIF.Row, 7)
    
    lsMovNro = oCont.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    oCaja.GrabaInteresDev lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, lsCtaHaber, _
            txtImporte, lsPersCod, lnTipoIf, lsCtaIFCod, txtFecha, lnDiasTrans
    
    ImprimeAsientoContable lsMovNro
    
    If MsgBox("Desea Registrar Otra Operación de Registro de Interés??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        fgIF.EliminaFila fgIF.Row
        txtCapital = "0.00"
        txtInteres = "0.00"
        txtCalculado = "0.00"
        txtImporte = "0.00"
        txtMovDesc = ""
        fgIF.SetFocus
    Else
        Unload Me
    End If
End If
End Sub
Private Sub cmdAceptar_Click()
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset
Dim oDocPago As clsDocPago
Dim lsDocumento As String
Dim lnTpoDoc As TpoDoc
Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim lsCuentaAho As String
Dim lnMotivo As MotivoNotaAbonoCargo
Dim lsObjetoPadre As String
Dim lsObjeto As String
Dim lsPersNombre As String
Dim lsPersDireccion As String
Dim lsUbigeo As String

Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim oCaja As nCajaGeneral

'variables para el ingreso de cheques
Dim lsPersCodIFChq As String
Dim lsNroCtaIfChq As String
Dim lsNroChq As String
Dim lnPlazaChq As ChequePlaza
Dim ldFechaRegChq As Date
Dim ldFechaValChq As Date
Dim lsConfCheque As String
Dim lsGlosa As String
Dim lsNombreIF As String
Dim lnImporte As Currency
Dim rsObjMot As ADODB.Recordset
Dim lsProductoCod As String
Dim lsAreaAgeCod As String
Dim lsCtaMotivo As String
Dim lnMonedaChq As Moneda

Dim lsCtaIntCalculado As String
Dim lsCtaInteres As String

Dim lsMovNro As String
Dim oCon As NContFunciones
Dim lsPersCodIf As String
Dim lnTipoIf As CGTipoIF
Dim lsCtaIFCod As String
Dim lnDiasTrans As Integer
Dim lsCtaContDifEfec As String
Dim lsCtaHaberCap As String
Dim lnImporteDif As Currency
Dim lsCtaIFCapital As String
Dim lnPlazo As Integer

On Error GoTo AceptarErr
Set oDocPago = New clsDocPago
If Valida = False Then Exit Sub

lsDocumento = ""
lnTpoDoc = -1
lsNroDoc = ""
lsNroVoucher = ""
Select Case lnOpcion
    Case 1 'EFECTIVO
            frmCajaGenEfectivo.inicio gsOpeCod, gsOpeDesc, CCur(txtImporte), Mid(gsOpeCod, 3, 1), False, True
            If frmCajaGenEfectivo.lbOk Then
                 Set rsBill = frmCajaGenEfectivo.rsBilletes
                 Set rsMon = frmCajaGenEfectivo.rsMonedas
                 lnImporteDif = frmCajaGenEfectivo.vnDiferencia
            Else
                Set frmCajaGenEfectivo = Nothing
                Exit Sub
            End If
            Set frmCajaGenEfectivo = Nothing
            If rsBill Is Nothing And rsMon Is Nothing Then
                MsgBox "Error en Ingreso de Billetaje", vbInformation, "Aviso"
                Exit Sub
            End If
    Case 2 'INGRESO CHEQUE
            'EJVG20140415 ***
            'Set frmIngCheques = Nothing
            'lnTpoDoc = TpoDocCheque
            
            'frmIngCheques.Inicio False, gsOpeCod, True, CCur(txtImporte), Mid(gsOpeCod, 3, 1), , 3, 4, True, txtMovDesc
            ' If frmIngCheques.OK Then
            '    lnMonedaChq = frmIngCheques.Moneda
            '    lsPersCodIFChq = frmIngCheques.PersCodIF
            '    lsNroCtaIfChq = frmIngCheques.NroCtaIf
            '    lsNroChq = frmIngCheques.NroChq
            '    lnPlazaChq = frmIngCheques.PlazaChq
            '    ldFechaRegChq = frmIngCheques.FechaRegChq
            '    ldFechaValChq = frmIngCheques.FechaValChq
            '    lsConfCheque = frmIngCheques.ConfCheque
            '    lsGlosa = frmIngCheques.Glosa
            '    lsNombreIF = frmIngCheques.NombreIF
            '    lnImporte = frmIngCheques.Importe
            '    Set rsObjMot = frmIngCheques.rsObjMotivo
            '    lsProductoCod = frmIngCheques.ProductoCod
            '    lsAreaAgeCod = frmIngCheques.AreaAgeCod
            '    lsCtaMotivo = frmIngCheques.CtaMotivo
                
            '    Unload frmIngCheques
            '    Set frmIngCheques = Nothing
            'Else
            '    Unload frmIngCheques
            '    Set frmIngCheques = Nothing
            '    Exit Sub
            'End If
            Exit Sub
            'END EJVG *******
            
    Case 3 'GIRO CHEQUE
        Set oCon = New NContFunciones
        lsNroVoucher = oCon.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
        Set oCon = Nothing
        oDocPago.InicioCheque "", True, Mid(txtBuscaEntidad, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, _
                     CCur(txtImporte), gdFecSis, gsNomCmacRUC, txtBuscaEntidad, lblDescIfTransf, _
                     lblDesCtaIfTransf, lsNroVoucher, True, , Mid(txtBuscaEntidad, 18, 10), , Mid(txtBuscaEntidad, 1, 2), Mid(txtBuscaEntidad, 4, 13), Mid(txtBuscaEntidad, 18, 10) 'EJVG20121130
                     'lblDesCtaIfTransf , lsNroVoucher, True, , Mid(txtBuscaEntidad, 18, 10)
        If oDocPago.vbOk Then
            lsDocumento = oDocPago.vsFormaDoc
            lnTpoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsNroVoucher = oDocPago.vsNroVoucher
        Else
            Exit Sub
        End If
        
    Case 4 'EMISION CARTA
        If Not chkCarta.value = vbChecked Then
            oDocPago.InicioCarta "", "", gsOpeCod, gsOpeDesc, txtMovDesc, "", CCur(txtImporte), _
                         gdFecSis, lblDescIfTransf, lblDesCtaIfTransf, gsNomCmac, "", ""
            If oDocPago.vbOk Then
                lsDocumento = oDocPago.vsFormaDoc
                lnTpoDoc = Val(oDocPago.vsTpoDoc)
                lsNroDoc = oDocPago.vsNroDoc
                lsNroVoucher = oDocPago.vsNroVoucher
            Else
                Exit Sub
            End If
        Else
            lnTpoDoc = -1
        End If
    Case 5 'NOTA DE ABONO
            lnTpoDoc = TpoDocNotaAbono
            frmNotaCargoAbono.inicio lnTpoDoc, CCur(txtImporte), gdFecSis, txtMovDesc, gsOpeCod, False
            
            If frmNotaCargoAbono.vbOk Then
                lsNroDoc = frmNotaCargoAbono.NroNotaCA
                lsDocumento = frmNotaCargoAbono.NotaCargoAbono
                lsPersNombre = frmNotaCargoAbono.PersNombre
                lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
                lnMotivo = frmNotaCargoAbono.Motivo
                lsObjetoPadre = frmNotaCargoAbono.ObjetoMotivoPadre
                lsObjeto = frmNotaCargoAbono.ObjetoMotivo
                lsPersDireccion = frmNotaCargoAbono.PersDireccion
                lsUbigeo = frmNotaCargoAbono.PersUbigeo
                
                Unload frmNotaCargoAbono
                Set frmNotaCargoAbono = Nothing
            Else
                Unload frmNotaCargoAbono
                Set frmNotaCargoAbono = Nothing
                Exit Sub
            End If
    Case 0
        Select Case gsOpeCod
            Case gOpeCGOpeCapIntPFMN, gOpeCGOpeCapIntPFME, gOpeCGOpeCMACCapIntDevPFMN, gOpeCGOpeCMACCapIntDevPFME
                MsgBox "Por favor Seleccione destino de la Capitalización de la Cuenta", vbInformation, "Aviso"
            Case Else
                MsgBox "Por favor Seleccione destino de la Cancelación de la Cuenta", vbInformation, "Aviso"
        End Select
        VMInteres.SetFocus
        Exit Sub
End Select
Set oCon = New NContFunciones
Set oCaja = New nCajaGeneral

lsPersCodIf = fgIF.TextMatrix(fgIF.Row, 9)
lnTipoIf = fgIF.TextMatrix(fgIF.Row, 10)
lsCtaIFCod = fgIF.TextMatrix(fgIF.Row, 11)
lnDiasTrans = fgIF.TextMatrix(fgIF.Row, 7)
lnPlazo = nVal(fgIF.TextMatrix(fgIF.Row, 12))

lsCtaIFCapital = Format(lnTipoIf, "00") + "." + lsPersCodIf + "." + lsCtaIFCod
lsCtaHaberCap = ""
If MsgBox(" ¿ Desea Grabar Datos ? ", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    If lnTipoCtaIf = gTpoCtaIFCtaPF Then
       lsCtaInteres = oOpe.EmiteOpeCta(gsOpeCod, "H", 0, lsCtaIFCapital, ObjEntidadesFinancieras)
       lsCtaIntCalculado = oOpe.EmiteOpeCta(gsOpeCod, "H", 1, lsCtaIFCapital, ObjEntidadesFinancieras)
       
       If Not oOpe.ValidaCtaCont(lsCtaInteres) Then
          MsgBox "Falta definir Cuenta Contable de Interes Provisionado Valida " & lsCtaInteres & " - " & lsCtaIFCapital, vbInformation, "¡Aviso!"
          Exit Sub
       End If
       If Not oOpe.ValidaCtaCont(lsCtaIntCalculado) Then
          MsgBox "Falta definir Cuenta Contable de Interes Calculado Valida " & lsCtaIntCalculado, vbInformation, "¡Aviso!"
          Exit Sub
       End If
    End If
    lsCtaHaberCap = oOpe.EmiteOpeCta(gsOpeCod, "H", 2, lsCtaIFCapital, ObjEntidadesFinancieras)
    If Not oOpe.ValidaCtaCont(lsCtaHaberCap) Then
       MsgBox "Falta definir Cuenta Contable de Capital", vbInformation, "¡Aviso!"
       Exit Sub
    End If
    
    Select Case lnOpcion
        Case 1 'EFECTIVO
            lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", Trim(Str(lnOpcion - 1)))
            If Not oOpe.ValidaCtaCont(lsCtaDebe) Then
                MsgBox "Falta definir Cuenta Contable para Abonar Capital en Orden " & Trim(Str(lnOpcion - 1)) & " " & lsCtaDebe, vbInformation, "¡Aviso!"
                Exit Sub
            End If
            lsCtaContDifEfec = oOpe.EmiteOpeCta(gsOpeCod, "D", 9)
            If lsCtaContDifEfec = "" And lnImporteDif <> 0 Then
                MsgBox "Falta definir Cuenta Contable para Ajustar Efectivo en Orden 9D", vbInformation, "¡Aviso!"
                Exit Sub
            End If
            
            oCaja.GrabaCapitalizaIFEfectivo lsMovNro, gsOpeCod, txtMovDesc, _
                rsBill, rsMon, lsCtaDebe, txtImporte, lsCtaContDifEfec, lnImporteDif, _
                lsCtaInteres, CCur(txtInteres), lsCtaIntCalculado, CCur(txtCalculado), _
                lsCtaHaberCap, CCur(txtCapital), lsPersCodIf, lnTipoIf, lsCtaIFCod, txtFecha, lnDiasTrans, lnPlazo, lbCancela
        
            ImprimeAsientoContable lsMovNro, , , , True, True
        Case 2 'INGRESO CHEQUE
            lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", Trim(Str(lnOpcion - 1)))
            lsCtaDebe = lsCtaDebe + gsCodAge
            If lsCtaDebe = "" Then
                MsgBox "Falta definir Cuenta Contable de Cheques en Orden " & Trim(Str(lnOpcion - 1)), vbInformation, "¡Aviso!"
                Exit Sub
            End If
            
            oCaja.GrabaCapitalizaRegCheque lsMovNro, gsOpeCod, lsGlosa, txtFecha, _
                lsCtaDebe, lsProductoCod, Mid(lsAreaAgeCod, 1, 3), Mid(lsAreaAgeCod, 4, 2), lsCtaInteres, _
                lsCtaIntCalculado, lsCtaHaberCap, CCur(txtCapital), CCur(txtInteres), CCur(txtCalculado), rsObjMot, lsPersCodIf, lnTipoIf, lsCtaIFCod, _
                lsNroChq, Mid(lsPersCodIFChq, 4, 13), Mid(lsPersCodIFChq, 1, 2), lnPlazaChq, lsNroCtaIfChq, CCur(txtImporte), ldFechaRegChq, _
                ldFechaValChq, lnMonedaChq, gChqEstEnValorizacion, gCGEstadosChqRecibido, ChqCGSinConfirmacion, _
                gsCodArea, gsCodAge, lnDiasTrans, lnPlazo, lbCancela
            
            ImprimeAsientoContable lsMovNro
            
        Case 3 'GIRO CHEQUE
            lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", Trim(Str(lnOpcion - 1)), txtBuscaEntidad, ObjEntidadesFinancieras)
            If lsCtaDebe = "" Then
                MsgBox "Falta definir Cuenta Contable para Abonar Capital en Orden " & Trim(Str(lnOpcion - 1)), vbInformation, "¡Aviso!"
                Exit Sub
            End If
            
            oCaja.GrabaCapitalizaDoc lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, _
                 CCur(txtImporte), ObjEntidadesFinancieras, txtBuscaEntidad, _
                 lsCtaInteres, CCur(txtInteres), lsCtaIntCalculado, CCur(txtCalculado), _
                 lsCtaHaberCap, CCur(txtCapital), lsPersCodIf, lnTipoIf, lsCtaIFCod, _
                 txtFecha, lnTpoDoc, lsNroDoc, txtFecha, lsNroVoucher, "", -1, "", "", lnDiasTrans, lnPlazo, lbCancela
                 
            ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento
            
        Case 4 'EMISION CARTA
            lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", Trim(Str(lnOpcion - 2)), txtBuscaEntidad, ObjEntidadesFinancieras)
            If Not oOpe.ValidaCtaCont(lsCtaDebe) Then
                MsgBox "Falta definir Cuenta Contable para Abonar Capital en Orden " & lsCtaDebe & " " & Trim(Str(lnOpcion - 2)), vbInformation, "¡Aviso!"
                Exit Sub
            End If
            oCaja.GrabaCapitalizaDoc lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, _
                 CCur(txtImporte), ObjEntidadesFinancieras, txtBuscaEntidad, _
                 lsCtaInteres, CCur(txtInteres), lsCtaIntCalculado, CCur(txtCalculado), _
                 lsCtaHaberCap, CCur(txtCapital), lsPersCodIf, lnTipoIf, lsCtaIFCod, _
                 txtFecha, lnTpoDoc, lsNroDoc, txtFecha, lsNroVoucher, "", -1, "", "", lnDiasTrans, lnPlazo, lbCancela
            Dim nTasa As Double
            Dim nDias As Integer
            nTasa = IIf(IsNumeric(fgIF.TextMatrix(fgIF.Row, 4)), fgIF.TextMatrix(fgIF.Row, 4), 0) 'nTasa
            nDias = IIf(IsNumeric(fgIF.TextMatrix(fgIF.Row, 4)), fgIF.TextMatrix(fgIF.Row, 7), 0) 'nDias
            
            ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento, , , , , , , , , , , , , , , , nTasa, nDias
            
        Case 5 'NOTA DE ABONO
            lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", Trim(Str(lnOpcion - 1)))
            If Not oOpe.ValidaCtaCont(lsCtaDebe) Then
                MsgBox "Falta definir Cuenta Contable para Abonar en Orden " & lsCtaDebe & " " & Trim(Str(lnOpcion - 1)), vbInformation, "¡Aviso!"
                Exit Sub
            End If

            oCaja.GrabaCapitalizaDoc lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, _
                 CCur(txtImporte), -1, "", _
                 lsCtaInteres, CCur(txtInteres), lsCtaIntCalculado, CCur(txtCalculado), _
                 lsCtaHaberCap, CCur(txtCapital), lsPersCodIf, lnTipoIf, lsCtaIFCod, _
                 txtFecha, lnTpoDoc, lsNroDoc, txtFecha, lsNroVoucher, _
                 lsCuentaAho, lnMotivo, lsObjetoPadre, lsObjeto, lnDiasTrans, lnPlazo, lbCancela
            
            Dim oContImp As NContImprimir
            Set oContImp = New NContImprimir
            lsDocumento = oContImp.ImprimeNotaCargoAbono(lsNroDoc, txtMovDesc, CCur(txtImporte), _
                    lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), _
                    lsCuentaAho, lnTpoDoc, gsNomAge, gsCodUser)
            
            Set oContImp = Nothing
            
            ImprimeAsientoContable lsMovNro, , lnTpoDoc, lsDocumento
        Case 0
            Exit Sub
    End Select
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
    If MsgBox("¿¿ Desea Realizar otra operación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        fgIF.EliminaFila fgIF.Row
        txtCalculado = "0.00"
        txtCapital = "0.00"
        txtInteres = "0.00"
        txtImporte = "0.00"
        txtBuscaEntidad = ""
        txtMovDesc = ""
        lblDesCtaIfTransf = ""
        lblDescIfTransf = ""
        Set rsBill = Nothing
        Set rsMon = Nothing
        
    Else
        Unload Me
    End If
End If
Exit Sub
AceptarErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdExportar_Click()
    Dim sMoneda As String
    Dim sFinan As String
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    If Mid(gsOpeCod, 3, 1) = "1" Then
        sMoneda = "MN"
    Else
        sMoneda = "ME"
    End If
    If Me.fgIF.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    If gsOpeCod = "401518" Or gsOpeCod = "402518" Then
        sFinan = "Bancos"
    Else
        sFinan = "Cajas"
    End If
    lsArchivoN = App.path & "\Spooler\" & sFinan & "IntDev" & sMoneda & Format(txtFecha, "yyyy") & Format(txtFecha, "mm") & Format(txtFecha, "dd") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
        xlHoja1.Columns("B:B").NumberFormat = "0"
        xlHoja1.Columns("J:J").NumberFormat = "0"
       GeneraReporteCajasME fgIF, xlHoja1
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
       OleExcel.Appearance = 0
       OleExcel.Width = 500
    End If
    MousePointer = 0

End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
On Error GoTo ProcesarErr
Select Case lnTipoCtaIf
    Case gTpoCtaIFCtaPF
        Set rs = oCtaIf.GetInteresDevengados(gsOpeCod, txtFecha, , gsCodCMAC, VMInteres.Visible)
    Case Else
        Set rs = oCtaIf.GetCuentasIFCanc(gsOpeCod, txtFecha)
End Select
fgIF.Clear
fgIF.FormaCabecera
fgIF.Rows = 2
If Not rs.EOF And Not rs.BOF Then
    Set fgIF.Recordset = rs
    '*****************Modificado PASI20151104
    'CargarRS fgIF, rs, gTpoCtaIFCtaPF 'RIRO20140430 ERS017
    CargarRS fgIF, rs, lnTipoCtaIf 'RIRO20140430 ERS017
    '*************end PASI
    fgIF.SetFocus
Else
    MsgBox "No existe información con la Fecha Ingresada.", vbOKOnly + vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing
Exit Sub
ProcesarErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgIF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If fraCapital.Visible Then
        txtCapital.SetFocus
    Else
        txtInteres.SetFocus
    End If
End If
End Sub

Private Sub fgIF_RowColChange()
If fgIF.TextMatrix(1, 0) <> "" Then
    If Me.fraCapital.Visible Then
        txtCapital = fgIF.TextMatrix(fgIF.Row, 3)
    End If
    If Me.fraInteres.Visible Then
        txtInteres = fgIF.TextMatrix(fgIF.Row, 5)
    End If
    'txtCalculado = fgIF.TextMatrix(fgIF.Row, 8) RIRO20140903
    txtCalculado = Round(fgIF.TextMatrix(fgIF.Row, 8), 2)
    CalculaTotal
End If
End Sub
Sub CalculaTotal()
Select Case gsOpeCod
    Case gOpeCGOpeCancCtaCteMN, gOpeCGOpeCancCtaCteME, _
         gOpeCGOpeCancCtaPFMN, gOpeCGOpeCancCtaPFME, _
         gOpeCGOpeCancCtaAhoMN, gOpeCGOpeCancCtaAhoME, _
         gOpeCGOpeCMACCancAhorrosMN, gOpeCGOpeCMACCancAhorrosME, _
         gOpeCGOpeCMACCancPFMN, gOpeCGOpeCMACCancPFME
         
        txtImporte = Format(CCur(txtCapital) + CCur(txtCalculado) + CCur(txtInteres), "#,#0.00")
    Case gOpeCGOpeIntDevPFMN, gOpeCGOpeIntDevPFME, gOpeCGOpeCMACIntDevPFMN, gOpeCGOpeCMACIntDevPFME
        txtImporte = Format(CCur(txtCalculado), "#,#0.00")
    Case Else
        txtImporte = Format(CCur(txtCalculado) + CCur(txtInteres), "#,#0.00")
End Select
End Sub
Private Sub Form_Load()
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
Me.Caption = gsOpeDesc

txtBuscaEntidad.psRaiz = "Cuentas de Entidades Financieras"
txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
txtFecha = gdFecSis
lnTipoCtaIf = -1
cmdAceptaInt.Visible = False
cmdAceptar.Visible = False
fraCapital.Visible = False
Me.fraInteres.Visible = True
txtCapital.Locked = True
txtInteres.Locked = True
txtCalculado.Locked = True
txtImporte.Locked = True
Select Case gsOpeCod
    Case gOpeCGOpeIntDevPFMN, gOpeCGOpeIntDevPFME, gOpeCGOpeCMACIntDevPFMN, gOpeCGOpeCMACIntDevPFME
        lnTipoCtaIf = gTpoCtaIFCtaPF
        cmdAceptaInt.Visible = True
        VMInteres.Visible = False
        Me.fradatosGen.Left = 100
        Me.FraConcepto.Left = 100
        cmdAceptaInt.Left = cmdAceptaInt.Left - (cmdAceptaInt.Width + 50)
        cmdSalir.Left = cmdSalir.Left - (cmdSalir.Width + 50)
        Me.Width = 10000
        Me.Height = 6000
        lbCancelaCuenta = False
        
    Case gOpeCGOpeCapIntPFMN, gOpeCGOpeCapIntPFME, gOpeCGOpeCMACCapIntDevPFMN, gOpeCGOpeCMACCapIntDevPFME
        lnTipoCtaIf = gTpoCtaIFCtaPF
        txtCalculado.Locked = False
        cmdAceptar.Visible = True
        lbCancelaCuenta = False
    
    Case gOpeCGOpeCancCtaAhoMN, gOpeCGOpeCancCtaAhoME, gOpeCGOpeCMACCancAhorrosMN, gOpeCGOpeCMACCancAhorrosME
        lnTipoCtaIf = gTpoCtaIFCtaAho
        fraCapital.Visible = True
        fraInteres.Visible = False
        cmdAceptar.Visible = True
        fgIF.EncabezadosAnchos = "350-3000-4000-1500-0-0-0-0-0-0-0"
        lbCancelaCuenta = True
        
    Case gOpeCGOpeCancCtaCteMN, gOpeCGOpeCancCtaCteME
        lnTipoCtaIf = gTpoCtaIFCtaCte
        fraCapital.Visible = True
        fraInteres.Visible = False
        cmdAceptar.Visible = True
        fgIF.EncabezadosAnchos = "350-3000-4000-1500-0-0-0-0-0-0-0"
        lbCancelaCuenta = True
        
    Case gOpeCGOpeCancCtaPFMN, gOpeCGOpeCancCtaPFME, gOpeCGOpeCMACCancPFMN, gOpeCGOpeCMACCancPFME
        lnTipoCtaIf = gTpoCtaIFCtaPF
        txtCalculado.Locked = False
        fraCapital.Visible = True
        cmdAceptar.Visible = True
        lbCancelaCuenta = True
End Select
CentraForm Me
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
lblDesCtaIfTransf = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
If txtBuscaEntidad <> "" Then
    If cmdAceptar.Visible Then cmdAceptar.SetFocus
    If cmdAceptaInt.Visible Then cmdAceptaInt.SetFocus
End If
End Sub

Private Sub txtCalculado_GotFocus()
fEnfoque txtCalculado
End Sub

Private Sub txtCalculado_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCalculado, KeyAscii)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub txtCalculado_LostFocus()
If Val(txtCalculado) = 0 Then txtCalculado = 0
txtCalculado = Format(txtCalculado, "#,#0.00")
CalculaTotal
End Sub

Private Sub txtcapital_GotFocus()
fEnfoque txtCapital
End Sub

Private Sub txtcapital_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCapital, KeyAscii)
If KeyAscii = 13 Then
    If fraInteres.Visible Then
        txtInteres.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub
Private Sub txtcapital_LostFocus()
If Trim(Len(txtCapital)) = 0 Then txtCapital = 0
txtCapital = Format(txtCapital, "#,#0.00")
CalculaTotal
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) Then
        fgIF.Clear
        fgIF.FormaCabecera
        fgIF.Rows = 2
        cmdProcesar.SetFocus
    End If
End If
End Sub
Private Sub txtFecha_Validate(Cancel As Boolean)
If ValFecha(txtFecha) Then
    Cancel = False
End If

End Sub

Private Sub txtImporte_GotFocus()
fEnfoque txtImporte
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 18, 2)
If KeyAscii = 13 Then
    If fraTransferencia.Visible = False Then
        If cmdAceptar.Visible Then cmdAceptar.SetFocus
        If cmdAceptaInt.Visible Then cmdAceptaInt.SetFocus
    Else
        txtBuscaEntidad.SetFocus
    End If
End If
End Sub
Private Sub txtImporte_LostFocus()
If Val(txtImporte) = 0 Then txtImporte = 0
txtImporte = Format(txtImporte, "#,#0.00")
End Sub
Private Sub txtInteres_GotFocus()
fEnfoque txtInteres
End Sub
Private Sub txtInteres_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtInteres, KeyAscii)
If KeyAscii = 13 Then
    txtCalculado.SetFocus
End If
End Sub
Private Sub txtInteres_LostFocus()
If Val(txtInteres) = 0 Then txtInteres = 0
txtInteres = Format(txtInteres, "#,#0.00")
CalculaTotal
End Sub
'Private Sub txtInteres_GotFocus()
'fEnfoque txtInteres
'End Sub
'Private Sub txtInteres_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosDecimales(txtInteres, KeyAscii)
'If KeyAscii = 13 Then
'    spnPeriodo.SetFocus
'End If
'End Sub
'Private Sub txtInteres_LostFocus()
'If Val(txtInteres) = 0 Then txtInteres = 0
'txtInteres = Format(txtInteres, "#,#0.00")
'End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtImporte.SetFocus
End If
End Sub
Private Sub VMInteres_MenuItemClick(MenuNumber As Long, MenuItem As Long)
fraTransferencia.Visible = False
txtBuscaEntidad = ""
lblDescIfTransf = ""
lblDesCtaIfTransf = ""
lnOpcion = -1
Select Case MenuNumber
    Case 1
        lnOpcion = MenuItem
        Select Case MenuItem
            Case 1 'EFECTIVO
                txtMovDesc.SetFocus
            Case 2 'INGRESO CHEQUE
            Case 3 'GIRO CHEQUE
                fraTransferencia.Visible = True
                txtBuscaEntidad.SetFocus
                chkCarta.Visible = False
            Case 4 'EMISION CARTA
                fraTransferencia.Visible = True
                txtBuscaEntidad.SetFocus
                chkCarta.Visible = True
            Case 5 'NOTA DE ABONO
            Case 6 'SIN SALDO
            Case 7 'OTROS
        End Select
    Case Else
End Select
If fraTransferencia.Visible And lbCancelaCuenta Then
    chkCapitaliza.Visible = False
End If
End Sub
Public Sub GeneraReporteCajasME(pflex As FlexEdit, pxlHoja1 As Excel.Worksheet, Optional pnColFiltroVacia As Integer = 0)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    For I = 0 To pflex.Rows - 1
        If pnColFiltroVacia = 0 Then
            For j = 0 To pflex.Cols - 1
                pxlHoja1.Cells(I + 1, j + 1) = pflex.TextMatrix(I, j)
            Next j
        Else
            If pflex.TextMatrix(I, pnColFiltroVacia) <> "" Then
                For j = 0 To pflex.Cols - 1
                    pxlHoja1.Cells(I + 1, j + 1) = pflex.TextMatrix(I, j)
                Next j
            End If
        End If
    Next I
    
End Sub

'RIRO20140430 ERS017 ***
Private Sub CargarRS(ByRef pFgIF As FlexEdit, ByVal rsCuentas As ADODB.Recordset, ByVal nTipo As CGTipoCtaIF)
    If rsCuentas Is Nothing Then Exit Sub
    Dim I As Long
    'pFgIF.ColWidth(11) = 0
    'pFgIF.ColWidth(12) = 0
    'pFgIF.ColWidth(13) = 0
    If nTipo = gTpoCtaIFCtaPF Then
        For I = 1 To rsCuentas.RecordCount
            pFgIF.TextMatrix(I, 4) = Format(rsCuentas!intvalor, "#,##0.00000")
            pFgIF.TextMatrix(I, 8) = Format(rsCuentas!intcalculado, "#,#0.00000")
            rsCuentas.MoveNext
        Next
    Else
        For I = 1 To rsCuentas.RecordCount
            pFgIF.TextMatrix(I, 4) = Format(rsCuentas!valorint, "#,##0.00000")
            pFgIF.TextMatrix(I, 8) = Format(rsCuentas!intcalculado, "#,#0.00000")
            rsCuentas.MoveNext
        Next
    End If
End Sub
'END RIRO **************
