VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPresGenPlanPend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adeudados: Proyecciones"
   ClientHeight    =   3765
   ClientLeft      =   780
   ClientTop       =   1440
   ClientWidth     =   9375
   Icon            =   "frmPresGenPlanPend.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar Proyectado"
      Height          =   390
      Left            =   2407
      TabIndex        =   5
      Top             =   3195
      Width           =   2220
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar Proyectado"
      Height          =   390
      Left            =   165
      TabIndex        =   4
      Top             =   3195
      Width           =   2220
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar Plan >>"
      Height          =   390
      Left            =   4650
      TabIndex        =   3
      Top             =   3195
      Width           =   1815
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7500
      TabIndex        =   2
      Top             =   3210
      Width           =   1320
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   315
      Left            =   930
      TabIndex        =   0
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraListaAdeuProy 
      Caption         =   "Lista de Adeudados Proyectados en Presupuesto"
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
      Height          =   2595
      Left            =   180
      TabIndex        =   6
      Top             =   480
      Width           =   9045
      Begin MSDataGridLib.DataGrid dtgAduedProy 
         Height          =   2265
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   3995
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   17
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "cCtaIFDesc"
            Caption         =   "N° Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cCtaIFCod"
            Caption         =   "N° Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cPersNombre"
            Caption         =   "Entidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "nMontoPrestado"
            Caption         =   "Monto Prestado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """S/"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "dCtaIFAper"
            Caption         =   "Fecha Apert"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "dCtaIFVenc"
            Caption         =   "Fecha Venc"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "nCtaIFCuotas"
            Caption         =   "Cuotas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "nCtaIFPlazo"
            Caption         =   "Plazo Cuota"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "nTasaInteres"
            Caption         =   "Interes"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "nComisionMonto"
            Caption         =   "Comisión Ini"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "nCtaIFIntPeriodo"
            Caption         =   "Periodo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2789.858
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraIngAdeud 
      Caption         =   "Datos de Pagaré"
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
      Height          =   2625
      Left            =   180
      TabIndex        =   8
      Top             =   450
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox txtTotApertura 
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   6960
         TabIndex        =   15
         Tag             =   "0"
         Text            =   "0.00"
         Top             =   2040
         Width           =   1650
      End
      Begin VB.TextBox txtCapital 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         Left            =   1020
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1590
         Width           =   1500
      End
      Begin VB.TextBox txtComisionMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3450
         TabIndex        =   13
         Text            =   "0"
         Top             =   2040
         Width           =   1410
      End
      Begin VB.TextBox txtComisionInicial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Text            =   "0"
         Top             =   2040
         Width           =   750
      End
      Begin VB.TextBox txtCuentaCod 
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1140
         Width           =   1005
      End
      Begin VB.TextBox txtCuentaDesc 
         Height          =   315
         Left            =   3030
         TabIndex        =   10
         Top             =   1140
         Width           =   2295
      End
      Begin Sicmact.TxtBuscar txtLinCredCod 
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         Top             =   1590
         Width           =   1485
         _extentx        =   2619
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmPresGenPlanPend.frx":030A
         appearance      =   1
         stitulo         =   ""
      End
      Begin Sicmact.TxtBuscar txtBuscaIF 
         Height          =   360
         Left            =   1020
         TabIndex        =   25
         Top             =   285
         Width           =   1995
         _extentx        =   3519
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmPresGenPlanPend.frx":0336
         appearance      =   1
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Institución Financiera :"
         Height          =   420
         Left            =   150
         TabIndex        =   27
         Top             =   270
         Width           =   900
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescTipoCta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3030
         TabIndex        =   26
         Top             =   690
         Width           =   5595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Neto a Desembolsar"
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
         Left            =   5130
         TabIndex        =   24
         Top             =   2070
         Width           =   1725
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Capital :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   150
         TabIndex        =   23
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label lblComision 
         Alignment       =   2  'Center
         Caption         =   "Monto"
         Height          =   285
         Left            =   2940
         TabIndex        =   22
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Comision Inicial :"
         Height          =   210
         Left            =   150
         TabIndex        =   21
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label Label5 
         Caption         =   "Tasa"
         Height          =   285
         Left            =   1590
         TabIndex        =   20
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "Código"
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
         Height          =   285
         Left            =   2310
         TabIndex        =   18
         Top             =   1155
         Width           =   645
      End
      Begin VB.Label Label9 
         Caption         =   "Linea de Crédito"
         Height          =   225
         Left            =   2970
         TabIndex        =   17
         Top             =   1635
         Width           =   1335
      End
      Begin VB.Label lblLinCredDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5700
         TabIndex        =   16
         Top             =   1590
         Width           =   2925
      End
      Begin VB.Label lblDescIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3030
         TabIndex        =   28
         Top             =   300
         Width           =   5595
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   5010
         Top             =   2025
         Width           =   3615
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   660
   End
End
Attribute VB_Name = "frmPresGenPlanPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim rs As New ADODB.Recordset
Dim oCon As DConecta
Dim lbEjecutado As Boolean
Dim lsNomCta  As String
Private Type Calendario
    Cuota As Long
    FechaPago As Date
    Capital As Currency
    Interes As Currency
    Comisión As Currency
    Total As Currency
    Saldo As Currency
End Type
Dim ltCalendario() As Calendario

Dim lsCodObjetoProy As String
Dim lnCapital As Currency
Dim lnCuotas As Integer
Dim lnPlazoCuota As Integer
Dim lnPeriodo As Long
Dim ldFechaIni As Date
Dim lnGracia As Integer
Dim lnInteres As Currency
Dim lnComisionIni As Currency
Dim lnComisionCuota As Currency
Dim lsImpre As String
Dim I As Integer
Dim lsFiltroIF As String
Dim lsObjeto As String
Dim lnCuotaK As Integer
Dim lbCreaSubCta As Boolean
Dim lsSubCtaIF As String
Dim lnTipoCtaIf As CGTipoCtaIF
Dim oNAdeudCal As nAdeudCal
Dim rsAdeud    As ADODB.Recordset
Dim lbCalendario As Boolean

Private Sub cmdAgregar_Click()
If cmdAgregar.Caption = "&Agregar Proyectado" Then
    Limpia
    Me.fraListaAdeuProy.Visible = False
    fraIngAdeud.Visible = True
    fraIngAdeud.Enabled = True
    Me.cmdAgregar.Caption = "&Adicionar a Presupuesto"
    cmdEliminar.Caption = "&Cancelar"
    txtBuscaIF.SetFocus
Else
    Dim oCon     As NContFunciones
    Dim oCaja As nCajaGeneral
    Dim lnTpoCuota As CGAdeudCalTpoCuota
    
    If Valida = False Then
       Exit Sub
    End If
    
    Set rsAdeud = frmAdeudCal.fgCronograma.GetRsNew(1)
    If frmAdeudCal.optTpoCuota(0).value Then
     lnTpoCuota = gAdeudTpoCuotaFija
    Else
     lnTpoCuota = gAdeudTpoCuotaVariable
    End If

    If MsgBox("Desea Grabar Operación de Apertura??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        Set oCon = New NContFunciones
        Set oCaja = New nCajaGeneral
        gsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
        Set oCon = Nothing
        oCaja.GrabaAdeudadoProyectado Mid(txtBuscaIF, 4, 13), Mid(txtBuscaIF, 1, 2), _
               Me.txtCuentaCod, txtCuentaDesc, txtFecha, lbCreaSubCta, lsSubCtaIF, lblDescTipoCta, _
               IIf(frmAdeudCal.optPeriodo(0).value, 360, 30), frmAdeudCal.txtInteres, nVal(txtCapital), _
               frmAdeudCal.SpnCuotas.Valor, frmAdeudCal.SpnGracia.Valor, txtComisionInicial, nVal(txtComisionMonto), IIf(frmAdeudCal.chkInterno.value = 1, "0", "1"), frmAdeudCal.txtCuotaPagoK, IIf(frmAdeudCal.chkVac = vbChecked And Mid(gsOpeCod, 3, 1) = "1", gMonedaExtranjera, Mid(gsOpeCod, 3, 1)), lnTpoCuota, CCur(frmAdeudCal.txtTramo), _
               nVal(frmAdeudCal.txtComision), txtLinCredCod, frmAdeudCal.txtFechaCuota, frmAdeudCal.txtPlazoCuotas, rsAdeud, gsMovNro
        RSClose rsAdeud
        Limpia
        CargaProyectados
        fraListaAdeuProy.Visible = True
        fraIngAdeud.Visible = False
        fraIngAdeud.Enabled = False
        cmdEliminar.Caption = "&Eliminar Proyectado"
        cmdAgregar.Caption = "&Agregar Proyectado"
    End If
End If
End Sub
Private Sub Limpia()
txtBuscaIF = ""
lblDescIF = ""
txtCuentaCod = ""
txtCuentaDesc = ""
txtCapital = "0.00"
txtLinCredCod = ""
lblLinCredDesc = ""
txtComisionInicial = "0.00"
txtComisionMonto = "0.00"
txtTotApertura = "0.00"
End Sub

Function Valida() As Boolean
Dim K As Integer
Valida = False
If Len(Trim(txtBuscaIF)) = 0 Then
    MsgBox "Entidad Financiera no seleccionada", vbInformation, "Aviso"
    txtBuscaIF.SetFocus
    Exit Function
End If
If nVal(txtTotApertura) = 0 Then
    MsgBox "No se indicó Monto del Préstamo. Por favor Verifique", vbInformation, "Aviso"
    txtCapital.SetFocus
    Exit Function
End If

If Not lbCalendario Then
    MsgBox "No se ha definido calendario de pagos para pagaré respectivo", vbInformation, "Aviso"
    Me.cmdGenerar.SetFocus
    Exit Function
End If
Valida = True
End Function

Private Sub cmdEliminar_Click()
Dim sql As String
Dim lbTransActiva As Boolean
Dim oMov As DMov
Dim lsMsgErr As String
On Error GoTo cmdEliminarErr
If cmdEliminar.Caption = "&Eliminar Proyectado" Then
    If rs.EOF Then
        MsgBox "No existen Registros para eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿ Desea Eliminar el Adeudado : [" & rs!cCtaIFDesc & "] ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oMov = New DMov
        lbTransActiva = True
        oMov.BeginTrans
        oMov.EliminaCuentaIFInteres rs!cPersCod, rs!ciftpo, rs!cCtaIfCod
        oMov.EliminaCuentaIFPresup rs!cPersCod, rs!ciftpo, rs!cCtaIfCod
        oMov.EliminaCuentaIFAdeudado rs!cPersCod, rs!ciftpo, rs!cCtaIfCod
        oMov.EliminaCuentaIF rs!cPersCod, rs!ciftpo, rs!cCtaIfCod
        oMov.CommitTrans
        lbTransActiva = False
        CargaProyectados
    End If
Else
    Limpia
    fraListaAdeuProy.Visible = True
    fraIngAdeud.Visible = False
    fraIngAdeud.Enabled = False
    cmdEliminar.Caption = "&Eliminar Proyectado"
    cmdAgregar.Caption = "&Agregar Proyectado"
End If
Exit Sub
cmdEliminarErr:
lsMsgErr = Err.Description
    If lbTransActiva Then
        oMov.RollbackTrans
    End If
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdGenerar_Click()
Dim lsCtaIF As String
Dim lsCtaIFDesc As String
Dim lnCapital As Currency
Dim ldFecha   As Date
Dim lbGetAdeudado As Boolean
If Me.fraIngAdeud.Visible Then
    lsCtaIF = txtBuscaIF + "." + txtCuentaCod
    lsCtaIFDesc = lblDescIF
    lnCapital = nVal(txtCapital)
    ldFecha = txtFecha
    lbGetAdeudado = False
Else
    If rs.EOF Then
        MsgBox "No Existen Adeudados para ver Cronograma", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    lsCtaIF = rs!ciftpo & "." & rs!cPersCod & "." & rs!cCtaIfCod
    lsCtaIFDesc = rs!cCtaIFDesc
    lnCapital = rs!nMontoPrestado
    ldFecha = rs!dCtaIFAper
    lbGetAdeudado = True
End If
If lsCtaIF <> "" And lnCapital > 0 Then
    frmAdeudCal.Inicio True, lsCtaIF, lsCtaIFDesc, lnCapital, ldFecha, , lbGetAdeudado
    If frmAdeudCal.OK Then
        lbCalendario = True
    Else
        lbCalendario = False
        Set frmAdeudCal = Nothing
    End If
Else
    MsgBox "Debe indicar Institución e Importe del Adeudado", vbInformation, "¡Aviso!"
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dtgAduedProy_GotFocus()
dtgAduedProy.MarqueeStyle = dbgHighlightRow
End Sub
Private Sub dtgAduedProy_LostFocus()
dtgAduedProy.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAgregar.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim N As Integer
Set oCon = New DConecta
Set oNAdeudCal = New nAdeudCal
oCon.AbreConexion

txtFecha = gdFecSis
CentraForm Me
Dim oOpe As New DOperacion
txtBuscaIF.psRaiz = "INSTITUCIONES FINANCIERAS"
txtBuscaIF.rs = oOpe.GetOpeObj(gsOpeCod, "0")
Set rs = oOpe.CargaOpeObj(gsOpeCod, , "1")
If rs.EOF Then
    MsgBox "Falta definir Filtro en Operación para seleccionar Adeudados Proyectados", vbInformation, "¡Aviso!"
Else
    lsFiltroIF = rs!cOpeObjFiltro
End If
RSClose rs
CargaProyectados
Set oOpe = Nothing
Me.lblDescTipoCta = gsOpeDescHijo
lnTipoCtaIf = Val(Mid(lsFiltroIF, 3, 2))
lbCreaSubCta = False
lbEjecutado = False
N = 0
Limpia
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub CargaProyectados()
Dim lnIndiceVac As Double
Dim oIF As New NCajaAdeudados
Dim oDAdeud As DCaja_Adeudados
    Set oDAdeud = New DCaja_Adeudados
    lnIndiceVac = oDAdeud.CargaIndiceVAC(CDate("01/" & Format(Month(gdFecSis), "00") & "/" & Format(Year(gdFecSis), "0000")) - 1)
    Set oDAdeud = Nothing
    Set rs = oIF.CargaDatosAdeudadosProyectados("", , lnIndiceVac, lsFiltroIF)
    Set dtgAduedProy.DataSource = rs
End Sub

Private Sub txtBuscaIF_EmiteDatos()
Dim oAdeud As New NCajaAdeudados
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
lblDescIF = txtBuscaIF.psDescripcion
lsSubCtaIF = ""
lbCreaSubCta = False
If txtBuscaIF <> "" Then
    txtLinCredCod = ""
    lblLinCredDesc = ""
    txtLinCredCod.rs = oAdeud.GetLineaCredito(Mid(gsOpeCod, 3, 1), Mid(txtBuscaIF, 4, 13))
    lbCreaSubCta = Not oCtaIf.GetVerificaSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
    lsSubCtaIF = oCtaIf.GetSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
    txtCuentaCod = oCtaIf.GetNewCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1), lsSubCtaIF)
    txtCuentaDesc = "Pendiente"
    txtCuentaDesc.SetFocus
End If
Set oAdeud = Nothing
End Sub

Private Sub txtcapital_GotFocus()
fEnfoque txtCapital
End Sub

Private Sub txtcapital_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCapital, KeyAscii, 15, 2)
If KeyAscii = 13 Then
    txtCapital = Format(txtCapital, gsFormatoNumeroView)
    txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
    txtLinCredCod.SetFocus
End If
End Sub

Private Sub txtCapital_Validate(Cancel As Boolean)
txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
End Sub

Private Sub txtComisionInicial_GotFocus()
fEnfoque txtComisionInicial
End Sub

Private Sub txtComisionInicial_KeyPress(KeyAscii As Integer)
If nVal(txtCapital) = 0 Then
    MsgBox "Primero ingresar Monto de Prestamo", vbInformation, "¡Aviso!"
    txtCapital.SetFocus
    Exit Sub
End If
KeyAscii = NumerosDecimales(txtComisionInicial, KeyAscii, 10, 4)
If KeyAscii = 13 Then
   txtComisionMonto = Format(Round(nVal(txtCapital) * txtComisionInicial / 100, 2), gsFormatoNumeroView)
   txtComisionInicial = Format(txtComisionInicial, gsFormatoNumeroView)
   txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
   txtComisionMonto.SetFocus
End If
End Sub

Private Sub txtComisionMonto_GotFocus()
fEnfoque txtComisionMonto
End Sub

Private Sub txtComisionMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtComisionMonto, KeyAscii, 14, 4)
If KeyAscii = 13 Then
   txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
   cmdGenerar.SetFocus
End If
End Sub

Private Sub txtComisionMonto_Validate(Cancel As Boolean)
txtTotApertura = Format(nVal(txtCapital) - nVal(txtComisionMonto), gsFormatoNumeroView)
End Sub

Private Sub txtCuentaCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCuentaDesc.SetFocus
End If
End Sub

Private Sub txtCuentaDesc_GotFocus()
fEnfoque txtCuentaDesc
End Sub

Private Sub txtCuentaDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCapital.SetFocus
End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If fraIngAdeud.Visible And fraIngAdeud.Enabled Then
        txtBuscaIF.SetFocus
    Else
        dtgAduedProy.SetFocus
    End If
End If
End Sub
Private Sub txtFecha_LostFocus()
If ValFecha(txtFecha) = False Then Exit Sub
End Sub
Private Sub txtFecha_Validate(Cancel As Boolean)
If ValFecha(txtFecha) = False Then Cancel = True
End Sub

Private Sub txtLinCredCod_EmiteDatos()
lblLinCredDesc = txtLinCredCod.psDescripcion
txtComisionInicial.SetFocus
End Sub

