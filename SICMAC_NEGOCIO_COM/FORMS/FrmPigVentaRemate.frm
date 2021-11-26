VERSION 5.00
Begin VB.Form FrmPigVentaRemate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas en Remate - Verificador y Bloqueo de Piezas"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "FrmPigVentaRemate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPieza 
      Caption         =   "Datos de Venta(Piezas)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   3075
      Left            =   135
      TabIndex        =   29
      Top             =   2745
      Width           =   9750
      Begin SICMACT.EditMoney txtValorVenta 
         Height          =   285
         Left            =   8115
         TabIndex        =   30
         Top             =   2670
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
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
         BackColor       =   -2147483624
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.FlexEdit fepiezas 
         Height          =   2400
         Left            =   90
         TabIndex        =   31
         Top             =   270
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   4233
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Item-Pieza-Tipo-Material-Descripcion-PNeto-ValBase-ValVenta-Cliente-prueba"
         EncabezadosAnchos=   "350-500-1100-1200-1250-700-1150-1150-2100-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-7-8-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-C-C-C-C-C-R-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-2-1-0"
         CantEntero      =   10
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label10 
         Caption         =   "Valor Vta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   7080
         TabIndex        =   32
         Top             =   2730
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4035
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   26
      Top             =   810
      Width           =   330
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   375
      Left            =   7455
      TabIndex        =   25
      Top             =   5955
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   24
      Top             =   5955
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   375
      Left            =   6270
      TabIndex        =   23
      Top             =   5970
      Width           =   1095
   End
   Begin SICMACT.EditMoney emvalrem 
      Height          =   285
      Left            =   8235
      TabIndex        =   22
      Top             =   1485
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      BackColor       =   -2147483637
      Text            =   "0"
   End
   Begin SICMACT.EditMoney emdeuda 
      Height          =   285
      Left            =   8235
      TabIndex        =   21
      Top             =   1125
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      BackColor       =   -2147483637
      Text            =   "0"
   End
   Begin SICMACT.EditMoney emreta 
      Height          =   285
      Left            =   8220
      TabIndex        =   20
      Top             =   765
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      BackColor       =   -2147483637
      Text            =   "0"
   End
   Begin VB.TextBox txtpneto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   1410
      TabIndex        =   19
      Top             =   1560
      Width           =   660
   End
   Begin VB.TextBox txtpieza 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   18
      Top             =   1245
      Width           =   660
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Venta(Lote)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   705
      Left            =   2760
      TabIndex        =   11
      Top             =   1995
      Width           =   6960
      Begin VB.TextBox txtcompra 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4305
         TabIndex        =   15
         Top             =   255
         Width           =   2520
      End
      Begin SICMACT.EditMoney emvalvtalote 
         Height          =   285
         Left            =   1170
         TabIndex        =   13
         Top             =   315
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   503
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
         BackColor       =   -2147483624
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label11 
         Caption         =   "Comprador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3300
         TabIndex        =   14
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label10 
         Caption         =   "Valor Vta."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rematado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   645
      Left            =   150
      TabIndex        =   5
      Top             =   2025
      Width           =   2130
      Begin VB.OptionButton oppieza 
         Caption         =   "Piezas"
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
         Left            =   1035
         TabIndex        =   7
         Top             =   315
         Width           =   915
      End
      Begin VB.OptionButton oplote 
         Caption         =   "Lote"
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
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   180
      TabIndex        =   0
      Top             =   75
      Width           =   9690
      Begin VB.TextBox txtubica 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4770
         TabIndex        =   4
         Top             =   180
         Width           =   4815
      End
      Begin VB.TextBox txtremate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   975
         TabIndex        =   3
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "UBICACION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "REMATE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   780
      End
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   300
      TabIndex        =   27
      Top             =   765
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.CommandButton cmdLote 
      Caption         =   "VENTAS POR LOTE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   165
      TabIndex        =   28
      Top             =   2805
      Width           =   9705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Index           =   1
      X1              =   150
      X2              =   9840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblEstadoCont 
      Height          =   330
      Left            =   3315
      TabIndex        =   33
      Top             =   5190
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   9825
      Y1              =   1905
      Y2              =   1905
   End
   Begin VB.Label Label6 
      Caption         =   "P.Neto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   405
      TabIndex        =   17
      Top             =   1620
      Width           =   705
   End
   Begin VB.Label Label5 
      Caption         =   "Piezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   405
      TabIndex        =   16
      Top             =   1275
      Width           =   690
   End
   Begin VB.Label Label9 
      Caption         =   "Valor de Remate"
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
      Left            =   6600
      TabIndex        =   10
      Top             =   1545
      Width           =   1425
   End
   Begin VB.Label Label8 
      Caption         =   "Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6585
      TabIndex        =   9
      Top             =   1185
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "Retasacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6585
      TabIndex        =   8
      Top             =   825
      Width           =   1035
   End
End
Attribute VB_Name = "FrmPigVentaRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lnTipoProceso As Integer
'
'Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
'Dim oContrato As DPigContrato
'Dim oDatos As DPigRemate
'Dim rsDatos As Recordset
'Dim psCuenta As String
'Dim rs As Recordset
'Dim Pieza As Boolean
'Dim lsEstado As String
'Dim nUbicaLote As Integer, lnUbicaRemate As Integer
'    Pieza = False
'    If KeyAscii = 13 Then
'
'        psCuenta = Me.AXCodCta.NroCuenta
'        Set oDatos = New DPigRemate
'        lnUbicaRemate = oDatos.GetUbicaLoteRemate(FrmPigVentaRemate.txtremate)
'        Set oContrato = New DPigContrato
'        Set rsDatos = oContrato.dObtieneContratosLotes(psCuenta, lnTipoProceso)
'        If (rsDatos.EOF And rsDatos.BOF) Then
'            MsgBox "Contrato no existe, o no se encuentra en Remate", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        lsEstado = rsDatos!nPrdEstado 'VERIFICADOR DE ESTADOS DE CONTRATO
'        nUbicaLote = rsDatos!nUbicaLote
'        If lsEstado = 2808 Or lsEstado = 2807 Or lsEstado = 2809 Then
'           If nUbicaLote = lnUbicaRemate Then
'                If Not (rsDatos.EOF And rsDatos.BOF) Then
'                    txtpieza.Text = rsDatos!npiezas
'                    txtpneto.Text = rsDatos!pesoneto
'                    emreta.Text = rsDatos!retasacion
'                    emdeuda.Text = rsDatos!valordeuda
'                    emvalrem.Text = rsDatos!valorproceso
'                    lblEstadoCont = rsDatos!nPrdEstado
'                    cmdGrabar.Enabled = True
'                Else
'                    MsgBox "Contrato no se encuentra en Remate o no tiene Piezas Disponibles para Remate", vbInformation, "Aviso"
'                   Exit Sub
'                End If
'                Set rsDatos = Nothing
'                Set rs = oContrato.dObtieneContratosPiezas(psCuenta, lnTipoProceso)
'                If Not (rs.EOF And rs.BOF) Then
'                        fepiezas.Clear
'                        fepiezas.Rows = 2
'                        fepiezas.FormaCabecera
'                        Do While Not rs.EOF
'                            fepiezas.AdicionaFila
'                            fepiezas.TextMatrix(fepiezas.Rows - 1, 1) = rs!nItemPieza
'                            fepiezas.TextMatrix(fepiezas.Rows - 1, 2) = rs!Descritipo
'                            fepiezas.TextMatrix(fepiezas.Rows - 1, 3) = rs!cConsDescripcion
'                            If IsNull(rs!cDescripcion) Then
'                                fepiezas.TextMatrix(fepiezas.Rows - 1, 4) = ""
'                            Else
'                                fepiezas.TextMatrix(fepiezas.Rows - 1, 4) = rs!cDescripcion
'                            End If
'                            fepiezas.TextMatrix(fepiezas.Rows - 1, 5) = rs!npesoneto
'                            fepiezas.TextMatrix(fepiezas.Rows - 1, 6) = rs!nValorProceso
'                            fepiezas.TextMatrix(fepiezas.Rows - 1, 7) = IIf(IsNull(rs!nValorventa), "", rs!nValorventa)
'                            If IsNull(rs!cComprador) Then
'                                fepiezas.TextMatrix(fepiezas.Rows - 1, 8) = ""
'                            Else
'                                fepiezas.TextMatrix(fepiezas.Rows - 1, 8) = rs!cComprador
'                                Pieza = True
'                            End If
'                            rs.MoveNext
'                        Loop
'                End If
'                If Pieza = True Then
'                    cmdLote.Visible = False
'                    frmPieza.Visible = True
'                    fepiezas.Enabled = True
'                    oppieza.value = True
'                    oplote.value = False
'                    emvalvtalote.Enabled = False
'                Else
'                    oplote.SetFocus
'                End If
'                suma_pieza
'                Set rs = Nothing
'                Set oContrato = Nothing
'           Else
'                   MsgBox "La ubicacion del Lote es distinta a la del Remate", vbExclamation, "Ubicacion"
'           End If
'         Else
'                MsgBox "El contrato no puede ser Rematado su Estado Actual no lo Permite", vbExclamation, "No se puede rematar"
'        End If
'
'        End If
'End Sub
'
'Private Sub cmdBuscar_Click()
'    'CRSF - 17/06
'    fepiezas.Clear
'    fepiezas.Rows = 2
'    fepiezas.FormaCabecera
'    FrmPigContratosRem.Show 1
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Limpia
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oGraba As NPigRemate
'Dim rs As Recordset
'Dim lnTipoVenta As Integer
'Dim lsCliente As String
'Dim psCuenta As String
'Dim lnValorVta As Currency
'Dim oCont As NContFunciones
'Dim lsMovNro As String
'Dim lsFechaHoraGrab As String
'Dim i As Integer
'
'    psCuenta = Me.AXCodCta.NroCuenta
'    Set rs = fepiezas.GetRsNew
'
'    Set oCont = New NContFunciones
'    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Set oCont = Nothing
'
'    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'
'    If oplote Then 'EN CASO DE QUE LA VENTA SEA POR LOTE
'        If txtcompra <> "" Then
'            lnTipoVenta = 1
'            lnValorVta = emvalvtalote
'            lsCliente = txtcompra.Text
'        Else
'            MsgBox "Debe Registrar al Comprador de las Piezas", vbInformation, "Aviso"
'            Exit Sub
'        End If
'
'    ElseIf oppieza Then
'        lnTipoVenta = 2
'        lnValorVta = txtValorVenta
'        lsCliente = ""
'
'        For i = 1 To Me.fepiezas.Rows - 1
'            If fepiezas.TextMatrix(i, 7) <> "" Then
'                If fepiezas.TextMatrix(i, 8) = "" And CCur(fepiezas.TextMatrix(i, 7) > 0) Then
'                    MsgBox "Debe ingresar al comprador", vbInformation, "Aviso"
'                    Exit Sub
'                End If
'            End If
'        Next i
'    End If
'
'    Set oGraba = New NPigRemate
'    oGraba.nPigVentaRemate psCuenta, lnValorVta, lnTipoVenta, lsFechaHoraGrab, rs, lnTipoProceso, lblEstadoCont, lsCliente
'    Set oGraba = Nothing
'
'    Limpia
'
'End Sub
'
'Private Sub Limpia()
'    txtpieza.Text = ""
'    txtpneto.Text = ""
'    emreta.Text = ""
'    emdeuda.Text = ""
'    emvalrem.Text = ""
'    emvalvtalote.Text = ""
'    txtcompra.Text = ""
'    oplote.value = True
'    AXCodCta.Cuenta = ""
'    AXCodCta.Age = ""
'    AXCodCta.SetFocusAge
'    cmdLote.Visible = True
'    fepiezas.Clear
'    fepiezas.Rows = 2
'    fepiezas.FormaCabecera
'    cmdGrabar.Enabled = False
'End Sub
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub emvalvtalote_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If emvalvtalote <> "" Then
'            If CCur(emvalrem) <= CCur(emvalvtalote) Then
'                  txtcompra.SetFocus
'            Else
'                  MsgBox "El Valor de Compra no puede ser menor al Valor del Remate"
'                  emvalvtalote.Text = 0
'                  emvalvtalote.SetFocus
'            End If
'        End If
'    End If
'End Sub
'Private Sub fepiezas_RowColChange()
'Dim Total As Double
'
'If fepiezas.Col = 8 Then
'    If fepiezas.TextMatrix(fepiezas.Row, 7) <> "" Then
'        'Determina si el Val.Vta > al Val. Base
'        If CCur(fepiezas.TextMatrix(fepiezas.Row, 6)) <= CCur(fepiezas.TextMatrix(fepiezas.Row, 7)) Then
'           If fepiezas.TextMatrix(fepiezas.Row, 8) = "" Then
'                 If CLng(txtValorVenta) >= CLng(emvalrem.Text) Then
'                    MsgBox "El Valor de la Venta ya cubrio la Deuda, No se puede Rematar Piezas", vbExclamation, "No Venta"
'                    fepiezas.TextMatrix(fepiezas.Row, 7) = "0"
'                  Else
'                    txtValorVenta = txtValorVenta + CCur(fepiezas.TextMatrix(fepiezas.Row, 7))
'                  End If
'           End If
'         Else
'             fepiezas.TextMatrix(fepiezas.Row, 7) = 0
'             MsgBox "El Valor de la Venta no debe ser menor al Valor de la Base "
'         End If
'    End If
'    suma_pieza
'End If
'End Sub
'
'Private Sub suma_pieza()
'    Dim Total As Currency
'    Dim CantFila As Integer
'    Dim i As Integer
'    Total = 0
'    i = 1
'    CantFila = Me.fepiezas.Rows - 1
'    Do While CantFila >= i
'             If IsNull(Me.fepiezas.TextMatrix(i, 7)) Or Me.fepiezas.TextMatrix(i, 7) = "" Then
'                Me.fepiezas.TextMatrix(i, 7) = 0
'             End If
'           Total = Total + CCur(Me.fepiezas.TextMatrix(i, 7))
'           i = i + 1
'      Loop
'      txtValorVenta = Total
'End Sub
'
'Private Sub incio()
'Me.AXCodCta.SetFocusCuenta
'Me.cmdBuscar.SetFocus
'End Sub
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
'        Dim sCuenta As String
'        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
'        If sCuenta <> "" Then
'            AXCodCta.NroCuenta = sCuenta
'            AXCodCta.SetFocusCuenta
'        End If
'    End If
'End Sub
'
'Private Sub Form_Load()
' Dim nRemate As DPigContrato
' Dim nDatosRemate As ADODB.Recordset
' Set nRemate = New DPigContrato
' Set nDatosRemate = nRemate.dObtieneDatosRemate(nRemate.dObtieneMaxRemate() - 1)
' If gdFecSis >= nDatosRemate!dInicio And gdFecSis <= nDatosRemate!dFin Then
'    AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
'    AXCodCta.Age = ""
'    If Not (nDatosRemate.EOF And nDatosRemate.BOF) Then
'        txtremate.Text = nDatosRemate!nRemate
'        txtubica.Text = nDatosRemate!cConsDescripcion
'        lnTipoProceso = nDatosRemate!nTipoProceso
'    End If
'    Set nDatosRemate = Nothing
'    Set nRemate = Nothing
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
'  Else
'    MsgBox "La Fecha no coninciden con las Fechas del Remate", vbExclamation, "Error de Fechas"
'   cmdBuscar.Enabled = False
'   AXCodCta.Enabled = False
'   Me.Caption = "NO SE PUEDE REMATAR FECHAS NO COINCIDEN"
'  End If
'End Sub
'
'Private Sub oplote_Click()
'
'    cmdLote.Visible = True
'    frmPieza.Visible = False
'    fepiezas.Enabled = False
'    emvalvtalote.Enabled = True
'    txtcompra.Enabled = True
'    emvalvtalote.SetFocus
'End Sub
'
'Private Sub oplote_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If emvalvtalote.Enabled And emvalvtalote.Visible Then emvalvtalote.SetFocus
'End If
'End Sub
'
'Private Sub oppieza_Click()
'    'OPCION DE PIEZA - NO VISIBLE LOTE
'    cmdLote.Visible = False
'    frmPieza.Visible = True
'    fepiezas.Enabled = True
'    emvalvtalote.Enabled = False
'End Sub
'Private Sub txtcompra_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        cmdGrabar.SetFocus
'    End If
'End Sub
