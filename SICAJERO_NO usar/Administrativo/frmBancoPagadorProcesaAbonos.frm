VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmBancoPagadorProcesaAbonos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesamiento de Abonos"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10890
   Icon            =   "frmBancoPagadorProcesaAbonos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10695
      Begin VB.CommandButton cmdProcesarAbono 
         Caption         =   "&Procesar Abono"
         Height          =   350
         Left            =   9000
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtMTSoles 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   800
         Width           =   1455
      End
      Begin VB.TextBox txtMTDolares 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox txtTotalReg 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1480
         Width           =   855
      End
      Begin VB.TextBox txtCantSoles 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   800
         Width           =   855
      End
      Begin VB.CommandButton cmdSubirArchivo 
         Caption         =   "..."
         Height          =   300
         Left            =   6120
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtArchivo 
         Height          =   300
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox txtMTDolaresRechazados 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   6720
         Width           =   1455
      End
      Begin VB.TextBox txtMTSolesRechazados 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   6720
         Width           =   1455
      End
      Begin VB.TextBox txtAbonosRechazados 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txtMTDolaresRealizados 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtMTSolesRealizados 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtAbonosRealizados 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   6240
         Width           =   855
      End
      Begin VB.CommandButton CmdCargaArch 
         Caption         =   "&Cargar"
         Height          =   350
         Left            =   9000
         TabIndex        =   13
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox txtCantDolares 
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1140
         Width           =   855
      End
      Begin TabDlg.SSTab SSTAB 
         Height          =   4095
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7223
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Abonos Realizados"
         TabPicture(0)   =   "frmBancoPagadorProcesaAbonos.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "flxAbonosRealizados"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Abonos Rechazados"
         TabPicture(1)   =   "frmBancoPagadorProcesaAbonos.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "flxAbonosRechazados"
         Tab(1).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco1 
            Height          =   3735
            Left            =   -74760
            TabIndex        =   6
            Top             =   600
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6588
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin TarjAdm.FlexEdit flxAbonosRealizados 
            Height          =   3615
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   10245
            _ExtentX        =   22040
            _ExtentY        =   5530
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Cuenta-Moneda-Imp. Neto-Saldo Disp.-Saldo Cont.-Estado Abono-Estado Cta.-Resultado"
            EncabezadosAnchos=   "300-1900-1000-1400-1400-1400-1400-1400-2500-0"
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
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C"
            TextArray0      =   "N°"
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin TarjAdm.FlexEdit flxAbonosRechazados 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   41
            Top             =   360
            Width           =   10245
            _ExtentX        =   22040
            _ExtentY        =   5530
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Cuenta-Moneda-Imp. Neto-Saldo Disp.-Saldo Cont.-Estado Abono-Estado Cta.-Resultado"
            EncabezadosAnchos=   "300-1900-1000-1400-1400-1400-1400-1400-2500-0"
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
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C"
            TextArray0      =   "N°"
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "No procesados :"
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
            Height          =   195
            Left            =   -72120
            TabIndex        =   9
            Top             =   4680
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe :"
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
            Height          =   195
            Left            =   -66720
            TabIndex        =   8
            Top             =   4680
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "F. Proceso:"
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
            Height          =   195
            Left            =   -69480
            TabIndex        =   7
            Top             =   4680
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   7680
         TabIndex        =   30
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "M. T. Soles :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3060
         TabIndex        =   38
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "M. T. Dolares :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         TabIndex        =   36
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Total Registros :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   1560
         Width           =   1425
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Soles :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   465
         TabIndex        =   32
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Archivo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "M. T. Dolares Rechazados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6705
         TabIndex        =   23
         Top             =   6780
         Width           =   2400
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "M. T. Soles Rechazados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         TabIndex        =   22
         Top             =   6780
         Width           =   2220
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Abonos Rechazados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   6780
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "M. T. Dolares Realizados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6705
         TabIndex        =   17
         Top             =   6300
         Width           =   2280
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "M. T. Soles Realizados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2880
         TabIndex        =   16
         Top             =   6300
         Width           =   2100
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Abonos Realizados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   6300
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   12
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Dolares :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   1200
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   10695
      Begin VB.CommandButton cmdGenerarArchivo 
         Caption         =   "&Generar Archivo Confirmación"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9480
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   8280
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   30
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   53
      Filtro          =   "Archivos de Texto (*.pagos)|*.pagos|Archivos de Texto (*.cobros)|*.cobros"
      Altura          =   0
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   9840
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBancoPagadorProcesaAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oBarra As clsProgressBar
Dim lnIdProc As Integer
Dim lsArc As String
Dim NumeroArchivo As Integer
Dim sLinea As String

Private Sub CmdCancelar_Click()
txtArchivo.Text = ""

txtCantSoles.Text = ""
txtCantDolares.Text = ""
txtTotalReg.Text = ""
txtMTSoles.Text = ""
txtMTDolares.Text = ""
txtAbonosRealizados.Text = ""
txtAbonosRechazados.Text = ""
txtMTSolesRealizados.Text = ""
txtMTSolesRechazados.Text = ""
txtMTDolaresRealizados.Text = ""
txtMTDolaresRechazados.Text = ""
Call LimpiaFlex(flxAbonosRealizados)
Call LimpiaFlex(flxAbonosRechazados)
End Sub

Private Sub CmdCargaArch_Click()
        Dim cad As String
        'Cabecera *******
        Dim lnEstadoProc As Integer
        Dim lsEstadoProc As String
        Dim lsCabTipoRegistro_1, lsCabEntidad_2, lsCabFecha_3, lsCapReservado_5 As String
        Dim lnCabNumRegistros_4 As String 'As Integer
        'End Cabecera ***
        'Detalle **********
        Dim lsTipoRegistro_1, lsEntidadFin_3, lsIdentCliente_4, lsReservado_5, lsCodComercio_6, lsCodCnta_8, lsMonedaCnta_9, lsSucursal_10, lsAgencia_11 As String
        Dim lsFechaProceso_13, lsFechaAbono_14, lsMotivo_15, lsNumDeposito_16, lsFechaDeposito_18, lsSucursal_19, lsAgencia_20, lsIdentOrigen_21, lsReservado_22, lsEstadoAbono_23, lsReservado_24 As String
        Dim lnTipoMov_2, lnTipoCuenta_7 As String 'As Integer
        Dim lnImporteNeto_12, lnImporteBruto_17 As Double
        'End Detalle ******
        'Fin ********
        Dim lsFinTipoRegistro_1, lsFinAdquiriente_3, lsFinMonedaSoles_4, lsFinAdquiriente_7, lsFinMonedaDolar_8, lsFinReservado_11 As String
        Dim lnFinNumRegistros_2, lnFinNumOpeSoles_5, lnFinNumOpeDolar_9 As String 'As Integer
        Dim lnFinMontoTotalSoles_6, lnFinMontoTotalDolar_10 As Double
        'End Fin ****
    
    If Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Fecha no valida.", vbInformation, "Aviso"
        mskFecha.SetFocus
        Exit Sub
    End If
    
    Dim lsRuta As String
    lsRuta = Me.txtArchivo.Text
    
    'Valida Archivo y Estado
    Dim cont As Integer
    cont = 0
    
        Open lsRuta For Input As #1
            Do Until EOF(1)
            Input #1, cad
                If cont = 0 Then
                    If Mid(cad, 1, 1) = "A" Then
                            If Format(CDate(Me.mskFecha.Text), "DDMMYYYY") <> Mid(cad, 4, 8) Then
                                    MsgBox "Fecha no valida. La fecha del archivo es : " & Mid(cad, 4, 8), vbInformation, "Aviso"
                                    Close #1
                                    mskFecha.SetFocus
                                    Exit Sub
                            End If
                            
                                Dim rsValEstProc As New ADODB.Recordset
                                Set rsValEstProc = Nothing
                                Set rsValEstProc = DevuelveEstadoProceso(dlgArchivo.FileTitle, Mid(cad, 4, 8))
                                
                                If Not rsValEstProc.EOF Then
                                    If rsValEstProc!nEstadoProceso = 4 Then
                                        lnEstadoProc = 0
                                            Close #1
                                            Exit Do
                                    End If
                                
                                    lnIdProc = rsValEstProc!nIDProc
                                    lnEstadoProc = rsValEstProc!nEstadoProceso
                                    
                                    Select Case lnEstadoProc
                                    Case 1
                                        lsEstadoProc = "Iniciado"
                                    Case 2
                                        lsEstadoProc = "Procesado"
                                    Case 3
                                        lsEstadoProc = "Confirmado"
                                    Case 4
                                        lsEstadoProc = "Extornado"
                                    End Select
                                    
                                    If MsgBox("Estado del proceso " & lsEstadoProc & " ¿Desea Continuar?", vbYesNo + vbInformation, "Aviso") = vbNo Then
                                        Close #1
                                        Exit Sub
                                    End If
                                Else
                                    lnEstadoProc = 0 'No existe registro
                                End If
                    End If
                Else
                 Exit Do
                End If
                cont = cont + 1
            Loop
        Close #1
        

    If lnEstadoProc = 0 Then
            Open lsRuta For Input As #1
                Do Until EOF(1)
                    Input #1, cad
                    
                        If Mid(cad, 1, 1) = "A" Then
                            lsCabTipoRegistro_1 = Mid(cad, 1, 1)
                            lsCabEntidad_2 = Mid(cad, 2, 2)
                            lsCabFecha_3 = Mid(cad, 4, 8)
                            lnCabNumRegistros_4 = Mid(cad, 12, 7)
                            lsCapReservado_5 = Mid(cad, 19, 102)
                            
                           lnIdProc = RegistraCabeceraTrama(dlgArchivo.FileTitle, lsCabTipoRegistro_1, lsCabEntidad_2, lsCabFecha_3, lnCabNumRegistros_4, lsCapReservado_5, gsCodUser)
                            
                        ElseIf Mid(cad, 1, 1) = "Z" Then
                            lsFinTipoRegistro_1 = Mid(cad, 1, 1)
                            lnFinNumRegistros_2 = Mid(cad, 2, 7)
                            lsFinAdquiriente_3 = Mid(cad, 9, 1)
                            lsFinMonedaSoles_4 = Mid(cad, 10, 3)
                            lnFinNumOpeSoles_5 = Mid(cad, 13, 6)
                            lnFinMontoTotalSoles_6 = Format(Mid(Mid(cad, 19, 12), 1, Len(Mid(cad, 19, 12)) - 2), ".") + Right(Mid(cad, 19, 12), 2) 'Mid(cad, 19, 12)
                            lsFinAdquiriente_7 = Mid(cad, 31, 1)
                            lsFinMonedaDolar_8 = Mid(cad, 32, 3)
                            lnFinNumOpeDolar_9 = Mid(cad, 35, 6)
                            lnFinMontoTotalDolar_10 = Format(Mid(Mid(cad, 41, 12), 1, Len(Mid(cad, 41, 12)) - 2), ".") + Right(Mid(cad, 41, 12), 2) 'Mid(cad, 41, 12)
                            lsFinReservado_11 = Mid(cad, 53, 68)
                            
                            Call RegistraFinTrama(lnIdProc, lsFinTipoRegistro_1, lnFinNumRegistros_2, lsFinAdquiriente_3, lsFinMonedaSoles_4, lnFinNumOpeSoles_5, lnFinMontoTotalSoles_6, lsFinAdquiriente_7, _
                                                      lsFinMonedaDolar_8, lnFinNumOpeDolar_9, lnFinMontoTotalDolar_10, lsFinReservado_11)
                            
                        Else
                        
                            If Mid(cad, 1, 1) <> " " Or Mid(cad, 1, 1) <> "V" Or Mid(cad, 1, 1) <> "X" Then
                                cad = " " + cad
                            End If
                        
                            If Mid(cad, 2, 1) = "1" Then
                                lsTipoRegistro_1 = Mid(cad, 1, 1)
                                lnTipoMov_2 = Mid(cad, 2, 1)
                                lsEntidadFin_3 = Mid(cad, 3, 2)
                                lsIdentCliente_4 = Mid(cad, 5, 1)
                                lsReservado_5 = Mid(cad, 6, 7)
                                lsCodComercio_6 = Mid(cad, 13, 9)
                                lnTipoCuenta_7 = Mid(cad, 22, 1)
                                lsCodCnta_8 = Mid(cad, 23, 16)
                                lsMonedaCnta_9 = Mid(cad, 39, 3)
                                lsSucursal_10 = Mid(cad, 42, 3)
                                lsAgencia_11 = Mid(cad, 45, 3)
                                lnImporteNeto_12 = Format(Mid(Mid(cad, 48, 12), 1, Len(Mid(cad, 48, 12)) - 2), ".") + Right(Mid(cad, 48, 12), 2) 'Mid(cad, 48, 12)
                                lsFechaProceso_13 = Mid(cad, 60, 8)
                                lsFechaAbono_14 = Mid(cad, 68, 8)
                                lsMotivo_15 = Mid(cad, 76, 2)
                                lsNumDeposito_16 = Mid(cad, 78, 8)
                                lnImporteBruto_17 = Format(Mid(Mid(cad, 86, 12), 1, Len(Mid(cad, 86, 12)) - 2), ".") + Right(Mid(cad, 86, 12), 2)  'Mid(cad, 86, 12)
                                lsFechaDeposito_18 = Mid(cad, 98, 8)
                                lsSucursal_19 = Mid(cad, 106, 3)
                                lsAgencia_20 = Mid(cad, 109, 3)
                                lsIdentOrigen_21 = Mid(cad, 112, 2)
                                lsReservado_22 = Mid(cad, 114, 1)
                                lsEstadoAbono_23 = Mid(cad, 115, 2)
                                lsReservado_24 = Mid(cad, 117, 4)
        
        
                            Call RegistraDetalleTrama(lnIdProc, lsTipoRegistro_1, lnTipoMov_2, lsEntidadFin_3, lsIdentCliente_4, lsReservado_5, lsCodComercio_6, lnTipoCuenta_7, _
                                                      lsCodCnta_8, lsMonedaCnta_9, lsSucursal_10, lsAgencia_11, lnImporteNeto_12, lsFechaProceso_13, lsFechaAbono_14, _
                                                      lsMotivo_15, lsNumDeposito_16, lnImporteBruto_17, lsFechaDeposito_18, lsSucursal_19, lsAgencia_20, lsIdentOrigen_21, _
                                                    lsReservado_22, lsEstadoAbono_23, lsReservado_24)
                            End If
                        End If
                Loop
            Close #1
               
            ValoresProcesoInicial (lnIdProc)
'            CargaDatosFinProcesoInicial (lnIdProc)
    ElseIf lnEstadoProc = 1 Then 'Estado Iniciado
            ValoresProcesoInicial (lnIdProc)
    ElseIf lnEstadoProc = 2 Then 'Estado Procesado
        ValoresProcesoInicial (lnIdProc)
        CargaDatosFinProcesoInicial (lnIdProc)
    End If
End Sub

Private Sub cmdGenerarArchivo_Click()
    Dim rsCabRes As New ADODB.Recordset
    Dim rsDetRes As ADODB.Recordset
    Dim rsFinRes As New ADODB.Recordset
    
On Error GoTo ArchivoErr
    NumeroArchivo = FreeFile
    
'    lsArc = App.Path & "\Spooler\" & Format(CDate(Me.mskFecha.Text), "DDMMYYYY") & ".txt"
     lsArc = App.Path & "\Spooler\" & "109CONFIRMA" & ".txt"
    
    Open lsArc For Output As #NumeroArchivo
    If LOF(1) > 0 Then
        If MsgBox("Existen Archivos Anteriores en el Directorio, Desea Remplazarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Close #1
            Kill lsArc
            Open lsArc For Append As #1
        Else
            Exit Sub
        End If
    End If
    
    
    ''Respuesta Cabecera

    Set rsCabRes = DevuelveCabeceraResBancoPag(lnIdProc)
    
    If Not rsCabRes.EOF Then
        sLinea = rsCabRes!cCabecera
        Print #NumeroArchivo, sLinea
    End If
    

    ''Detalle

    Set rsDetRes = DevuelveDetalleResBancoPag(lnIdProc)
    
    If rsDetRes.RecordCount > 0 Then
        Do While Not rsDetRes.EOF
            sLinea = rsDetRes!cDetalle
            If Mid(sLinea, 1, 1) <> " " Or Mid(sLinea, 1, 1) <> "V" Or Mid(sLinea, 1, 1) <> "X" Then
                sLinea = " " + sLinea
            End If
            
            Print #NumeroArchivo, sLinea
            rsDetRes.MoveNext
        Loop
    Else
        MsgBox "No se encontraron datos en detalle", vbInformation, "Mensaje del Sistema"
    End If
    
    ''Respuesta Fin

    Set rsFinRes = DevuelveFinResBancoPag(lnIdProc)
    
    If Not rsFinRes.EOF Then
        sLinea = rsFinRes!cFin
        Print #NumeroArchivo, sLinea
    End If

    Close #NumeroArchivo
    MsgBox "Archivo Generado en su Spooler", vbInformation, "Aviso"
    Exit Sub
ArchivoErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    cmdGenerarArchivo.Enabled = False
    Close #NumeroArchivo   ' Cierra el archivo.
End Sub

Private Sub cmdProcesarAbono_Click()
Dim oBancoPag As New DBancoPagador
Dim bTransac As Boolean


    If MsgBox("¿Esta seguro de realizar el proceso de abono?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If

On Error GoTo ErrAbonoBancPag
     Set oBancoPag = New DBancoPagador
     bTransac = False
     Call oBancoPag.dBeginTrans
     bTransac = True

     If (oBancoPag.InsertaAbonoBancoPagadorInicial(lnIdProc, gsCodUser, gAhoBancPagIni, gsCodAge)) Then
         oBancoPag.dCommitTrans
         bTransac = False
         MsgBox "Se ha realizado el proceso de abono a las cuentas con éxito", vbInformation, "Aviso"
         cmdGenerarArchivo.Visible = True
         cmdProcesarAbono.Enabled = False
     Else
         MsgBox "Se ha generado un error un el proceso de abono. Comunicar a TI", vbInformation, "Aviso"
         Call oBancoPag.dRollbackTrans
     End If

    Set oBancoPag = Nothing
    Call CargaDatosFinProcesoInicial(lnIdProc)
    Exit Sub
ErrAbonoBancPag:
    If bTransac Then
        oBancoPag.dRollbackTrans
        Set oBancoPag = Nothing
    End If
    Err.Raise Err.Number, "Error En Proceso", Err.Description
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSubirArchivo_Click()
    txtArchivo.Text = Empty
    
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Texto (*.txt)|*.txt|Todos los Archivo (*.*)|*.*"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
    Else
        MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
        txtArchivo.Text = ""
    End If
End Sub

Public Sub ValoresProcesoInicial(ByVal pnIdProc As Integer)
    Dim rsValProcIni As New ADODB.Recordset
    Set rsValProcIni = DevuelveValoresProcesoInicial(pnIdProc)
    
    If Not rsValProcIni.EOF Then
        txtCantSoles.Text = rsValProcIni!nCantSoles
        txtCantDolares.Text = rsValProcIni!nCantDolares
        txtTotalReg.Text = rsValProcIni!nTotalReg
        txtMTSoles.Text = Format(rsValProcIni!nMontoTotalSoles, "#0.00")
        txtMTDolares.Text = Format(rsValProcIni!nMontoTotalDolares, "#0.00")
    End If
    
    CmdCargaArch.Visible = False
    cmdProcesarAbono.Visible = True
    
    cmdProcesarAbono.Top = 210
    cmdProcesarAbono.Left = 9000
End Sub

Public Sub CargaDatosFinProcesoInicial(ByVal pnIdProc As Integer)
    Dim rsValFinProcIni As New ADODB.Recordset
    Set rsValFinProcIni = DevuelveValoresFinProcesoInicial(pnIdProc)
    
    If Not rsValFinProcIni.EOF Then
        txtAbonosRealizados.Text = rsValFinProcIni!nCantRealizados
        txtAbonosRechazados.Text = rsValFinProcIni!nCantRechazados
        txtMTSolesRealizados.Text = Format(rsValFinProcIni!nMTSolesRealizados, "#0.00")
        txtMTSolesRechazados.Text = Format(rsValFinProcIni!nMTSolesRechazados, "#0.00")
        txtMTDolaresRealizados.Text = Format(rsValFinProcIni!nMTDolaresRealizados, "#0.00")
        txtMTDolaresRechazados.Text = Format(rsValFinProcIni!nMTDolaresRechazados, "#0.00")
    End If
    
    Call CargaDatosRealizados(pnIdProc)
    Call CargaDatosRechazados(pnIdProc)
End Sub


Public Sub CargaDatosRealizados(ByVal pnIdProc As Integer)
    Dim nNumFila As Integer
    Dim rsRegReali As ADODB.Recordset
    Set rsRegReali = DevuelveRegistrosxEstado(pnIdProc, 2) '''Realizados
    
    Call LimpiaFlex(flxAbonosRealizados)
    If rsRegReali.RecordCount > 0 Then
        Me.flxAbonosRealizados.Clear
        Me.flxAbonosRealizados.Rows = 2
        Me.flxAbonosRealizados.FormaCabecera
        
        Do While Not rsRegReali.EOF
            
            If flxAbonosRealizados.TextMatrix(1, 1) <> "" Then
                flxAbonosRealizados.AdicionaFila
            End If
            
            nNumFila = flxAbonosRealizados.Rows - 1
            
            flxAbonosRealizados.TextMatrix(nNumFila, 0) = nNumFila
            flxAbonosRealizados.TextMatrix(nNumFila, 1) = rsRegReali!cCodCnta_8
            flxAbonosRealizados.TextMatrix(nNumFila, 2) = rsRegReali!cMonedaCnta_9
            flxAbonosRealizados.TextMatrix(nNumFila, 3) = Format(rsRegReali!nImporteNeto_12, "##,##0.00")
            flxAbonosRealizados.TextMatrix(nNumFila, 4) = Format(rsRegReali!nSaldoDisp, "##,##0.00")
            flxAbonosRealizados.TextMatrix(nNumFila, 5) = Format(rsRegReali!nSaldoCont, "##,##0.00")
            flxAbonosRealizados.TextMatrix(nNumFila, 6) = rsRegReali!cEstadoAbono_23
            flxAbonosRealizados.TextMatrix(nNumFila, 7) = rsRegReali!cEstadoCta
            flxAbonosRealizados.TextMatrix(nNumFila, 8) = rsRegReali!cDescripcion
            rsRegReali.MoveNext
        Loop
    Else
        MsgBox "No se encontraron datos", vbInformation, "Mensaje del Sistema"
    End If
End Sub

Public Sub CargaDatosRechazados(ByVal pnIdProc As Integer)
    Dim nNumFila As Integer
    Dim rsRegRecha As ADODB.Recordset
    Set rsRegRecha = DevuelveRegistrosxEstado(pnIdProc, 3) '''Rechazados
    
    Call LimpiaFlex(flxAbonosRechazados)
    If rsRegRecha.RecordCount > 0 Then
        Me.flxAbonosRechazados.Clear
        Me.flxAbonosRechazados.Rows = 2
        Me.flxAbonosRechazados.FormaCabecera
        
        Do While Not rsRegRecha.EOF
            
            If flxAbonosRechazados.TextMatrix(1, 1) <> "" Then
                flxAbonosRechazados.AdicionaFila
            End If
            
            nNumFila = flxAbonosRechazados.Rows - 1
            
            flxAbonosRechazados.TextMatrix(nNumFila, 0) = nNumFila
            flxAbonosRechazados.TextMatrix(nNumFila, 1) = rsRegRecha!cCodCnta_8
            flxAbonosRechazados.TextMatrix(nNumFila, 2) = rsRegRecha!cMonedaCnta_9
            flxAbonosRechazados.TextMatrix(nNumFila, 3) = Format(rsRegRecha!nImporteNeto_12, "##,##0.00")
            flxAbonosRechazados.TextMatrix(nNumFila, 4) = Format(rsRegRecha!nSaldoDisp, "##,##0.00")
            flxAbonosRechazados.TextMatrix(nNumFila, 5) = Format(rsRegRecha!nSaldoCont, "##,##0.00")
            flxAbonosRechazados.TextMatrix(nNumFila, 6) = rsRegRecha!cEstadoAbono_23
            flxAbonosRechazados.TextMatrix(nNumFila, 7) = rsRegRecha!cEstadoCta
            flxAbonosRechazados.TextMatrix(nNumFila, 8) = rsRegRecha!cDescripcion
            rsRegRecha.MoveNext
        Loop
    Else
        MsgBox "No se encontraron datos", vbInformation, "Mensaje del Sistema"
    End If
End Sub
