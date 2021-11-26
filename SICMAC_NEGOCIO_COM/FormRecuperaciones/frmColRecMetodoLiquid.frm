VERSION 5.00
Begin VB.Form frmColRecMetodoLiquid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones : Recuperaciones - Método de Liquidación"
   ClientHeight    =   5310
   ClientLeft      =   1230
   ClientTop       =   2940
   ClientWidth     =   7095
   Icon            =   "frmColRecMetodoLiquid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   390
      Left            =   3600
      TabIndex        =   5
      Top             =   4800
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   390
      Left            =   5940
      TabIndex        =   4
      Top             =   4800
      Width           =   1020
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
      Height          =   390
      Left            =   4770
      TabIndex        =   3
      Top             =   4800
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   4665
      Left            =   165
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.CheckBox chkDistribucion 
         Caption         =   "Distribuir Montos a Cobrar"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Frame fraDistribucion 
         Caption         =   "Distribución de Montos"
         Enabled         =   0   'False
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   1830
         Width           =   6495
         Begin SICMACT.EditMoney AXMontos 
            Height          =   285
            Index           =   0
            Left            =   3480
            TabIndex        =   16
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney AXMontos 
            Height          =   285
            Index           =   1
            Left            =   3480
            TabIndex        =   17
            Top             =   1020
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney AXMontos 
            Height          =   285
            Index           =   2
            Left            =   3480
            TabIndex        =   18
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin SICMACT.EditMoney AXMontos 
            Height          =   285
            Index           =   3
            Left            =   3480
            TabIndex        =   19
            Top             =   1620
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblMensaje 
            Caption         =   "Comisión de Abogado : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label lblTotalS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   285
            Left            =   5040
            TabIndex        =   34
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Label lblTotalD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   285
            Left            =   2160
            TabIndex        =   33
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   3
            Left            =   5040
            TabIndex        =   32
            Top             =   1620
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   2
            Left            =   5040
            TabIndex        =   31
            Top             =   1320
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   1
            Left            =   5040
            TabIndex        =   30
            Top             =   1020
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   0
            Left            =   5040
            TabIndex        =   29
            Top             =   720
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   28
            Top             =   1620
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   2
            Left            =   2160
            TabIndex        =   27
            Top             =   1320
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   1
            Left            =   2160
            TabIndex        =   26
            Top             =   1020
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Index           =   0
            Left            =   2160
            TabIndex        =   25
            Top             =   720
            Width           =   1245
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Dist. Cobranza"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Left            =   2400
            TabIndex        =   22
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblComision 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   285
            Left            =   3480
            TabIndex        =   21
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Comi. de Abogado :"
            Height          =   255
            Index           =   9
            Left            =   600
            TabIndex        =   20
            Top             =   1920
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "Total"
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
            Index           =   4
            Left            =   720
            TabIndex        =   15
            Top             =   2280
            Width           =   1245
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
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
            Height          =   285
            Left            =   3480
            TabIndex        =   14
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Gastos :"
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   13
            Top             =   1620
            Width           =   1365
         End
         Begin VB.Label Label7 
            Caption         =   "Int. Compensatorio :"
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   12
            Top             =   1020
            Width           =   1425
         End
         Begin VB.Label Label7 
            Caption         =   "Int. Moratorio :"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   11
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label7 
            Caption         =   "Capital :"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   10
            Top             =   720
            Width           =   1185
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   360
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   960
      End
      Begin SICMACT.ActXCodCta AxCodCta 
         Height          =   465
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   820
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.TextBox txtMetLiq 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Mét.Liquidación"
         Height          =   195
         Index           =   17
         Left            =   1320
         TabIndex        =   36
         Top             =   870
         Width           =   1230
      End
      Begin VB.Label lblMetJudi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   2670
         TabIndex        =   35
         Top             =   840
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label9 
         Caption         =   "Mét. de Liquidación Nuevo :"
         Height          =   195
         Left            =   615
         TabIndex        =   2
         Top             =   1200
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmColRecMetodoLiquid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* RECUPERACIONES - CAMBIO DE METODO LIQUIDACION
'Archivo:  frmColRecMetodoLiquid.frm
'LAYG   :  10/08/2001.
'Resumen:  Nos permite cambiar el metodo de liquidacion de un credito en
'          recuperaciones

Option Explicit
Dim pDias As Integer

'**DAOR 20070421***********************************************************
Dim fnSaldoCap As Currency, fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency, fnSaldoGasto As Currency
Dim fnCapPag As Currency, fnIntCompPag As Currency, fnIntMoratPag As Currency, fnGastoPag As Currency
Dim fnNewSaldoCap As Currency, fnNewSaldoIntComp As Currency, fnNewSaldoIntMorat As Currency, fnNewSaldoGasto As Currency
Dim fnSaldoIntCompGen As Currency, fnNewSaldoIntCompGen As Currency
Dim fnNroUltGastoCta As Integer, fnNroCalend As Integer
Dim fmMatGastos As Variant
Dim nPorcComision As Double
Dim fsFecUltPago As String
Dim fnTasaInt As Double, fnTasaIntMorat As Double
Dim fnTipoCalcIntComp As Integer, fnTipoCalcIntMora As Integer
'0 --> No Calcula
'1 --> Capital
'2 --> Capital + Int Comp
'3 --> Capital + Int comp + Int Morat
Dim fnFormaCalcIntComp As Integer, fnFormaCalcIntMora As Integer
'0 INTERES SIMPLE
'1 INTERES COMPUESTO
Dim fnIntCompGenerado As Currency, fnIntMoraGenerado As Currency
'**************************************************************************
Dim objPista As COMManejador.Pista  '' *** PEAC 20090126



Private Sub limpiar()
    AxCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    'Comentado por DAOR 20070421
    'lblDeuCap = "":    lblDeuCom = ""
    'lblDeuMor = "":    lblDeuGas = ""
    
    '**DAOR 20070421******************************
    lblDeuda(0) = "":    lblDeuda(1) = ""
    lblDeuda(2) = "":    lblDeuda(3) = ""
    
    lblSaldo(0) = "":    lblSaldo(1) = ""
    lblSaldo(2) = "":    lblSaldo(3) = ""
    lblTotalD = ""
    lblTotalS = ""
    '*********************************************
    lblMetJudi = ""
    txtMetLiq.Text = ""
    AXMontos(0).Text = ""
    AXMontos(1).Text = ""
    AXMontos(2).Text = ""
    AXMontos(3).Text = ""
    lblTotal.Caption = ""
    chkDistribucion.value = 0
End Sub
Private Sub HabilitaControles(ByVal pbNuevoMetLiquid As Boolean, ByVal pbGrabar As Boolean, _
        ByVal pbCancelar As Boolean, ByVal pbSalir As Boolean)
    txtMetLiq.Enabled = pbNuevoMetLiquid
    cmdGrabar.Enabled = pbGrabar
    cmdCancelar.Enabled = pbCancelar
    cmdSalir.Enabled = pbSalir
End Sub
Function ValidaDatos()
Dim i, K As Integer
'Valida Que Las Letras Ingresadas en Metodo de liquidacionsolo sean G=gasto, I=Interes, M=Mora, C=Capital
    For i = 1 To Len(txtMetLiq.Text)
        If Mid(txtMetLiq.Text, i, 1) <> "i" And Mid(txtMetLiq.Text, i, 1) <> "G" And Mid(txtMetLiq.Text, i, 1) <> "I" And Mid(txtMetLiq.Text, i, 1) <> "M" And Mid(txtMetLiq.Text, i, 1) <> "C" Then
            MsgBox "Existe Una Letra No Valida En Metodo de Liquidacion", vbInformation, "Aviso"
            If txtMetLiq.Enabled Then
                txtMetLiq.SetFocus
            End If
            ValidaDatos = False
            Exit Function
        End If
    Next i
'Valida Que Las Letras Ingresadas en Metodo de liquidacionsolo sean G=gasto, I=Interes, M=Mora, C=Capital i=intvertical
    For i = 1 To Len(txtMetLiq.Text)
        For K = 1 To Len(txtMetLiq.Text)
            If Mid(txtMetLiq.Text, i, 1) = Mid(txtMetLiq.Text, K, 1) And i <> K Then
                MsgBox "Existe Letra Duplicada en Metodo de Liquidacion", vbInformation, "Aviso"
                If txtMetLiq.Enabled Then
                    txtMetLiq.SetFocus
                End If
                ValidaDatos = False
                Exit Function
            End If
        Next K
    Next i
'Que Halla ingresado texto en met de liquid
    If Len(Trim(txtMetLiq.Text)) = 0 Then
        MsgBox "Ingrese el Metodo de Liquidacion", vbInformation, "Aviso"
            If txtMetLiq.Enabled Then
                txtMetLiq.SetFocus
            End If
        ValidaDatos = False
        Exit Function
    End If
    ValidaDatos = True
End Function

Private Sub BuscaDatos(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValCredito As COMNColocRec.NColRecValida
Dim lsmensaje As String
'**DAOR 20070414*****************************************
Dim loCredRec As COMDColocRec.DCOMColRecCredito
Dim lrCIMG As ADODB.Recordset
Dim loCalcula As COMNColocRec.NCOMColRecCalculos
Dim lnDiasUltTrans As Integer
Dim lrDatGastos As New ADODB.Recordset
'********************************************************
'On Error GoTo ControlError
    'fnTipoCalcIntComp = 0
    fnIntCompGenerado = 0 'JUEZ 20130821
    fnIntMoraGenerado = 0
    fnSaldoGasto = 0
   'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValCredito = New COMNColocRec.NColRecValida
        Set lrValida = loValCredito.nValidaCambioMetodoLiquid(psNroContrato, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    
    '**DAOR 20070421*************************************************
    Set loCredRec = New COMDColocRec.DCOMColRecCredito
        Set lrDatGastos = loCredRec.dObtieneListaGastosxCredito(psNroContrato, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loCredRec = Nothing
    '****************************************************************
    Set loValCredito = Nothing
    
    If lrValida Is Nothing Then ' Hubo un Error
        limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    'Muestra Datos
    'Comentado por DAOR 20070421
    'lblDeuCap.Caption = Format(lrValida!nSaldo, "#0.00")
    'lblDeuCom.Caption = Format(lrValida!nSaldoIntComp, "#0.00")
    'lblDeuMor.Caption = Format(lrValida!nSaldoIntMor, "#0.00")
    'lblDeuGas.Caption = Format(lrValida!nSaldoGasto, "#0.00")
    
    '**DAOR 20070421, Obtener deuda a la fecha***************************************
    ' Asigna Valores a las Variables
    fnSaldoCap = lrValida!nSaldo
    fnSaldoIntComp = lrValida!nSaldoIntComp
    fnSaldoIntMorat = lrValida!nSaldoIntMor
    fnSaldoGasto = lrValida!nSaldoGasto
    fnSaldoIntCompGen = lrValida!nIntCompGen
    fsFecUltPago = CDate(fgFechaHoraGrab(lrValida!cUltimaActualizacion))
    fnTasaInt = IIf(IsNull(lrValida!nTasaInt), 0, lrValida!nTasaInt)
    fnTasaIntMorat = lrValida!nTasaIntMor
    lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
    
    Dim nMontoSaldo As Double
    '*** Carga Gastos en Matriz
    Dim i As Integer
    ReDim fmMatGastos(0)
    ReDim fmMatGastos(lrDatGastos.RecordCount, 11)
    Do While Not lrDatGastos.EOF
        If lrDatGastos!nColocRecGastoEstado = gColRecGastoEstPendiente Then
            fmMatGastos(i, 1) = lrDatGastos!nNroGastoCta
            fmMatGastos(i, 2) = lrDatGastos!nMonto
            fmMatGastos(i, 3) = lrDatGastos!nMontoPagado
            fmMatGastos(i, 4) = lrDatGastos!nColocRecGastoEstado
            fmMatGastos(i, 5) = "N" ' Estado del Gasto
            fmMatGastos(i, 6) = 0 '(fmMatGastos(i, 2) - fmMatGastos(i, 3)) 'avmm 0  ' Monto a Cubrir del Gasto
            fmMatGastos(i, 7) = lrDatGastos!nPrdConceptoCod
            nMontoSaldo = nMontoSaldo + (fmMatGastos(i, 2) - fmMatGastos(i, 3))
            i = i + 1
        End If
        lrDatGastos.MoveNext
    Loop
    fnSaldoGasto = nMontoSaldo
    
    Set loCalcula = New COMNColocRec.NCOMColRecCalculos
    'Calcula el Int Comp Generado
    If fnTipoCalcIntComp = 0 Then ' NoCalcula
        fnIntCompGenerado = 0
    ElseIf fnTipoCalcIntComp = 1 Then ' En base al capital
        If fnFormaCalcIntComp = 1 Then 'INTERES COMPUESTO
            fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
        Else
            'INTERES SIMPLE
            fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap)
        End If
    ElseIf fnTipoCalcIntComp = 2 Then ' En base a capit + int Comp
        If fnFormaCalcIntComp = 1 Then
            fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp)
        Else
            fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp)
        End If
    ElseIf fnTipoCalcIntComp = 3 Then ' En base a capit + int Comp + int Morat
        If fnFormaCalcIntComp = 1 Then
            fnIntCompGenerado = loCalcula.nCalculaIntCompGenerado(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
        Else
            fnIntCompGenerado = loCalcula.nCalculaIntCompGeneradoICA(lnDiasUltTrans, fnTasaInt, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
        End If
    End If
    'Calcula el Int Mora Generado
    If fnTipoCalcIntMora = 0 Then  ' NoCalcula
        fnIntMoraGenerado = 0
    ElseIf fnTipoCalcIntMora = 1 Then ' En base al capital
        If fnFormaCalcIntMora = 1 Then 'INTERES COMPUESTO
            fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
        Else
            'INTERES SIMPLE
            fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap)
        End If
    ElseIf fnTipoCalcIntMora = 2 Then ' En base a capit + int Comp
        If fnFormaCalcIntMora = 1 Then
            fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp)
        Else
            fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp)
        End If
    ElseIf fnTipoCalcIntMora = 3 Then ' En base a capit + int Comp + int Morat
        If fnFormaCalcIntMora = 1 Then
            fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGenerado(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
        Else
            fnIntMoraGenerado = loCalcula.nCalculaIntMoratorioGeneradoICA(lnDiasUltTrans, fnTasaIntMorat, fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat)
        End If
    End If
          
    Set loCalcula = Nothing
    'Agregamos el Int Calculado al Saldo Int Comp
    fnSaldoIntComp = lrValida!nSaldoIntComp + fnIntCompGenerado
    fnSaldoIntMorat = lrValida!nSaldoIntMor + fnIntMoraGenerado

    lblDeuda(0).Caption = Format(fnSaldoCap, "#0.00")
    lblDeuda(1).Caption = Format(fnSaldoIntComp, "#0.00")
    lblDeuda(2).Caption = Format(fnSaldoIntMorat, "#0.00")
    lblDeuda(3).Caption = Format(fnSaldoGasto, "#0.00")
    lblTotalD.Caption = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "#0.00")
    '*****************************************************************
    lblMetJudi.Caption = lrValida!cMetLiquid
    
    '**DAOR 20070421********************************************
    lblMensaje.Caption = "Comisión de Abogado : " & Format(lrValida!nValorCom, "#0.00") & "%"
    nPorcComision = lrValida!nValorCom
    '***********************************************************
    
    Set lrValida = Nothing
    
    'Comentado por DAOR 20070414
    'Me.txtMetLiq.Enabled = True
    ''cmdGrabar.Enabled = True
    'txtMetLiq.SetFocus
    
    '**DAOR 20070414*********************************************
    Set loCredRec = New COMDColocRec.DCOMColRecCredito
    Set lrCIMG = loCredRec.dObtieneDistribucionCIMGCobranza(psNroContrato, gdFecSis)
    Set loCredRec = Nothing
    If Not lrCIMG.EOF Then
        chkDistribucion.value = 1
        fraDistribucion.Enabled = True
        AXMontos(0).Text = Format(lrCIMG!nCapital, "#0.00")
        AXMontos(1).Text = Format(lrCIMG!nIntComp, "#0.00")
        AXMontos(2).Text = Format(lrCIMG!nMora, "#0.00")
        AXMontos(3).Text = Format(lrCIMG!nGasto, "#0.00")
        lblComision.Caption = Format(lrCIMG!nComiAbog, "#0.00")
        
        lblSaldo(0).Caption = CDbl(lblDeuda(0).Caption) - CDbl(AXMontos(0).Text)
        lblSaldo(1).Caption = CDbl(lblDeuda(1).Caption) - CDbl(AXMontos(1).Text)
        lblSaldo(2).Caption = CDbl(lblDeuda(2).Caption) - CDbl(AXMontos(2).Text)
        lblSaldo(3).Caption = CDbl(lblDeuda(3).Caption) - CDbl(AXMontos(3).Text)
        
        lblTotalS.Caption = CDbl(lblSaldo(0).Caption) + CDbl(lblSaldo(1).Caption) + CDbl(lblSaldo(2).Caption) + CDbl(lblSaldo(3).Caption)
        
        lblTotal.Caption = Format(lrCIMG!nCapital + lrCIMG!nIntComp + lrCIMG!nMora + lrCIMG!nGasto + lrCIMG!nComiAbog, "#0.00")
        AXMontos(0).SetFocus
    Else
        Me.txtMetLiq.Enabled = True
        txtMetLiq.SetFocus
    End If
    Set lrCIMG = Nothing
    Me.txtMetLiq.Enabled = True
    '************************************************************
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaDatos(AxCodCta.NroCuenta)
End Sub

'**DAOR 20070414
Private Sub chkDistribucion_Click()
    If chkDistribucion.value = 1 Then
        fraDistribucion.Enabled = True
        cmdGrabar.Enabled = True
    Else
        fraDistribucion.Enabled = False
        cmdGrabar.Enabled = False
    End If
End Sub

Private Sub CmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AxCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AxCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub CmdGrabar_Click()
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito

Dim lsMovNro As String
Dim lsFechaHoraGrab As String


'********* VERIFICAR VISTO AVMM - 13-12-2006 **********************
'Dim loVisto As COMDColocRec.DCOMColRecCredito
'Set loVisto = New COMDColocRec.DCOMColRecCredito
'    '1=Metodo
'    If loVisto.bVerificarVisto(AxCodCta.NroCuenta, 1) = False Then
'        MsgBox "No existe Visto para realizar cambio del Metodo de Liquidación", vbInformation, "Aviso"
'        Exit Sub
'    End If
'Set loVisto = Nothing
'********************************************************************
    
'**Comentado por DAOR 20070414
'    If MsgBox(" Desea Grabar Nuevo Método de Liquidación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'        cmdGrabar.Enabled = False
'
'            'Genera el Mov Nro
'            Set loContFunct = New COMNContabilidad.NCOMContFunciones
'                lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'            Set loContFunct = Nothing
'
'            lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
'            Set loGrabar = New COMNColocRec.NCOMColRecCredito
'                'Grabar cambio de metodo de liquidacion
'                Call loGrabar.nCambiaMetodoLiquidCredRecup(AxCodCta.NroCuenta, txtMetLiq.Text, lsFechaHoraGrab, _
'                     lsMovNro, False)
'            Set loGrabar = Nothing
'
'            Limpiar
'
'            AxCodCta.Enabled = True
'            AxCodCta.SetFocus
'
'    Else
'        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
'    End If
    
    '**DAOR 20070414*****************************************************************************
    
     'MAVM 12112009 ***
    If Len(txtMetLiq.Text) = 4 Then
    '***
    If chkDistribucion.value <> 1 Then
        If MsgBox(" Desea Grabar Nuevo Método de Liquidación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            cmdGrabar.Enabled = False
                
                'Genera el Mov Nro
                Set loContFunct = New COMNContabilidad.NCOMContFunciones
                    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set loContFunct = Nothing
                
                lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
                Set loGrabar = New COMNColocRec.NCOMColRecCredito
                    'Grabar cambio de metodo de liquidacion
                    Call loGrabar.nCambiaMetodoLiquidCredRecup(AxCodCta.NroCuenta, txtMetLiq.Text, lsFechaHoraGrab, _
                         lsMovNro, False)

                    ''*** PEAC 20090126
                    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Grabar Nuevo Método de Liquidación", AxCodCta.NroCuenta, gCodigoCuenta
                                                  
                Set loGrabar = Nothing
        
                limpiar
                
                AxCodCta.Enabled = True
                AxCodCta.SetFocus
                
        Else
            MsgBox " Grabación cancelada ", vbInformation, " Aviso "
        End If
    Else
        If ValidarMontos Then
            If MsgBox(" Desea Grabar La Distribución de los Montos de Cobranza ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                'Genera el Mov Nro
                Set loContFunct = New COMNContabilidad.NCOMContFunciones
                    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set loContFunct = Nothing
                
                Set loGrabar = New COMNColocRec.NCOMColRecCredito
                    Call loGrabar.nRegistraDistribucionCIMGCobranza(AxCodCta.NroCuenta, gdFecSis, CCur(AXMontos(0).Text), CCur(AXMontos(1).Text), CCur(AXMontos(2).Text), CCur(AXMontos(3).Text), CCur(lblComision.Caption), lsMovNro)
                    '*******RECO 2013-07-19*******
                    Call loGrabar.nRegistraAutorizacionPagoJud(AxCodCta.NroCuenta, lsMovNro, gdFecSis, lblMetJudi.Caption, 1, txtMetLiq.Text)
                    '*********END RECO************

                    ''*** PEAC 20090126
                    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Grabar La Distribución de los Montos de Cobranza", AxCodCta.NroCuenta, gCodigoCuenta
                    
                Set loGrabar = Nothing
        
                limpiar
                
                AxCodCta.Enabled = True
                AxCodCta.SetFocus
            End If
        Else
            MsgBox "Los montos distribuidos no deben superar a los montos de la deuda", vbInformation, "Alerta"
        End If
    End If
    
    'MAVM 12112009 ***
    Else
        MsgBox "Debe Ingresar el nuevo Metodo de Liquidacion", vbInformation, "Alerta"
    End If
    '***
    '******************************************************************************************
    Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    limpiar
    Call HabilitaControles(False, False, True, True)
    AxCodCta.Enabled = True
    AxCodCta.SetFocusCuenta
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    limpiar
    AxCodCta.Enabled = True
    CargaParametros 'DAOR 20070421
    'AxCodCta.SetFocusAge
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gRecRegistrarMetodoLiquid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub txtMetLiq_GotFocus()
    Call fEnfoque(txtMetLiq)
End Sub

Private Sub txtMetLiq_KeyPress(KeyAscii As Integer)
    Dim vCadMet As String
    Dim x As Byte
    KeyAscii = fgIntfMayusculas(KeyAscii)
    vCadMet = "GCIM"
    For x = 1 To Len(txtMetLiq)
        vCadMet = Replace(vCadMet, Mid(txtMetLiq, x, 1), "", , , vbTextCompare)
    Next
    If InStr(1, vCadMet, Chr(KeyAscii), vbTextCompare) > 0 Or KeyAscii = 8 Or KeyAscii = 13 Then
        'vCadMet = Replace(vCadMet, Chr(KeyAscii), "", , , vbTextCompare)
        If KeyAscii = 13 Then
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        End If
    Else
        KeyAscii = 0
    End If
End Sub

'**DAOR 20070414
Private Sub AXMontos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValidarMontos Then
            If Index = 3 Then cmdGrabar.SetFocus Else AXMontos(Index + 1).SetFocus
        Else
            MsgBox "Los montos distribuidos no deben superar a los montos de la deuda", vbInformation, "Alerta"
        End If
    End If
End Sub

'**DAOR 20070414
Private Function ValidarMontos() As Boolean
Dim i As Integer, nTotal As Double, nTotalCIM As Double, nTotalS As Double
    ValidarMontos = True
    If CDbl(AXMontos(0).Text) > CDbl(lblDeuda(0).Caption) Then
        AXMontos(0).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(1).Text) > CDbl(lblDeuda(1).Caption) Then
        AXMontos(1).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(2).Text) > CDbl(lblDeuda(2).Caption) Then
        AXMontos(2).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(3).Text) > CDbl(lblDeuda(3).Caption) Then
        AXMontos(1).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    For i = 0 To 3
        'If Val(AXMontos(i).Text) > 0 Then
        lblSaldo(i).Caption = CDbl(lblDeuda(i).Caption) - CDbl(AXMontos(i).Text)
        nTotal = nTotal + CDbl(AXMontos(i).Text)
        nTotalS = nTotalS + CDbl(lblSaldo(i).Caption)
        If i <> 3 Then nTotalCIM = nTotalCIM + CDbl(AXMontos(i).Text)
        'End If
    Next
    lblComision.Caption = Format(nCalculaComisionAbogadoSimple(nPorcComision, nTotalCIM), "#0.00")
    nTotal = nTotal + CDbl(lblComision.Caption)
    lblTotal.Caption = Format(nTotal, "#0.00")
    lblTotalS.Caption = Format(nTotalS, "#0.00")
End Function

'**DAOR 20070421
Private Function nCalculaComisionAbogadoSimple(ByVal pnPorcComision As Double, ByVal pnMontoCIM As Currency) As Currency
    nCalculaComisionAbogadoSimple = pnMontoCIM * (pnPorcComision / 100)
End Function

'**DAOR 20070421
Private Sub CargaParametros()
Dim loParam As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDConstSistema.NCOMConstSistema
    fnTipoCalcIntComp = loParam.LeeConstSistema(151)
    fnTipoCalcIntMora = loParam.LeeConstSistema(152)
    fnFormaCalcIntComp = loParam.LeeConstSistema(202) ' CMACICA
    fnFormaCalcIntMora = loParam.LeeConstSistema(203) ' CMACICA
    
Set loParam = Nothing
End Sub
