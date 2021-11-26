VERSION 5.00
Begin VB.Form frmColRecCancPagoJudicial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperaciones - Cancelación de Créditos con Pago Judicial"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmColRecCancPagoJudicial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5250
      TabIndex        =   49
      Top             =   6720
      Width           =   990
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
      Left            =   6420
      TabIndex        =   48
      Top             =   6720
      Width           =   1020
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
      Height          =   390
      Left            =   4080
      TabIndex        =   47
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cancelación"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   7335
      Begin VB.CheckBox chkAut 
         Caption         =   "Con autorización de Gerencia"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   2610
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtGlosa 
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   2880
         Width           =   6855
      End
      Begin SICMACT.EditMoney AXMontos 
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   38
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   3720
         TabIndex        =   39
         Top             =   780
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   3720
         TabIndex        =   40
         Top             =   1080
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   3720
         TabIndex        =   41
         Top             =   1380
         Width           =   1260
         _ExtentX        =   2223
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
      Begin VB.Label lblTotalDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   285
         Left            =   5280
         TabIndex        =   50
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   45
         Top             =   2625
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Index           =   5
         Left            =   1440
         TabIndex        =   43
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label lblComisionP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   3720
         TabIndex        =   42
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Label lblTotalP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Height          =   285
         Left            =   3720
         TabIndex        =   37
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label lblGastosDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   36
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblIntMoratDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   35
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblIntCompDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   34
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label lblCapitalDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   5280
         TabIndex        =   33
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label lblTotalD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label lblGastosD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   31
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblIntMoratD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   30
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblIntCompD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   29
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
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
         Index           =   10
         Left            =   5400
         TabIndex        =   26
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Capital:"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   25
         Top             =   525
         Width           =   525
      End
      Begin VB.Label lblCapitalD 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Height          =   265
         Left            =   2160
         TabIndex        =   24
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Int. Moratorio:"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   23
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Int. Compensatorio:"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   22
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comision Abogado:"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   20
         Top             =   1725
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   8
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto a Canc."
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
         Index           =   9
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crédito"
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtMetLiq 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   28
         Top             =   2070
         Width           =   1050
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   360
         Left            =   6090
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   820
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblIngRecup 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6240
         TabIndex        =   52
         Top             =   2070
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Recup."
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   51
         Top             =   2115
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid. Nuevo"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2120
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Int."
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   16
         Top             =   1365
         Width           =   630
      End
      Begin VB.Label lblTasaInt 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6240
         TabIndex        =   15
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Demanda"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   14
         Top             =   990
         Width           =   690
      End
      Begin VB.Label lblDemanda 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6240
         TabIndex        =   13
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   990
         Width           =   480
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
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
         Height          =   285
         Left            =   930
         TabIndex        =   11
         Top             =   960
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Condición"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1365
         Width           =   705
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   930
         TabIndex        =   9
         Top             =   1330
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cobranza"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   8
         Top             =   1360
         Width           =   1035
      End
      Begin VB.Label lblTipoCobranza 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3810
         TabIndex        =   7
         Top             =   1335
         Width           =   1425
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Abogado"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1725
         Width           =   645
      End
      Begin VB.Label lblEstudioJur 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   930
         TabIndex        =   5
         Top             =   1710
         Width           =   4305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Met.Liquid."
         Height          =   195
         Index           =   4
         Left            =   5400
         TabIndex        =   4
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label lblMetLiquid 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6240
         TabIndex        =   3
         Top             =   1710
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmColRecCancPagoJudicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'* RECUPERACIONES - CANCELACION DE CREDITOS EN RECUPERACIONES
'** Nombre : frmColRecCancPagoJudicial.frm
'** Descripción : Permite registrar el Monto Final y su Metodo de liquidación que se cobrará en Operaciones para cancelar el credito
'** Creación : JUEZ, 20120419 10:00:00 AM
'********************************************************************
Option Explicit

Dim fnSaldoCap As Currency, fnSaldoIntComp As Currency, fnSaldoIntMorat As Currency, fnSaldoGasto As Currency
Dim fnNewSaldoCap As Currency ',fnNewSaldoIntComp As Currency, fnNewIntCompGen As Currency
Dim fnSaldoIntCompGen As Currency, fnNewSaldoIntCompGen As Currency
Dim fnNroCalen As Integer
Dim fmMatGastos As Variant

Dim fnFechaIngRecup As Date

Dim fnPorcComision As Double
Dim fnIntCompGenerado As Currency
Dim fnIntMoraGenerado As Currency
Dim fnTipoCalcIntComp As Integer, fnTipoCalcIntMora As Integer
Dim fnFormaCalcIntComp As Integer, fnFormaCalcIntMora As Integer
Dim fsFecUltPago As String
Dim fnTasaInt As Double, fnTasaIntMorat As Double
Dim fsCondicion As String, fsDemanda As String
Dim fsCancSKMayorCero As String

Dim objPista As COMManejador.Pista

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaCredito (AXCodCta.NroCuenta)
End Sub


Private Sub AXMontos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValidarMontos Then
            If Index = 3 Then cmdGrabar.SetFocus Else AXMontos(Index + 1).SetFocus
        Else
            MsgBox "Los montos a pagar no deben superar a los montos de la deuda", vbInformation, "Alerta"
        End If
    End If
End Sub

Private Function ValidarMontos() As Boolean
Dim i As Integer, nTotalP As Double, nTotalCIM As Double, nTotalDS As Double
    ValidarMontos = True
    If CDbl(AXMontos(0).Text) > CDbl(lblCapitalD.Caption) Then
        AXMontos(0).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(1).Text) > CDbl(lblIntCompD.Caption) Then
        AXMontos(1).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(2).Text) > CDbl(lblIntMoratD.Caption) Then
        AXMontos(2).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    If CDbl(AXMontos(3).Text) > CDbl(lblGastosD.Caption) Then
        AXMontos(3).SetFocus
        ValidarMontos = False
        Exit Function
    End If
    For i = 0 To 3
        'If Val(AXMontos(i).Text) > 0 Then
        If i = 0 Then
            lblCapitalDS.Caption = Format(CDbl(lblCapitalD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblCapitalDS.Caption)
        ElseIf i = 1 Then
            lblIntCompDS.Caption = Format(CDbl(lblIntCompD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblIntCompDS.Caption)
        ElseIf i = 2 Then
            lblIntMoratDS.Caption = Format(CDbl(lblIntMoratD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblIntMoratDS.Caption)
        Else
            lblGastosDS.Caption = Format(CDbl(lblGastosD.Caption) - CDbl(AXMontos(i).Text), "#0.00")
            nTotalDS = nTotalDS + CDbl(lblGastosDS.Caption)
        End If

        nTotalP = nTotalP + CDbl(AXMontos(i).Text)
        
        If i <> 3 Then nTotalCIM = nTotalCIM + CDbl(AXMontos(i).Text)
        'End If
    Next
    'MAVM 20120717 ***
    'fnPorcComision = nTotalCIM * (fnPorcComision / 100)
    'nTotalP = nTotalP + fnPorcComision
    'lblTotalP.Caption = Format(nTotalP, "#0.00")
    'lblTotalDS.Caption = Format(nTotalDS, "#0.00")
    'lblComisionP.Caption = Format(fnPorcComision, "#0.00")
    nTotalP = nTotalP + (nTotalCIM * (fnPorcComision / 100))
    lblTotalP.Caption = Format(nTotalP, "#0.00")
    lblTotalDS.Caption = Format(nTotalDS, "#0.00")
    lblComisionP.Caption = Format(nTotalCIM * (fnPorcComision / 100), "#0.00")
    '***
End Function

Private Function ValidarUmbralesReglamento(ByRef lsmensaje As String, ByRef PorcPagMin As String) As Boolean
Dim fnCantAñosIngRecup As Integer
Dim fnPorcIntComp As Double
Dim fnIntCompReglam As Double
ValidarUmbralesReglamento = True
'DEBE PAGAR SEGUN REGLAMENTEO
'Jud Hasta 1 año: Capital 100%,Gastos 100%, IntComp 50%, Int.Morat 0%
'Jud A partir 1 Hasta 2 años: Capital 100%,Gastos 100%, IntComp 25%, Int.Morat 0%
'Jud A partir 2 años: Capital 100%,Gastos 100%, IntComp 5%, Int.Morat 0%

fnCantAñosIngRecup = DateDiff("D", fnFechaIngRecup, gdFecSis)
If fnCantAñosIngRecup < 365 Then '365 >> 1 Año
    fnPorcIntComp = 0.5
ElseIf fnCantAñosIngRecup >= 365 And fnCantAñosIngRecup < 730 Then '730 >> 2 años
    fnPorcIntComp = 0.25
ElseIf fnCantAñosIngRecup >= 730 Then
    fnPorcIntComp = 0.05
End If

fnIntCompReglam = Round(CDbl(lblIntCompD.Caption) * fnPorcIntComp, 2)

If CDbl(AXMontos(0).Text) <> CDbl(lblCapitalD.Caption) Then
    AXMontos(0).SetFocus
    lsmensaje = "el Capital"
    PorcPagMin = "100%"
    ValidarUmbralesReglamento = False
    Exit Function
End If
If CDbl(AXMontos(1).Text) < CDbl(fnIntCompReglam) Then
    AXMontos(1).SetFocus
    lsmensaje = "el Interes Comp."
    PorcPagMin = (fnPorcIntComp * 100) & "%"
    ValidarUmbralesReglamento = False
    Exit Function
End If
If CDbl(AXMontos(3).Text) <> CDbl(lblGastosD.Caption) Then
    AXMontos(3).SetFocus
    lsmensaje = "Gastos"
    PorcPagMin = "100%"
    ValidarUmbralesReglamento = False
    Exit Function
End If
End Function

Private Sub CmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If Not loPers Is Nothing Then
        lsPersCod = loPers.sPersCod
        lsPersNombre = loPers.sPersNombre
    Else
        Exit Sub
    End If
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
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    LimpiarCampos
    Call HabilitaControles(False, False, True, False)
    AXCodCta.Enabled = True
    cmdBuscar.Enabled = True
End Sub

Private Sub CmdGrabar_Click()
Dim lsmensaje As String
Dim PorcPagMin As String
If Me.chkAut.value = 1 Then
    If Me.txtGlosa.Text <> "" Then
        RegistrarPagoCancelacion
    Else
        MsgBox "No registró el detalle de la autorización de la Gerencia", vbInformation, "Aviso"
        Me.txtGlosa.SetFocus
    End If
Else
    If ValidarUmbralesReglamento(lsmensaje, PorcPagMin) Then
         RegistrarPagoCancelacion
    Else
        MsgBox "El monto a pagar en " & lsmensaje & " no cumple con el umbral de rebaja definido en el Reglamento de Creditos. (Pago mínimo: " & PorcPagMin & ")", vbInformation, "¡Alerta!"
    End If
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LimpiarCampos
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    Call HabilitaControles(False, False, True, False)
End Sub

Private Sub LimpiarCampos()
    Me.AXCodCta.NroCuenta = ""
    Me.AXCodCta.CMAC = "109"
    Me.AXCodCta.Age = gsCodAge
    Me.lblCliente.Caption = ""
    Me.lblCondicion.Caption = ""
    Me.lblTipoCobranza.Caption = ""
    Me.lblEstudioJur.Caption = ""
    Me.lblDemanda.Caption = ""
    Me.lblTasaInt.Caption = ""
    Me.lblMetLiquid.Caption = ""
    Me.lblIngRecup.Caption = ""
    Me.txtMetLiq.Text = ""
    Me.lblCapitalD.Caption = ""
    Me.lblIntCompD.Caption = ""
    Me.lblIntMoratD.Caption = ""
    Me.lblGastosD.Caption = ""
    Me.lblTotalD.Caption = ""
    Me.AXMontos(0).Text = ""
    Me.AXMontos(1).Text = ""
    Me.AXMontos(2).Text = ""
    Me.AXMontos(3).Text = ""
    Me.lblComisionP.Caption = ""
    Me.lblTotalP.Caption = ""
    Me.lblCapitalDS.Caption = ""
    Me.lblIntCompDS.Caption = ""
    Me.lblIntMoratDS.Caption = ""
    Me.lblGastosDS.Caption = ""
    Me.lblTotalDS.Caption = ""
    Me.chkAut.value = 0
    Me.txtGlosa.Text = ""
    '** Juez 20120716 ****************
    If gsCodCargo = "002017" Then 'Solo Jefe Recuperaciones
        Me.chkAut.Visible = True
    Else
        Me.chkAut.Visible = False
    End If
    '** End Juez *********************
End Sub

Private Sub BuscaCredito(ByVal psCtaCod As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim loValCredito As COMNColocRec.NColRecValida
Dim lrDatCredito As ADODB.Recordset
Dim lrDatGastos As New ADODB.Recordset
Dim loValCred As COMDColocRec.DCOMColRecCredito
Dim loCredRec As COMDColocRec.DCOMColRecCredito
Dim lrCIMG As ADODB.Recordset
Dim lnDiasUltTrans As Integer
Dim lnIntCompGenCal As Double
'On Error GoTo ControlError

Dim lsmensaje As String

Call CargaParametros

    'valida Contrato
    Set loValCredito = New COMNColocRec.NColRecValida
    Set lrValida = loValCredito.nValidaCambioMetodoLiquid(psCtaCod, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If

    'Carga Datos
    Set loValCred = New COMDColocRec.DCOMColRecCredito
        Set lrDatCredito = loValCred.dObtieneDatosCancelaCredRecup(psCtaCod, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Then   ' Hubo un Error
        MsgBox "No se Encontro el Credito o No se ha realizado registro de expediente", vbInformation, "Aviso"
        LimpiarCampos
        Set lrDatCredito = Nothing
        Exit Sub
    End If
    
    Set loCredRec = New COMDColocRec.DCOMColRecCredito
        Set lrDatGastos = loCredRec.dObtieneListaGastosxCredito(psCtaCod, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loCredRec = Nothing
    
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
        ' Asigna Valores a las Variables
'        fnSaldoCap = lrDatCredito!nSaldo
'        fnSaldoIntComp = lrDatCredito!nSaldoIntComp
'        fnSaldoIntMorat = lrDatCredito!nSaldoIntMor
'        fnSaldoGasto = lrDatCredito!nSaldoGasto
'        fnTasaInt = lrDatCredito!nTasaInteres
'        fnIntCompGenerado = lrDatCredito!nIntCompGen
'        fnNroCalen = lrDatCredito!nNroCalen
        
        fnSaldoCap = lrValida!nSaldo
        fnSaldoIntComp = lrValida!nSaldoIntComp
        fnSaldoIntMorat = lrValida!nSaldoIntMor
'        fnSaldoGasto = lrValida!nSaldoGasto
        fnSaldoGasto = nMontoSaldo
        fnSaldoIntCompGen = lrDatCredito!nIntCompGen
        fsFecUltPago = CDate(fgFechaHoraGrab(lrValida!cUltimaActualizacion))
        fnTasaInt = IIf(IsNull(lrValida!nTasaInt), 0, lrValida!nTasaInt)
        fnTasaIntMorat = lrValida!nTasaIntMor
        lnDiasUltTrans = CDate(Format(gdFecSis, "dd/mm/yyyy")) - CDate(Format(fsFecUltPago, "dd/mm/yyyy"))
        
        'Muestra Datos
        Me.lblCliente.Caption = PstaNombre(Trim(lrDatCredito!cPersNombre))
        Me.lblCondicion.Caption = fgCondicionColRecupDesc(lrDatCredito!nPrdEstado)
        Me.lblTipoCobranza.Caption = IIf(lrDatCredito!nTipCJ = gColRecTipCobJudicial, "Judicial", "ExtraJudicial")
        Me.lblEstudioJur.Caption = lrDatCredito!cPersNombreAbog
        
        Me.lblDemanda.Caption = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        fsDemanda = IIf(lrDatCredito!nDemanda = gColRecDemandaSi, "S", "N")
        Me.lblMetLiquid.Caption = lrDatCredito!cMetLiquid
        fnPorcComision = lrDatCredito!nValorCom
        Me.lblComisionP.Caption = Format(lrDatCredito!nValorCom, "#,##0.00")
        fsCondicion = IIf(lrDatCredito!nPrdEstado = gColocEstRecVigJud, "J", "A")
        Me.lblTasaInt.Caption = Format(lrDatCredito!nTasaInteres, "#,##0.00")
        fnFechaIngRecup = lrDatCredito!dIngRecup
        Me.lblIngRecup = fnFechaIngRecup
        
    Set lrDatCredito = Nothing
    
    Call HabilitaControles(True, True, True, True)
    
    'Obtiene los montos grabados de la misma fecha
    Set loCredRec = New COMDColocRec.DCOMColRecCredito
    Set lrCIMG = loCredRec.dObtieneDistribucionCIMGCobranza(psCtaCod, gdFecSis)
    Set loCredRec = Nothing
    If Not lrCIMG.EOF Then
        AXMontos(0).Text = Format(lrCIMG!nCapital, "#0.00")
        AXMontos(1).Text = Format(lrCIMG!nIntComp, "#0.00")
        AXMontos(2).Text = Format(lrCIMG!nMora, "#0.00")
        AXMontos(3).Text = Format(lrCIMG!nGasto, "#0.00")
        lblComisionP.Caption = Format(lrCIMG!nComiAbog, "#0.00")
        
        lblTotalP.Caption = Format(lrCIMG!nCapital + lrCIMG!nIntComp + lrCIMG!nMora + lrCIMG!nGasto + lrCIMG!nComiAbog, "#0.00")
        AXMontos(0).SetFocus
    Else
        lblCapitalDS.Caption = Format(0, "#0.00")
        lblIntCompDS.Caption = Format(0, "#0.00")
        lblIntMoratDS.Caption = Format(0, "#0.00")
        lblGastosDS.Caption = Format(0, "#0.00")
        
        lblTotalDS.Caption = Format(0, "#0.00")
        Me.txtMetLiq.Enabled = True
        txtMetLiq.SetFocus
    End If
    Set lrCIMG = Nothing
        
    Dim loCalcula As COMNColocRec.NCOMColRecCalculos
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
    
     Me.lblCapitalD.Caption = Format(fnSaldoCap, "#,##0.00")
     Me.lblIntCompD.Caption = Format(fnSaldoIntComp, "#,##0.00")
     Me.lblIntMoratD.Caption = Format(fnSaldoIntMorat, "#,##0.00")
     Me.lblGastosD.Caption = Format(fnSaldoGasto, "#,##0.00")
     Me.lblTotalD.Caption = Format(fnSaldoCap + fnSaldoIntComp + fnSaldoIntMorat + fnSaldoGasto, "#,##0.00")
    
        lblCapitalDS.Caption = CDbl(lblCapitalD.Caption) - CDbl(AXMontos(0).Text)
        lblIntCompDS.Caption = CDbl(lblIntCompD.Caption) - CDbl(AXMontos(1).Text)
        lblIntMoratDS.Caption = CDbl(lblIntMoratD.Caption) - CDbl(AXMontos(2).Text)
        lblGastosDS.Caption = CDbl(lblGastosD.Caption) - CDbl(AXMontos(3).Text)
        
        lblTotalDS.Caption = CDbl(lblCapitalDS.Caption) + CDbl(lblIntCompDS.Caption) + CDbl(lblIntMoratDS.Caption) + CDbl(lblGastosDS.Caption)
    
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    AXCodCta.Enabled = False
    cmdBuscar.Enabled = False
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub HabilitaControles(ByVal pbCmdGrabar As Boolean, ByVal pbCmdCancelar As Boolean, _
            ByVal pbCmdSalir As Boolean, ByVal pbTxt As Boolean)
    cmdGrabar.Enabled = pbCmdGrabar
    cmdCancelar.Enabled = pbCmdCancelar
    cmdSalir.Enabled = pbCmdSalir
    Me.AXMontos(0).Enabled = pbTxt
    Me.AXMontos(1).Enabled = pbTxt
    Me.AXMontos(2).Enabled = pbTxt
    Me.AXMontos(3).Enabled = pbTxt
    txtMetLiq.Enabled = pbTxt
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


Public Sub RegistrarPagoCancelacion()
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito
Set objPista = New COMManejador.Pista
'En el anterior frm el nro ope era 130600    Cancelacion de Credito
gsOpeCod = "191040" 'Colocaciones: Recuperaciones - Cancelaciones de Creditos

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
'Dim lsmensaje As String

If Len(txtMetLiq.Text) = 4 Then

    If MsgBox(" Desea Grabar la cancelación del crédito? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
        
            Call loGrabar.nRegistraPagoCancelacionCredJudicial(AXCodCta.NroCuenta, gdFecSis, CCur(AXMontos(0).Text), CCur(AXMontos(1).Text), CCur(AXMontos(2).Text), CCur(AXMontos(3).Text), _
                                                            CCur(lblComisionP.Caption), txtMetLiq.Text, txtGlosa.Text, fnNroCalen, lsMovNro, gsOpeCod)
            
            '*******RECO 2013-07-19*******
           Call loGrabar.nRegistraAutorizacionPagoJud(AXCodCta.NroCuenta, lsMovNro, gdFecSis, lblMetLiquid.Caption, 1, txtMetLiq.Text)
           '*********END RECO************

            MsgBox "El crédito " & AXCodCta.NroCuenta & " está listo para ser cancelado en Operaciones", vbInformation, "Aviso"
            
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Grabar Cancelación Credito con Pago Judicial", AXCodCta.NroCuenta, gCodigoCuenta
                                          
        Set loGrabar = Nothing
        
        Call HabilitaControles(False, False, True, False)
        LimpiarCampos
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        cmdBuscar.Enabled = True
            
    Else
        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
    End If
    
Else
    MsgBox "Debe Ingresar correctamente el Metodo de Liquidacion para cancelar", vbExclamation, "Alerta"
    Me.txtMetLiq.SetFocus
End If
Exit Sub

ControlError:
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub CargaParametros()
Dim loParam As COMDConstSistema.NCOMConstSistema
Set loParam = New COMDConstSistema.NCOMConstSistema
    fnTipoCalcIntComp = loParam.LeeConstSistema(151)
    fnTipoCalcIntMora = loParam.LeeConstSistema(152)
    fnFormaCalcIntComp = loParam.LeeConstSistema(202) ' CMACICA
    fnFormaCalcIntMora = loParam.LeeConstSistema(203) ' CMACICA
    
Set loParam = Nothing
End Sub
