VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenMostrarBilletaje 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Billetaje"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmCajaGenMostrarBilletaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Moneda Extranjera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   4335
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   9
         Left            =   480
         TabIndex        =   27
         Top             =   1275
         Width           =   630
      End
      Begin VB.Label lblTotalME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1910
         TabIndex        =   26
         Top             =   1250
         Width           =   1965
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL MONEDAS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   8
         Left            =   465
         TabIndex        =   25
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label lblTotalMonedaME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1900
         TabIndex        =   24
         Top             =   810
         Width           =   1965
      End
      Begin VB.Label lblTotalBilletesME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1900
         TabIndex        =   23
         Top             =   380
         Width           =   1965
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   7
         Left            =   470
         TabIndex        =   22
         Top             =   460
         Width           =   1410
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   6
         Left            =   360
         Top             =   360
         Width           =   3525
      End
      Begin VB.Label lblTotalBilletes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Index           =   6
         Left            =   1905
         TabIndex        =   20
         Top             =   375
         Width           =   1965
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   6
         Left            =   465
         TabIndex        =   19
         Top             =   450
         Width           =   1410
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   5
         Left            =   360
         Top             =   800
         Width           =   3525
      End
      Begin VB.Label lblTotalBilletes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Index           =   5
         Left            =   1905
         TabIndex        =   18
         Top             =   820
         Width           =   1965
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   5
         Left            =   465
         TabIndex        =   17
         Top             =   890
         Width           =   1410
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   4
         Left            =   360
         Top             =   1230
         Width           =   3525
      End
      Begin VB.Label lblTotalBilletes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Index           =   4
         Left            =   1905
         TabIndex        =   16
         Top             =   1250
         Width           =   1965
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   4
         Left            =   465
         TabIndex        =   15
         Top             =   1320
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Moneda Nacional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4335
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   3
         Left            =   465
         TabIndex        =   13
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label lblTotalMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1905
         TabIndex        =   12
         Top             =   1250
         Width           =   1965
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   3
         Left            =   360
         Top             =   1230
         Width           =   3525
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL MONEDAS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   2
         Left            =   465
         TabIndex        =   11
         Top             =   885
         Width           =   1395
      End
      Begin VB.Label lblTotalMonedaMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1920
         TabIndex        =   10
         Top             =   825
         Width           =   1965
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   2
         Left            =   360
         Top             =   800
         Width           =   3525
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Index           =   0
         Left            =   465
         TabIndex        =   7
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label lblTotalBilletesMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   1905
         TabIndex        =   6
         Top             =   375
         Width           =   1965
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   0
         Left            =   360
         Top             =   360
         Width           =   3525
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   320
         Left            =   2760
         TabIndex        =   3
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblDescUser 
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   840
         TabIndex        =   2
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblCaptionUser 
         AutoSize        =   -1  'True
         Caption         =   "Cajero :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label lblMensaje 
      Caption         =   "Cajero no registrado, se tiene que registrar de nuevo el billetaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOTAL BILLETES :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   150
      Index           =   1
      Left            =   585
      TabIndex        =   9
      Top             =   1770
      Width           =   1410
   End
   Begin VB.Label lblTotalBilletes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   1
      Left            =   2025
      TabIndex        =   8
      Top             =   1695
      Width           =   1965
   End
   Begin VB.Shape ShapeS 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Index           =   1
      Left            =   480
      Top             =   1680
      Width           =   3525
   End
End
Attribute VB_Name = "frmCajaGenMostrarBilletaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DESARROLLADO POR: FRHU - 20160317
Option Explicit
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Public Sub Inicio(ByVal pnMovNroRegEfec As Long)
    Dim oNCaja As New COMNCajaGeneral.NCOMCajero
    Dim rs As New ADODB.Recordset
    
    lblMensaje.Caption = ""
    Set rs = oNCaja.ObtieneRegistroBilletaje(pnMovNroRegEfec)
    If Not (rs.EOF And rs.BOF) Then
        lblDescUser.Caption = rs!cUser
        If lblDescUser.Caption = "" Then
            lblMensaje.Caption = "Cajero no registrado, se tiene que registrar de nuevo el billetaje"
        Else
            lblMensaje.Caption = ""
        End If
        txtfecha.Text = Mid(rs!Fecha, 7, 2) & "/" & Mid(rs!Fecha, 5, 2) & "/" & Left(rs!Fecha, 4)
        lblTotalBilletesMN.Caption = Format(rs!BilleteSoles, "#,##0.00")
        lblTotalMonedaMN.Caption = Format(rs!MonedaSoles, "#,##0.00")
        lblTotalMN.Caption = Format(rs!TotalSoles, "#,##0.00")
        lblTotalBilletesME.Caption = Format(rs!TotalDolares, "#,##0.00")
        lblTotalMonedaME.Caption = "0.00"
        lblTotalME.Caption = Format(rs!TotalDolares, "#,##0.00")
    Else
        Call MsgBox("No se registro el billetaje correctamente.", vbInformation, "AVISO")
    End If
    Me.Show 1
End Sub
