VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepConsolAho 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimiento de Dep. según monto y Nro."
   ClientHeight    =   5100
   ClientLeft      =   1665
   ClientTop       =   2970
   ClientWidth     =   8970
   Icon            =   "FRMREPCONSOLAHO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRep 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   5055
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   5415
      Begin MSComctlLib.TreeView TV 
         Height          =   4695
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   8281
         _Version        =   393217
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   3
         ImageList       =   "imglPic"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraTC 
      Caption         =   "Tipo de Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1455
      Left            =   5640
      TabIndex        =   17
      Top             =   0
      Width           =   3255
      Begin VB.TextBox TxtTipCamFijAnt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   325
         Left            =   360
         TabIndex        =   1
         Text            =   "0"
         Top             =   420
         Width           =   1125
      End
      Begin VB.TextBox txtTipCamCompraSBS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   325
         Left            =   360
         TabIndex        =   3
         Text            =   "0"
         Top             =   1040
         Width           =   1125
      End
      Begin VB.TextBox txtTipCamFij 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   325
         Left            =   1920
         TabIndex        =   2
         Text            =   "0"
         Top             =   420
         Width           =   1005
      End
      Begin VB.TextBox txtTipCamVentaSBS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   325
         Left            =   1920
         TabIndex        =   4
         Text            =   "0"
         Top             =   1040
         Width           =   1005
      End
      Begin VB.Label lblTipcambiAnt 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Fijo.Ant"
         Height          =   195
         Left            =   510
         TabIndex        =   21
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Comp.SBS"
         Height          =   195
         Left            =   390
         TabIndex        =   20
         Top             =   800
         Width           =   1065
      End
      Begin VB.Label lblTipcamFij 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Fijo.Act"
         Height          =   195
         Left            =   2010
         TabIndex        =   19
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T.C.Venta.SBS"
         Height          =   195
         Left            =   1882
         TabIndex        =   18
         Top             =   800
         Width           =   1080
      End
   End
   Begin VB.Frame fraAge 
      Caption         =   "Agencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2295
      Left            =   5640
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
      Begin VB.ListBox lstAge 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdConcolidar 
      Caption         =   "&Consolida"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   4650
      Width           =   1000
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   150
      Left            =   8520
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   265
      _Version        =   393217
      TextRTF         =   $"FRMREPCONSOLAHO.frx":030A
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
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   4650
      Width           =   1000
   End
   Begin MSComctlLib.ImageList imglPic 
      Left            =   720
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMREPCONSOLAHO.frx":038B
            Key             =   "Cerrado"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMREPCONSOLAHO.frx":07DD
            Key             =   "Abierto"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMREPCONSOLAHO.frx":0C2F
            Key             =   "Hoja"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   4650
      Width           =   1000
   End
   Begin VB.Frame fraMes 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   825
      Left            =   5640
      TabIndex        =   15
      Top             =   3735
      Visible         =   0   'False
      Width           =   3255
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   825
      Left            =   5640
      TabIndex        =   12
      Top             =   3735
      Visible         =   0   'False
      Width           =   3135
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Fec. Inicio:"
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
         TabIndex        =   14
         Top             =   200
         Width           =   975
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fec. Fin:"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   200
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmRepConsolAho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnAnchoNormal As Integer
Dim lnAnchoMax As Integer
Dim lsPerson As String
Dim dbConexionConsol As New ADODB.Connection
Dim oConConsol As DConecta
Dim gcCentralCom As String

Dim lsImpre As String
'VARIABLES PARA EL CALCULO
'* Proceso de Calculo previo a la Generacion del Reporte
Dim lnNumCueVigAhoSol As Long
Dim lnNumCueVigPFSol As Long
Dim lnMonCueVigAhoSol As Currency
Dim lnMonCueVigPFSol As Currency

Dim lnNumCueNueAhoSol As Long
Dim lnNumCueNuePFSol As Long
Dim lnMonCueNueAhoSol As Currency
Dim lnMonCueNuePFSol As Currency

Dim lnNumCueAbonoAhoSol As Long
Dim lnMonCueAbonoAhoSol As Currency
Dim lnNumCueAbonoPFSol As Long
Dim lnMonCueAbonoPFSol As Currency

Dim lnMonCueIntCapAhoSol As Currency
Dim lnMonCueIntCapPFSol As Currency

Dim lnNumCueCanAhoSol As Long
Dim lnMonCueCanAhoSol As Currency
Dim lnNumCueCanPFSol As Long
Dim lnMonCueCanPFSol As Currency

Dim lnNumCueRetAhoSol As Long
Dim lnMonCueRetAhoSol As Currency
Dim lnNumCueRetPFSol As Long
Dim lnMonCueRetPFSol As Currency

Dim lnTotCueVigAhoSol As Long
Dim lnTotCueVigPFSol As Long
Dim lnMonTotCueVigAhoSol As Currency
Dim lnMonTotCueVigPFSol As Currency

'0000000
Dim lnNumCueVigAhoDol As Long
Dim lnNumCueVigPFDol As Long
Dim lnMonCueVigAhoDol As Currency
Dim lnMonCueVigPFDol As Currency

Dim lnNumCueNueAhoDol As Long
Dim lnNumCueNuePFDol As Long
Dim lnMonCueNueAhoDol As Currency
Dim lnMonCueNuePFDol As Currency

Dim lnNumCueAbonoAhoDol As Long
Dim lnMonCueAbonoAhoDol As Currency
Dim lnNumCueAbonoPFDol As Long
Dim lnMonCueAbonoPFDol As Currency

Dim lnMonCueIntCapAhoDol As Currency
Dim lnMonCueIntCapPFDol As Currency

Dim lnNumCueCanAhoDol As Long
Dim lnMonCueCanAhoDol As Currency
Dim lnNumCueCanPFDol As Long
Dim lnMonCueCanPFDol As Currency

Dim lnNumCueRetAhoDol As Long
Dim lnMonCueRetAhoDol As Currency
Dim lnNumCueRetPFDol As Long
Dim lnMonCueRetPFDol As Currency

Dim lnTotCueVigAhoDol As Long
Dim lnTotCueVigPFDol As Long
Dim lnMonTotCueVigAhoDol As Currency
Dim lnMonTotCueVigPFDol As Currency

Dim lnMonTotCueVigAhoDolVVV As Currency
Dim lnMonTotCueVigPFDolVVV As Currency

Dim lsRTFRfm As String
Dim lsReporte As String

Dim lsTipAho As String

'FIN VARIABLES PARA EL CALCULO
Private Function GetMovMonMun(pbConsol As Boolean, Optional psCodAge As String = "", Optional psNomAge As String = "") As String
    Dim lsNombreAge As String

    Dim lsFechaIni As String
    Dim lsFechaFin As String
    Dim lsFechaAnt As String

    Dim ldFechaIni As Date
    Dim ldFechaFin As Date
    Dim ldFechaAnt As Date

    Dim lsMes As String
    Dim lsDia As String
    Dim lsAño As String

    Dim lsNumAho As String * 11
    Dim lsMonAho As String * 17
    Dim lsNumPF As String * 11
    Dim lsMonPF As String * 17
    Dim lsNumT As String * 11
    Dim lsMonT As String * 17

    Dim lnMax As Integer

    Dim oGen As DGeneral
    Set oGen = New DGeneral

    lnMax = 5

    If pbConsol Then
        lsNombreAge = "CONSOLIDADO"
    Else
        lsNombreAge = psNomAge
    End If

    lsMes = Mid(Trim(txtFecha.Text), 4, 2)
    lsDia = "01"
    lsAño = Mid(Trim(txtFecha.Text), 7, 4)
    ldFechaIni = CDate(lsDia & "/" & lsMes & "/" & lsAño)
    ldFechaFin = DateAdd("m", 1, ldFechaIni)

    ldFechaAnt = DateAdd("d", -1, ldFechaIni)

    lsFechaIni = Format(ldFechaIni, "mm/dd/yyyy")
    lsFechaFin = Format(ldFechaFin, "mm/dd/yyyy")
    lsFechaAnt = Format(ldFechaAnt, "mm/dd/yyyy")

    'Calculo Soles
    lnNumCueVigAhoSol = CuentasVigentesAho(lsFechaAnt, lsFechaIni, "1", lnMonCueVigAhoSol, pbConsol, psCodAge)
    If psCodAge = "11201" Or pbConsol Then
        lnMonCueVigAhoSol = lnMonCueVigAhoSol + CCur(oGen.GetParametro(2000, 1001)) '  ReadVarSis("AHO", "nMonCtaAddMDMN"))
        lnNumCueVigAhoSol = lnNumCueVigAhoSol + CCur(oGen.GetParametro(2000, 1002))
    End If
    lnNumCueVigPFSol = CuentasVigentesPF(lsFechaAnt, lsFechaIni, "1", lnMonCueVigPFSol, pbConsol, psCodAge)

    lnNumCueNueAhoSol = CuentasApeAho(lsFechaIni, lsFechaFin, "1", lnMonCueNueAhoSol, pbConsol, psCodAge)
    lnNumCueNuePFSol = CuentasApePF(lsFechaIni, lsFechaFin, "1", lnMonCueNuePFSol, pbConsol, psCodAge)

    lnNumCueAbonoAhoSol = CuentasAbonoAho(lsFechaIni, lsFechaFin, "1", lnMonCueAbonoAhoSol, pbConsol, psCodAge)
    lnNumCueAbonoPFSol = CuentasAbonoPF(lsFechaIni, lsFechaFin, "1", lnMonCueAbonoPFSol, pbConsol, psCodAge)

    lnMonCueIntCapAhoSol = CuentasIntCapAho(lsFechaIni, lsFechaFin, "1", pbConsol, psCodAge)
    lnMonCueIntCapPFSol = CuentasIntCapPF(lsFechaIni, lsFechaFin, "1", pbConsol, psCodAge)

    lnNumCueCanAhoSol = CuentasCanceladaAho(lsFechaIni, lsFechaFin, "1", lnMonCueCanAhoSol, pbConsol, psCodAge) * -1
    lnNumCueCanPFSol = CuentasCanceladaPF(lsFechaIni, lsFechaFin, "1", lnMonCueCanPFSol, pbConsol, psCodAge) * -1

    lnNumCueRetAhoSol = CuentasRetAho(lsFechaIni, lsFechaFin, "1", lnMonCueRetAhoSol, pbConsol, psCodAge)
    lnNumCueRetPFSol = CuentasRetPF(lsFechaIni, lsFechaFin, "1", lnMonCueRetPFSol, pbConsol, psCodAge)

    lnTotCueVigAhoSol = lnNumCueVigAhoSol + lnNumCueNueAhoSol + lnNumCueCanAhoSol
    lnTotCueVigPFSol = lnNumCueVigPFSol + lnNumCueNuePFSol + lnNumCueCanPFSol
    lnMonTotCueVigAhoSol = lnMonCueVigAhoSol + lnMonCueNueAhoSol + lnMonCueAbonoAhoSol + lnMonCueIntCapAhoSol + lnMonCueCanAhoSol + lnMonCueRetAhoSol
    lnMonTotCueVigPFSol = lnMonCueVigPFSol + lnMonCueNuePFSol + lnMonCueAbonoPFSol + lnMonCueIntCapPFSol + lnMonCueCanPFSol + lnMonCueRetPFSol

    'Calculo Dolares
    lnNumCueVigAhoDol = CuentasVigentesAho(lsFechaAnt, lsFechaIni, "2", lnMonCueVigAhoDol, pbConsol, psCodAge)
    lnNumCueVigPFDol = CuentasVigentesPF(lsFechaAnt, lsFechaIni, "2", lnMonCueVigPFDol, pbConsol, psCodAge)

    lnNumCueNueAhoDol = CuentasApeAho(lsFechaIni, lsFechaFin, "2", lnMonCueNueAhoDol, pbConsol, psCodAge)
    lnNumCueNuePFDol = CuentasApePF(lsFechaIni, lsFechaFin, "2", lnMonCueNuePFDol, pbConsol, psCodAge)

    lnNumCueAbonoAhoDol = CuentasAbonoAho(lsFechaIni, lsFechaFin, "2", lnMonCueAbonoAhoDol, pbConsol, psCodAge)
    lnNumCueAbonoPFDol = CuentasAbonoPF(lsFechaIni, lsFechaFin, "2", lnMonCueAbonoPFDol, pbConsol, psCodAge)

    lnMonCueIntCapAhoDol = CuentasIntCapAho(lsFechaIni, lsFechaFin, "2", pbConsol, psCodAge)
    lnMonCueIntCapPFDol = CuentasIntCapPF(lsFechaIni, lsFechaFin, "2", pbConsol, psCodAge)

    lnNumCueCanAhoDol = CuentasCanceladaAho(lsFechaIni, lsFechaFin, "2", lnMonCueCanAhoDol, pbConsol, psCodAge) * -1
    lnNumCueCanPFDol = CuentasCanceladaPF(lsFechaIni, lsFechaFin, "2", lnMonCueCanPFDol, pbConsol, psCodAge) * -1

    lnNumCueRetAhoDol = CuentasRetAho(lsFechaIni, lsFechaFin, "2", lnMonCueRetAhoDol, pbConsol, psCodAge)
    lnNumCueRetPFDol = CuentasRetPF(lsFechaIni, lsFechaFin, "2", lnMonCueRetPFDol, pbConsol, psCodAge)

    'multiplicacion
    lnMonCueVigAhoDol = Round(lnMonCueVigAhoDol, 2)
    lnMonCueVigPFDol = Round(lnMonCueVigPFDol, 2)

    lnMonCueNueAhoDol = Round(lnMonCueNueAhoDol, 2)
    lnMonCueNuePFDol = Round(lnMonCueNuePFDol, 2)

    lnMonCueAbonoAhoDol = Round(lnMonCueAbonoAhoDol, 2)

    lnMonCueIntCapAhoDol = Round(lnMonCueIntCapAhoDol, 2)
    lnMonCueIntCapPFDol = Round(lnMonCueIntCapPFDol, 2)

    lnMonCueCanAhoDol = Round(lnMonCueCanAhoDol, 2)
    lnMonCueCanPFDol = Round(lnMonCueCanPFDol, 2)

    lnMonCueRetAhoDol = Round(lnMonCueRetAhoDol, 2)
    lnMonCueRetPFDol = Round(lnMonCueRetPFDol, 2)

    lnTotCueVigAhoDol = lnNumCueVigAhoDol + lnNumCueNueAhoDol + lnNumCueCanAhoDol
    lnTotCueVigPFDol = lnNumCueVigPFDol + lnNumCueNuePFDol + lnNumCueCanPFDol

    'Montos totales
    lnMonTotCueVigAhoDolVVV = lnMonCueVigAhoDol + lnMonCueNueAhoDol + lnMonCueAbonoAhoDol + lnMonCueIntCapAhoDol + lnMonCueCanAhoDol + lnMonCueRetAhoDol
    lnMonTotCueVigPFDolVVV = lnMonCueVigPFDol + lnMonCueNuePFDol + lnMonCueIntCapPFDol + lnMonCueCanPFDol + lnMonCueRetPFDol + lnMonCueAbonoPFDol

    lnMonTotCueVigAhoDolVVV = lnMonTotCueVigAhoDolVVV * Me!txtTipCamFij
    lnMonTotCueVigPFDolVVV = lnMonTotCueVigPFDolVVV * Me!txtTipCamFij

    'Montos totales
    'lnMonTotCueVigAhoDol = lnMonCueVigAhoDol + lnMonCueNueAhoDol + lnMonCueAbonoAhoDol + lnMonCueIntCapAhoDol + lnMonCueCanAhoDol + lnMonCueRetAhoDol
    'lnMonTotCueVigPFDol = lnMonCueVigPFDol + lnMonCueNuePFDol + lnMonCueIntCapPFDol + lnMonCueCanPFDol + lnMonCueRetPFDol

    '00000
    lnMonCueVigAhoDol = Round(lnMonCueVigAhoDol * (Me!TxtTipCamFijAnt), 2)
    lnMonCueVigPFDol = Round(lnMonCueVigPFDol * Me!TxtTipCamFijAnt, 2)

    lnMonCueNueAhoDol = Round(lnMonCueNueAhoDol * Me!txtTipCamFij, 2)
    lnMonCueNuePFDol = Round(lnMonCueNuePFDol * Me!txtTipCamFij, 2)

    lnMonCueAbonoAhoDol = Round(lnMonCueAbonoAhoDol * Me!txtTipCamFij, 2)
    lnMonCueAbonoPFDol = Round(lnMonCueAbonoPFDol * Me!txtTipCamFij, 2)

    lnMonCueIntCapAhoDol = Round(lnMonCueIntCapAhoDol * Me!txtTipCamFij, 2)
    lnMonCueIntCapPFDol = Round(lnMonCueIntCapPFDol * Me!txtTipCamFij, 2)

    lnMonCueCanAhoDol = Round(lnMonCueCanAhoDol * Me!txtTipCamFij, 2)
    lnMonCueCanPFDol = Round(lnMonCueCanPFDol * Me!txtTipCamFij, 2)

    lnMonCueRetAhoDol = Round(lnMonCueRetAhoDol * Me!txtTipCamFij, 2)
    lnMonCueRetPFDol = Round(lnMonCueRetPFDol * Me!txtTipCamFij, 2)

    lnMonTotCueVigAhoDol = lnMonCueVigAhoDol + lnMonCueNueAhoDol + lnMonCueAbonoAhoDol + lnMonCueIntCapAhoDol + lnMonCueCanAhoDol + lnMonCueRetAhoDol
    lnMonTotCueVigPFDol = lnMonCueVigPFDol + lnMonCueNuePFDol + lnMonCueAbonoPFDol + lnMonCueIntCapPFDol + lnMonCueCanPFDol + lnMonCueRetPFDol

    lnMonCueRetAhoDol = lnMonCueRetAhoDol + (lnMonTotCueVigAhoDolVVV - lnMonTotCueVigAhoDol)
    lnMonCueCanPFDol = lnMonCueCanPFDol + (lnMonTotCueVigPFDolVVV - lnMonTotCueVigPFDol)


    lnMonTotCueVigAhoDol = lnMonTotCueVigAhoDolVVV
    lnMonTotCueVigPFDol = lnMonTotCueVigPFDolVVV

    'MsgBox "hola"
    '* Proceso de Generacion del Reporte

    Dim lsCadena As String
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String

    Dim lsMonto As String
    Dim lsNroCta As String

    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF

    lsNroCta = "  NRO.CTAS "
    lsMonto = "      MONTO      "

    lsCadena = ""
    lsCadena = lsCadena & Space(lnMax) & lsNegritaOn & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "SUPERINTENDENCIA DE BANCA Y SEGUROS" & Space(79) & "CODIGO: " & gsCodCMAC & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & CentrarCadena(gsNomCmac, 126) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & CentrarCadena("MOVIMIENTO DE LOS DEPOSITOS SEGUN MONTO Y NUMERO DE CUENTAS - " & lsNombreAge, 126) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & CentrarCadena("AL: " & Format$(ldFechaIni, "dd/mm/yyyy"), 126) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & CentrarCadena("(En Miles Nuevos Soles)", 126) & lsNegritaOff & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "+" & String(124, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & " CODIGO " & "¦" & Space(25) & "¦" & CentrarCadena("AHORROS", 28) & "¦" & CentrarCadena("PLAZOFIJO", 28) & "¦" & CentrarCadena("TOTAL", 28) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & " S.B.S. " & "¦" & CentrarCadena("DETALLE", 24) & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & lsNroCta & "¦" & lsMonto & "¦" & lsNroCta & "¦" & lsMonto & "¦" & lsNroCta & "¦" & lsMonto & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "+" & String(8, "-") & "+" & String(25, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & CentrarCadena("MONEDA NACIONAL", 24) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueVigAhoSol, "#,##0"))) & Format(lnNumCueVigAhoSol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueVigAhoSol, "#,##0.00"))) & Format(lnMonCueVigAhoSol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueVigPFSol, "#,##0"))) & Format(lnNumCueVigPFSol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueVigPFSol, "#,##0.00"))) & Format(lnMonCueVigPFSol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueVigAhoSol + lnNumCueVigPFSol, "#,##0"))) & Format(lnNumCueVigAhoSol + lnNumCueVigPFSol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueVigAhoSol + lnMonCueVigPFSol, "#,##0.00"))) & Format(lnMonCueVigAhoSol + lnMonCueVigPFSol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10001 " & "¦" & " Saldo Mes Anterior" & Space(25 - Len(" Saldo Mes Anterior")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueNueAhoSol, "#,##0"))) & Format(lnNumCueNueAhoSol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueNueAhoSol, "#,##0.00"))) & Format(lnMonCueNueAhoSol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueNuePFSol, "#,##0"))) & Format(lnNumCueNuePFSol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueNuePFSol, "#,##0.00"))) & Format(lnMonCueNuePFSol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueNueAhoSol + lnNumCueNuePFSol, "#,##0"))) & Format(lnNumCueNueAhoSol + lnNumCueNuePFSol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueNueAhoSol + lnMonCueNuePFSol, "#,##0.00"))) & Format(lnMonCueNueAhoSol + lnMonCueNuePFSol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10002 " & "¦" & " Mas: Cuentas Nuevas" & Space(25 - Len(" Mas: Cuentas Nuevas")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueAbonoAhoSol, "#,##0"))) & Format(lnNumCueAbonoAhoSol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueAbonoAhoSol, "#,##0.00"))) & Format(lnMonCueAbonoAhoSol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueAbonoPFSol, "#,##0"))) & Format(lnNumCueAbonoPFSol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueAbonoPFSol, "#,##0.00"))) & Format(lnMonCueAbonoPFSol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueAbonoAhoSol + lnNumCueAbonoPFSol, "#,##0"))) & Format(lnNumCueAbonoAhoSol + lnNumCueAbonoPFSol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueAbonoAhoSol + lnMonCueAbonoPFSol, "#,##0.00"))) & Format(lnMonCueAbonoAhoSol + lnMonCueAbonoPFSol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10003 " & "¦" & "      Abonos" & Space(25 - Len("      Abonos")) & "¦" & Space(11) & "¦" & lsMonAho & "¦" & Space(11) & "¦" & lsMonPF & "¦" & Space(11) & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsMonAho = Space(17 - Len(Format(lnMonCueIntCapAhoSol, "#,##0.00"))) & Format(lnMonCueIntCapAhoSol, "#,##0.00")
    lsMonPF = Space(17 - Len(Format(lnMonCueIntCapPFSol, "#,##0.00"))) & Format(lnMonCueIntCapPFSol, "#,##0.00")
    lsMonT = Space(17 - Len(Format((lnMonCueIntCapAhoSol + lnMonCueIntCapPFSol), "#,##0.00"))) & Format((lnMonCueIntCapAhoSol + lnMonCueIntCapPFSol), "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10004 " & "¦" & "      Intereses Capit." & Space(25 - Len("      Intereses Capit.")) & "¦" & Space(11) & "¦" & lsMonAho & "¦" & Space(11) & "¦" & lsMonPF & "¦" & Space(11) & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueCanAhoSol, "#,##0"))) & Format(lnNumCueCanAhoSol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueCanAhoSol, "#,##0.00"))) & Format(lnMonCueCanAhoSol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueCanPFSol, "#,##0"))) & Format(lnNumCueCanPFSol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueCanPFSol, "#,##0.00"))) & Format(lnMonCueCanPFSol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueCanAhoSol + lnNumCueCanPFSol, "#,##0"))) & Format(lnNumCueCanAhoSol + lnNumCueCanPFSol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueCanAhoSol + lnMonCueCanPFSol, "#,##0.00"))) & Format(lnMonCueCanAhoSol + lnMonCueCanPFSol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10020 " & "¦" & " Menos: Ctas Canceladas." & Space(25 - Len(" Menos: Ctas Canceladas.")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueRetAhoSol, "#,##0"))) & Format(lnNumCueRetAhoSol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueRetAhoSol, "#,##0.00"))) & Format(lnMonCueRetAhoSol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueRetPFSol, "#,##0"))) & Format(lnNumCueRetPFSol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueRetPFSol, "#,##0.00"))) & Format(lnMonCueRetPFSol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueRetAhoSol + lnNumCueRetPFSol, "#,##0"))) & Format(lnNumCueRetAhoSol + lnNumCueRetPFSol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueRetAhoSol + lnMonCueRetPFSol, "#,##0.00"))) & Format(lnMonCueRetAhoSol + lnMonCueRetPFSol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10021 " & "¦" & "        Retiros y Cargos." & Space(25 - Len("        Retiros y Cargos.")) & "¦" & Space(11) & "¦" & lsMonAho & "¦" & Space(11) & "¦" & lsMonPF & "¦" & Space(11) & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnTotCueVigAhoSol, "#,##0"))) & Format(lnTotCueVigAhoSol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonTotCueVigAhoSol, "#,##0.00"))) & Format(lnMonTotCueVigAhoSol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnTotCueVigPFSol, "#,##0"))) & Format(lnTotCueVigPFSol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonTotCueVigPFSol, "#,##0.00"))) & Format(lnMonTotCueVigPFSol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnTotCueVigAhoSol + lnTotCueVigPFSol, "#,##0"))) & Format(lnTotCueVigAhoSol + lnTotCueVigPFSol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonTotCueVigAhoSol + lnMonTotCueVigPFSol, "#,##0.00"))) & Format(lnMonTotCueVigAhoSol + lnMonTotCueVigPFSol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "+" & String(8, "-") & "+" & String(25, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & "  10040 " & "¦" & " Saldo a fin de Mes" & Space(25 - Len(" Saldo a fin de Mes")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "+" & String(8, "-") & "+" & String(25, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & oImpresora.gPrnSaltoLinea

    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea

    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & CentrarCadena("MONEDA EXTRANJERA", 24) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea

'00000000000000000000000000
    lsNumAho = Space(11 - Len(Format(lnNumCueVigAhoDol, "#,##0"))) & Format(lnNumCueVigAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueVigAhoDol, "#,##0.00"))) & Format(lnMonCueVigAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueVigPFDol, "#,##0"))) & Format(lnNumCueVigPFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueVigPFDol, "#,##0.00"))) & Format(lnMonCueVigPFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueVigAhoDol + lnNumCueVigPFDol, "#,##0"))) & Format(lnNumCueVigAhoDol + lnNumCueVigPFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueVigAhoDol + lnMonCueVigPFDol, "#,##0.00"))) & Format(lnMonCueVigAhoDol + lnMonCueVigPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20001 " & "¦" & " Saldo Mes Anterior" & Space(25 - Len(" Saldo Mes Anterior")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueNueAhoDol, "#,##0"))) & Format(lnNumCueNueAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueNueAhoDol, "#,##0.00"))) & Format(lnMonCueNueAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueNuePFDol, "#,##0"))) & Format(lnNumCueNuePFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueNuePFDol, "#,##0.00"))) & Format(lnMonCueNuePFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueNueAhoDol + lnNumCueNuePFDol, "#,##0"))) & Format(lnNumCueNueAhoDol + lnNumCueNuePFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueNueAhoDol + lnMonCueNuePFDol, "#,##0.00"))) & Format(lnMonCueNueAhoDol + lnMonCueNuePFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20002 " & "¦" & " Mas: Cuentas Nuevas" & Space(25 - Len(" Mas: Cuentas Nuevas")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueAbonoAhoDol, "#,##0"))) & Format(lnNumCueAbonoAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueAbonoAhoDol, "#,##0.00"))) & Format(lnMonCueAbonoAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueAbonoPFDol, "#,##0"))) & Format(lnNumCueAbonoPFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueAbonoPFDol, "#,##0.00"))) & Format(lnMonCueAbonoPFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueAbonoAhoDol + lnNumCueAbonoPFDol, "#,##0"))) & Format(lnNumCueAbonoAhoDol + lnNumCueAbonoPFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueAbonoAhoDol + lnMonCueAbonoPFDol, "#,##0.00"))) & Format(lnMonCueAbonoAhoDol + lnMonCueAbonoPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20003 " & "¦" & "      Abonos" & Space(25 - Len("      Abonos")) & "¦" & Space(11) & "¦" & lsMonAho & "¦" & Space(11) & "¦" & lsMonPF & "¦" & Space(11) & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsMonAho = Space(17 - Len(Format(lnMonCueIntCapAhoDol, "#,##0.00"))) & Format(lnMonCueIntCapAhoDol, "#,##0.00")
    lsMonPF = Space(17 - Len(Format(lnMonCueIntCapPFDol, "#,##0.00"))) & Format(lnMonCueIntCapPFDol, "#,##0.00")
    lsMonT = Space(17 - Len(Format(lnMonCueIntCapAhoDol + lnMonCueIntCapPFDol, "#,##0.00"))) & Format(lnMonCueIntCapAhoDol + lnMonCueIntCapPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20004 " & "¦" & "      Intereses Capit." & Space(25 - Len("      Intereses Capit.")) & "¦" & Space(11) & "¦" & lsMonAho & "¦" & Space(11) & "¦" & lsMonPF & "¦" & Space(11) & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueCanAhoDol, "#,##0"))) & Format(lnNumCueCanAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueCanAhoDol, "#,##0.00"))) & Format(lnMonCueCanAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueCanPFDol, "#,##0"))) & Format(lnNumCueCanPFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueCanPFDol, "#,##0.00"))) & Format(lnMonCueCanPFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueCanAhoDol + lnNumCueCanPFDol, "#,##0"))) & Format(lnNumCueCanAhoDol + lnNumCueCanPFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueCanAhoDol + lnMonCueCanPFDol, "#,##0.00"))) & Format(lnMonCueCanAhoDol + lnMonCueCanPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20020 " & "¦" & " Menos: Ctas Canceladas." & Space(25 - Len(" Menos: Ctas Canceladas.")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnNumCueRetAhoDol, "#,##0"))) & Format(lnNumCueRetAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonCueRetAhoDol, "#,##0.00"))) & Format(lnMonCueRetAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnNumCueRetPFDol, "#,##0"))) & Format(lnNumCueRetPFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonCueRetPFDol, "#,##0.00"))) & Format(lnMonCueRetPFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnNumCueRetAhoDol + lnNumCueRetPFDol, "#,##0"))) & Format(lnNumCueRetAhoDol + lnNumCueRetPFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonCueRetAhoDol + lnMonCueRetPFDol, "#,##0.00"))) & Format(lnMonCueRetAhoDol + lnMonCueRetPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20021 " & "¦" & "        Retiros y Cargos." & Space(25 - Len("        Retiros y Cargos.")) & "¦" & Space(11) & "¦" & lsMonAho & "¦" & Space(11) & "¦" & lsMonPF & "¦" & Space(11) & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea

    lsNumAho = Space(11 - Len(Format(lnTotCueVigAhoDol, "#,##0"))) & Format(lnTotCueVigAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonTotCueVigAhoDol, "#,##0.00"))) & Format(lnMonTotCueVigAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnTotCueVigPFDol, "#,##0"))) & Format(lnTotCueVigPFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonTotCueVigPFDol, "#,##0.00"))) & Format(lnMonTotCueVigPFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnTotCueVigAhoDol + lnTotCueVigPFDol, "#,##0"))) & Format(lnTotCueVigAhoDol + lnTotCueVigPFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonTotCueVigAhoDol + lnMonTotCueVigPFDol, "#,##0.00"))) & Format(lnMonTotCueVigAhoDol + lnMonTotCueVigPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "+" & String(8, "-") & "+" & String(25, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & "  20040 " & "¦" & " Saldo a fin de Mes" & Space(25 - Len(" Saldo a fin de Mes")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "+" & String(8, "-") & "+" & String(25, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & String(11, "-") & "+" & String(17, "-") & "+" & oImpresora.gPrnSaltoLinea
'00000000000000000000000000

    lsNumAho = Space(11 - Len(Format(lnTotCueVigAhoSol + lnTotCueVigAhoDol, "#,##0"))) & Format(lnTotCueVigAhoSol + lnTotCueVigAhoDol, "#,##0")
    lsMonAho = Space(17 - Len(Format(lnMonTotCueVigAhoSol + lnMonTotCueVigAhoDol, "#,##0.00"))) & Format(lnMonTotCueVigAhoSol + lnMonTotCueVigAhoDol, "#,##0.00")
    lsNumPF = Space(11 - Len(Format(lnTotCueVigPFSol + lnTotCueVigPFDol, "#,##0"))) & Format(lnTotCueVigPFSol + lnTotCueVigPFDol, "#,##0")
    lsMonPF = Space(17 - Len(Format(lnMonTotCueVigPFSol + lnMonTotCueVigPFDol, "#,##0.00"))) & Format(lnMonTotCueVigPFSol + lnMonTotCueVigPFDol, "#,##0.00")
    lsNumT = Space(11 - Len(Format(lnTotCueVigAhoSol + lnTotCueVigPFSol + lnTotCueVigAhoDol + lnTotCueVigPFDol, "#,##0"))) & Format(lnTotCueVigAhoSol + lnTotCueVigPFSol + lnTotCueVigAhoDol + lnTotCueVigPFDol, "#,##0")
    lsMonT = Space(17 - Len(Format(lnMonTotCueVigAhoSol + lnMonTotCueVigPFSol + lnMonTotCueVigAhoDol + lnMonTotCueVigPFDol, "#,##0.00"))) & Format(lnMonTotCueVigAhoSol + lnMonTotCueVigPFSol + lnMonTotCueVigAhoDol + lnMonTotCueVigPFDol, "#,##0.00")

    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & "  01000 " & "¦" & " Saldo a fin de Mes" & Space(25 - Len(" Saldo a fin de Mes")) & "¦" & lsNumAho & "¦" & lsMonAho & "¦" & lsNumPF & "¦" & lsMonPF & "¦" & lsNumT & "¦" & lsMonT & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & "  02000 " & "¦" & " Promedio de Saldo Diario" & Space(25 - Len(" Promedio de Saldo Diario")) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "¦" & Space(8) & "¦" & Space(25) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & Space(11) & "¦" & Space(17) & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "+" & String(124, "-") & "+" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "Periodicidad Mensual" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "Total Vista fin de Mes   : Se anotará el saldo de la cuenta Nro. 2101" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "Total Ahorro fin de Mes  : Se anotará el saldo del rubro 23" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "Total Plazo fin de Mes   : Se anotará el saldo del rubro 24" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMax) & "Total Monto              : Se anotará la suma de los saldo de la cuenta Nro.2101 y rubros 23 y 24" & oImpresora.gPrnSaltoLinea

    GetMovMonMun = lsCadena

End Function

Private Sub cmdConcolidar_Click()
    R.Text = ""

    If TV.SelectedItem.Key = "M" Or TV.SelectedItem.Key = "D" Then
        MsgBox "Debe Elegir un Reporte.", vbInformation, "Aviso"
        TV.SetFocus
        Exit Sub
    ElseIf TV.SelectedItem.Key = "D1" Then
        R.Text = GeneraD1Consolida
    End If

    If R.Text <> "" Then
       EnviaPrevio lsImpre, Me.TV.SelectedItem.Text, True, 66
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    R.Text = ""

    If TV.SelectedItem.Key = "M" Or TV.SelectedItem.Key = "D" Then
        MsgBox "Debe Elegir un Reporte.", vbInformation, "Aviso"
        TV.SetFocus
        Exit Sub
    ElseIf TV.SelectedItem.Key = "D1" Then
        R.Text = GeneraD1
    ElseIf TV.SelectedItem.Key = "M1" Then
        R.Text = GeneraM1
    ElseIf TV.SelectedItem.Key = "M2" Then
        If ValidaM Then Exit Sub
        R.Text = GeneraM2
    ElseIf TV.SelectedItem.Key = "M3" Then
        Dim ldFecha As Date
        ldFecha = CDate(txtFecha.Text)
        R.Text = GetGanCompVenta(ldFecha, TxtTipCamFijAnt, txtTipCamCompraSBS, txtTipCamVentaSBS)
    ElseIf TV.SelectedItem.Key = "M4" Then
        R.Text = DepositosEscalaMontos(CDbl(txtTipCamFij))
    ElseIf TV.SelectedItem.Key = "M5" Then
        Dim nFSD As Double
        nFSD = oGen.GetParametro(2000, 2019)
        R.Text = GeneraFSD(nFSD, CDbl(txtTipCamFij))
        oPrevio.Show R.Text, Me.TV.SelectedItem.Text, True, 66
        If MsgBox("Desea generar archivo en Excel del FSD??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            GeneraFSDExcel nFSD, CDbl(txtTipCamFij)
        End If
        Exit Sub
    ElseIf TV.SelectedItem.Key = "M6" Then
'        Dim lsFecha As String, sCodAge As String
'        Dim i As Integer
'        ldFecha = CDate("01/" & Right(cmbMes, 2) & "/" & mskAño)
'        ldFecha = DateAdd("m", 1, ldFecha)
'        ldFecha = DateAdd("d", -1, ldFecha)
'        lsFecha = Format$(ldFecha, "dd/mm/yyyy")
'        For i = 0 To lstAge.ListCount - 1
'            If lstAge.Selected(i) Then
'                If Right(lstAge.List(i), 5) = "ONSOL" Then
'                    sCodAge = "__"
'                Else
'                    sCodAge = Right(lstAge.List(i), 2)
'                End If
'                'Ahorros Soles
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "1", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "1", "1", sCodAge, lsFecha, 1)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "2", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "2", "1", sCodAge, lsFecha, 1)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "3", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "3", "1", sCodAge, lsFecha, 1)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "6", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "6", "1", sCodAge, lsFecha, 1)
'                'Ahorros Dolares
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "1", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "1", "2", sCodAge, lsFecha, 1)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "2", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "2", "2", sCodAge, lsFecha, 1)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "3", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "3", "2", sCodAge, lsFecha, 1)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "6", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProAC, "6", "2", sCodAge, lsFecha, 1)
'
'                'Plazo Fijo Soles
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "1", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "2", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "3", "1", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "6", "1", sCodAge, lsFecha, 0)
'                'Plazo Fijo Soles
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "1", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "2", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "3", "2", sCodAge, lsFecha, 0)
'                R.Text = R.Text + GenSdosFM(gsCodProPF, "6", "2", sCodAge, lsFecha, 0)
'
'                'CTS Soles
'                R.Text = R.Text + GenSdosFM(gsCodProCTS, "1", "1", sCodAge, lsFecha, 0, "")
'                R.Text = R.Text + GenSdosFM(gsCodProCTS, "1", "1", sCodAge, lsFecha, 0, gsInstCmac)
'                'CTS Dolares
'                R.Text = R.Text + GenSdosFM(gsCodProCTS, "1", "2", sCodAge, lsFecha, 0, "")
'                R.Text = R.Text + GenSdosFM(gsCodProCTS, "1", "2", sCodAge, lsFecha, 0, gsInstCmac)
'            End If
'        Next i

    ElseIf TV.SelectedItem.Key = "M7" Then
        GeneraFSDExcelFuncionarios nFSD, CDbl(txtTipCamFij)
    End If

    If R.Text <> "" Then oPrevio.Show R.Text, Me.TV.SelectedItem.Text, True, 66, gImpresora
End Sub

Private Function GeneraM1() As String
    Dim I As Integer
    Dim lsCadena As String
    Dim lsCadCon As String

    If ValidaM Then Exit Function

    oConConsol.AbreConexion 'Remota "07", , , "03"

    lsCadena = ""
    For I = 0 To lstAge.ListCount - 1
        If lstAge.Selected(I) Then
            If Right(lstAge.List(I), 5) = "ONSOL" Then
                lsCadena = lsCadena & GetMovMonMun(True)
            Else
                lsCadena = lsCadena & GetMovMonMun(False, Right(lstAge.List(I), 5)) & oImpresora.gPrnSaltoPagina
            End If
        End If
    Next I

    oConConsol.CierraConexion

    GeneraM1 = lsCadena
End Function

Private Function GeneraM2() As String
    Dim I As Integer
    Dim lsCadena As String

    Dim dFecha As Date

    Dim lsMes As String
    Dim lsDia As String
    Dim lsAño As String

    If ValidaM Then Exit Function

    dFecha = CDate(txtFecha.Text)
    lsCadena = SaldoCaptaciones(dFecha)

    GeneraM2 = lsCadena
End Function

Private Function GeneraD1() As String
    Dim I As Integer
    Dim lsCadena As String
    Dim J As Integer

    If ValidaD Then Exit Function

    lsCadena = ""
    For I = 0 To lstAge.ListCount - 1
        If lstAge.Selected(I) Then
            If Right(lstAge.List(I), 5) = "ONSOL" Then
                For J = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    If J Mod 5 = 0 And J <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    lsCadena = lsCadena & GetMovCompraVentaME(DateAdd("d", J, CDate(Me.mskFecIni)), DateAdd("d", J, CDate(Me.mskFecIni)))
                Next J
            Else
                For J = 0 To DateDiff("d", CDate(Me.mskFecIni), CDate(Me.mskFecFin))
                    If J Mod 5 = 0 And J <> 0 Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    lsCadena = lsCadena & GetMovCompraVentaME(DateAdd("d", J, CDate(Me.mskFecIni)), DateAdd("d", J, CDate(Me.mskFecIni)), Right(lstAge.List(I), 5))
                Next J

                'lsCadena = lsCadena & GetMovCompraVentaME(CDate(Me.mskFecIni), CDate(Me.mskFecFin), Right(lstAge.List(i), 5)) & oImpresora.gPrnSaltoPagina
            End If
        End If
    Next I

    GeneraD1 = lsCadena
End Function

Private Function GeneraD1Consolida() As String
    Dim I As Integer
    Dim lsCadena As String
    Dim J As Integer
    Dim lbBan As Boolean

    If ValidaD Then Exit Function

    lsCadena = ""
    For I = 0 To lstAge.ListCount - 1
        If lstAge.Selected(I) Then
            If Right(lstAge.List(I), 5) = "ONSOL" Then
                lsCadena = lsCadena & GetMovCompraVentaMEConsolida(CDate(Me.mskFecIni), DateAdd("d", 1, CDate(mskFecFin)))
            Else
                lsCadena = lsCadena & GetMovCompraVentaMEConsolida(CDate(Me.mskFecIni), DateAdd("d", 1, CDate(mskFecFin)), Right(lstAge.List(I), 5))
            End If
        End If
    Next I

    GeneraD1Consolida = lsCadena
End Function


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sqlAge As String
    Dim rsAge As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oCambio As nTipoCambio
    Set oCambio = New nTipoCambio
    Set oConConsol = New DConecta

    oCon.AbreConexion
    gcCentralCom = "DBCOMUNES.."

    Set rsAge = New ADODB.Recordset
    rsAge.CursorLocation = adUseClient

    sqlAge = "Select cAgeDescripcion cNomtab, '" & gsCodCMAC & "' + cAgeCod cValor From Agencias"
    Set rsAge = oCon.CargaRecordSet(sqlAge)

    Me.cmdConcolidar.Enabled = False
    lstAge.Clear

    If Not RSVacio(rsAge) Then
        While Not rsAge.EOF
            lstAge.AddItem Trim(rsAge!cNomtab) & Space(500) & Trim(rsAge!cValor)
            rsAge.MoveNext
        Wend
        lstAge.AddItem "Consolidado" & Space(500) & "CONSOL"
    End If

    rsAge.Close
    Set rsAge = Nothing

    TV.Nodes.Clear
    TV.Nodes.Add , tvwPrevious, "D", "Reportes Diarios", "Cerrado"
    TV.Nodes.Add "D", tvwChild, "D1", "Reporte Compra Venta de Moneda Extranjera", "Hoja"
    TV.Nodes.Add , tvwPrevious, "M", "Reportes Mensuales", "Cerrado"
    TV.Nodes.Add "M", tvwChild, "M1", "Mov. Depósitos Según Monto y Núm. Cuentas", "Hoja"
    TV.Nodes.Add "M", tvwChild, "M2", "Situacion de Captaciones", "Hoja"
    TV.Nodes.Add "M", tvwChild, "M3", "Movimiento de Compra Venta (Asiento 021)", "Hoja"
    TV.Nodes.Add "M", tvwChild, "M4", "Depósitos Según Escala de Montos", "Hoja"
    TV.Nodes.Add "M", tvwChild, "M5", "Listado para el FSD", "Hoja"
    TV.Nodes.Add "M", tvwChild, "M6", "Listado General de Cuentas", "Hoja"
    TV.Nodes.Add "M", tvwChild, "M7", "FSD - Cuentas Funcionarios", "Hoja"

    txtTipCamFij = Format(gnTipCambio, "#,##0.0000")
    TxtTipCamFijAnt = Format(oCambio.EmiteTipoCambio(DateAdd("m", -1, gdFecSis), TCFijoMes), "#,##0.0000")
    txtTipCamCompraSBS = "0.0000"
    txtTipCamVentaSBS = "0.0000"

    txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    mskFecIni.Text = Format(gdFecSis, "dd/mm/yyyy")
    mskFecFin.Text = Format(gdFecSis, "dd/mm/yyyy")

    lnAnchoNormal = 6240
    lnAnchoMax = 11700
    
    CentraForm Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub lstAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TV.SelectedItem.Key = "D1" Then
            mskFecIni.SetFocus
        ElseIf TV.SelectedItem.Key = "D" Or TV.SelectedItem.Key = "M" Then
            cmdImprimir.SetFocus
        Else
            txtFecha.SetFocus
        End If
    End If
End Sub

Private Sub txtFecha_GotFocus()
    txtFecha.SelStart = 0
    txtFecha.SelLength = Len(txtFecha.Text)
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdImprimir.SetFocus
    End If
End Sub

Private Sub MskFecFin_GotFocus()
    mskFecFin.SelStart = 0
    mskFecFin.SelLength = 40
End Sub

Private Sub MskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdImprimir.SetFocus
    End If
End Sub

Private Sub mskFecIni_GotFocus()
    mskFecIni.SelStart = 0
    mskFecIni.SelLength = 40
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mskFecFin.SetFocus
    End If
End Sub

Private Sub TV_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = "Cerrado"
End Sub

Private Sub TV_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = "Abierto"
End Sub

Private Sub TV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtTipCamFijAnt.SetFocus
    End If
End Sub

Private Sub TV_NodeClick(ByVal Node As MSComctlLib.Node)
    If IsNumeric(Right(Node.Key, 1)) Then
        If Left(Node.Key, 1) = "D" Then
            fraFechas.Visible = True
            fraMes.Visible = False
            cmdConcolidar.Enabled = True
        Else
            fraMes.Visible = True
            fraFechas.Visible = False
            cmdConcolidar.Enabled = False
        End If
    Else
        Me.fraMes.Visible = False
        Me.fraFechas.Visible = False
        Me.cmdConcolidar.Enabled = False
    End If
End Sub

Private Sub txtTipCamCompraSBS_GotFocus()
    txtTipCamCompraSBS.SelStart = 0
    txtTipCamCompraSBS.SelLength = 50
End Sub

Private Sub txtTipCamCompraSBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTipCamVentaSBS.SetFocus
    Else
        KeyAscii = NumerosDecimales(txtTipCamCompraSBS, KeyAscii, 8, 3)
    End If
End Sub

Private Sub txtTipCamFij_GotFocus()
    txtTipCamFij.SelStart = 0
    txtTipCamFij.SelLength = 50
End Sub

Private Sub txtTipCamFij_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtTipCamCompraSBS.SetFocus
    End If
End Sub

Private Function CuentasVigentesAho(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlM As String
    Dim rsM As ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Numero = Sum(nCtaVigAC), Monto = Sum(abs(nSaldoAC) + Abs(nChqCMAC) + Abs(nMonChqVal)), a=Sum(abs(nSaldoAC+nMonChqVal)) from EstadDiaACConsol where destadac between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Numero = Sum(nCtaVigAC), Monto = Sum(abs(nSaldoAC) + abs(nChqCMAC) + abs(nMonChqVal)) from EstadDiaACConsol where destadac between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If
    Set rsM = New ADODB.Recordset
    rsM.CursorLocation = adUseClient
    Set rsM = oConConsol.CargaRecordSet(sqlM)
    Set rsM.ActiveConnection = Nothing

    If rsM.EOF And rsM.BOF Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasVigentesAho = 0
        pnMonto = 0
    Else
        CuentasVigentesAho = IIf(IsNull(rsM!Numero), 0, rsM!Numero)
        pnMonto = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasVigentesPF(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlPF As String
    Dim rsPF As ADODB.Recordset
    Dim sqlCTS As String
    Dim rsCTS As ADODB.Recordset
    Dim lnTotCantidad As Long

    If pbConsol Then
        sqlCTS = "Select Numero = Sum(nNumVigCTS), Monto = Sum(nSaldoCTS+nMonChqVal) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
        sqlPF = "Select Numero = Sum(nNumVigPF), Monto = Sum(nSaldoPF+nMonChqVal) From EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' and dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlCTS = "Select Numero = Sum(nNumVigCTS),Monto = Sum(nSaldoCTS+nMonChqVal) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
        sqlPF = "Select Numero = Sum(nNumVigPF), Monto = Sum(nSaldoPF+nMonChqVal) From EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' and dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If
    Set rsPF = New ADODB.Recordset
    rsPF.CursorLocation = adUseClient
    Set rsPF = oConConsol.CargaRecordSet(sqlPF)
    Set rsPF.ActiveConnection = Nothing

    If rsPF.EOF And rsPF.BOF Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        lnTotCantidad = IIf(IsNull(rsPF!Numero), 0, rsPF!Numero)
        pnMonto = IIf(IsNull(rsPF!Monto), 0, Round(rsPF!Monto, 2))
    End If

    rsPF.Close
    Set rsPF = Nothing
    Set rsCTS = New ADODB.Recordset
    rsCTS.CursorLocation = adUseClient
    Set rsCTS = oConConsol.CargaRecordSet(sqlCTS)
    Set rsCTS.ActiveConnection = Nothing
    If rsCTS.EOF And rsCTS.BOF Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        CuentasVigentesPF = lnTotCantidad + IIf(IsNull(rsCTS!Numero), 0, rsCTS!Numero)
        pnMonto = pnMonto + IIf(IsNull(rsCTS!Monto), 0, Round(rsCTS!Monto, 2))
    End If

    rsCTS.Close
    Set rsCTS = Nothing

End Function

Private Function CuentasApeAho(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlM As String
    Dim rsM As New ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Numero = Sum(nNumAperAC), Monto = Sum(nMonAperAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Numero = Sum(nNumAperAC), Monto = Sum(nMonAperAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If

    Set rsM = oConConsol.CargaRecordSet(sqlM)

    If RSVacio(rsM) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasApeAho = 0
        pnMonto = 0
    Else
        CuentasApeAho = IIf(IsNull(rsM!Numero), 0, rsM!Numero)
        pnMonto = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasApePF(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlPF As String
    Dim rsPF As New ADODB.Recordset
    Dim sqlCTS As String
    Dim rsCTS As New ADODB.Recordset
    Dim lnTotCantidad As Long

    If pbConsol Then
        sqlCTS = "Select Numero = Sum(nNumAperCTS), Monto = Sum(nMonAperCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' And destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
        sqlPF = "Select Numero = Sum(nNumAperPF), Monto = Sum(nMonAperPF) from EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' And dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlCTS = "Select Numero = Sum(nNumAperCTS), Monto = Sum(nMonAperCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' And destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
        sqlPF = "Select Numero = Sum(nNumAperPF), Monto = Sum(nMonAperPF) from EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' And dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If

    Set rsPF = oConConsol.CargaRecordSet(sqlPF)

    If RSVacio(rsPF) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        lnTotCantidad = IIf(IsNull(rsPF!Numero), 0, rsPF!Numero)
        pnMonto = IIf(IsNull(rsPF!Monto), 0, Round(rsPF!Monto, 2))
    End If

    rsPF.Close
    Set rsPF = Nothing

    Set rsCTS = oConConsol.CargaRecordSet(sqlCTS)

    If RSVacio(rsCTS) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        CuentasApePF = lnTotCantidad + IIf(IsNull(rsCTS!Numero), 0, rsCTS!Numero)
        pnMonto = pnMonto + IIf(IsNull(rsCTS!Monto), 0, Round(rsCTS!Monto, 2))
    End If

    rsCTS.Close
    Set rsCTS = Nothing

End Function

Private Function CuentasAbonoAho(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlM As String
    Dim rsM As New ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Numero = Sum(nNumDepAC), Monto = Sum(nMonDepAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Numero = Sum(nNumDepAC), Monto = Sum(nMonDepAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If

    Set rsM = oConConsol.CargaRecordSet(sqlM)

    If RSVacio(rsM) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasAbonoAho = 0
        pnMonto = 0
    Else
        CuentasAbonoAho = IIf(IsNull(rsM!Numero), 0, rsM!Numero)
        pnMonto = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasAbonoPF(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlM As String
    Dim rsM As New ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Numero = Sum(nNumDepCTS), Monto = Sum(nMonDepCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Numero = Sum(nNumDepCTS), Monto = Sum(nMonDepCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If

    Set rsM = oConConsol.CargaRecordSet(sqlM)

    If RSVacio(rsM) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasAbonoPF = 0
        pnMonto = 0
    Else
        CuentasAbonoPF = IIf(IsNull(rsM!Numero), 0, rsM!Numero)
        pnMonto = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasIntCapAho(psFecIni As String, psFecFin As String, psMoneda As String, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Currency
    Dim sqlM As String
    Dim rsM As New ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Monto = Sum(nIntCapAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Monto = Sum(nIntCapAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If

    Set rsM = oConConsol.CargaRecordSet(sqlM)

    If RSVacio(rsM) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasIntCapAho = 0
    Else
        CuentasIntCapAho = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasIntCapPF(psFecIni As String, psFecFin As String, psMoneda As String, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Currency
    Dim sqlPF As String
    Dim rsPF As New ADODB.Recordset
    Dim sqlCTS As String
    Dim rsCTS As New ADODB.Recordset
    Dim lnTotCantidad As Currency

    If pbConsol Then
        sqlCTS = "Select Monto = Sum(nIntCapCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
        sqlPF = "Select Monto = Sum(nIntCapPF) from EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' and dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlCTS = "Select Monto = Sum(nIntCapCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
        sqlPF = "Select Monto = Sum(nIntCapPF) from EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' and dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If

    Set rsPF = oConConsol.CargaRecordSet(sqlPF)

    If RSVacio(rsPF) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        lnTotCantidad = IIf(IsNull(rsPF!Monto), 0, Round(rsPF!Monto, 2))
    End If

    rsPF.Close
    Set rsPF = Nothing

    Set rsCTS = oConConsol.CargaRecordSet(sqlCTS)

    If RSVacio(rsCTS) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        CuentasIntCapPF = lnTotCantidad + IIf(IsNull(rsCTS!Monto), 0, Round(rsCTS!Monto, 2))
    End If

    rsCTS.Close
    Set rsCTS = Nothing

End Function

Private Function CuentasCanceladaAho(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlM As String
    Dim rsM As New ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Numero = Sum(nNumCancAC + nNumCancInacAC), Monto = Sum(nMonCancAC + nMonCancInacAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Numero = Sum(nNumCancAC + nNumCancInacAC), Monto = Sum(nMonCancAC + nMonCancInacAC) from EstadDiaACConsol where destadac >= '" & psFecIni & "' and destadac < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If
    'sqlM = "Select Numero = Sum(nNumCancAC + nNumCancInacAC), Monto = Sum(nMonCancAC + nMonCancInacAC) from EstadDiaACConsol where destadac between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Set rsM = oConConsol.CargaRecordSet(sqlM)

    If RSVacio(rsM) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasCanceladaAho = 0
        pnMonto = 0
    Else
        CuentasCanceladaAho = IIf(IsNull(rsM!Numero), 0, rsM!Numero)
        pnMonto = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasCanceladaPF(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlPF As String
    Dim rsPF As New ADODB.Recordset
    Dim sqlCTS As String
    Dim rsCTS As New ADODB.Recordset
    Dim lnTotCantidad As Long

    If pbConsol Then
        sqlCTS = "Select Numero = Sum(nNumCancCTS), Monto = Sum(nMonCancCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
        sqlPF = "Select Numero = Sum(nNumCancPF), Monto = Sum(nMonCancPF) from EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' and dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlCTS = "Select Numero = Sum(nNumCancCTS), Monto = Sum(nMonCancCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge ='" & psCodAge & "'"
        sqlPF = "Select Numero = Sum(nNumCancPF), Monto = Sum(nMonCancPF) from EstadDiaPFConsol where dEstadPF >= '" & psFecIni & "' and dEstadPF < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge ='" & psCodAge & "'"
    End If

    Set rsPF = oConConsol.CargaRecordSet(sqlPF)

    If RSVacio(rsPF) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        lnTotCantidad = IIf(IsNull(rsPF!Numero), 0, rsPF!Numero)
        pnMonto = IIf(IsNull(rsPF!Monto), 0, Round(rsPF!Monto, 2))
    End If

    rsPF.Close
    Set rsPF = Nothing

    Set rsCTS = oConConsol.CargaRecordSet(sqlCTS)

    If RSVacio(rsCTS) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        CuentasCanceladaPF = lnTotCantidad + IIf(IsNull(rsCTS!Numero), 0, rsCTS!Numero)
        pnMonto = pnMonto + IIf(IsNull(rsCTS!Monto), 0, Round(rsCTS!Monto, 2))
    End If

    rsCTS.Close
    Set rsCTS = Nothing

End Function

Private Function CuentasRetAho(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlM As String
    Dim rsM As New ADODB.Recordset

    If pbConsol Then
        sqlM = "Select Numero = Sum(nNumRetAC), Monto = Sum(nRetInacAC + nMonRetAC) from EstadDiaACConsol where destadac between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlM = "Select Numero = Sum(nNumRetAC), Monto = Sum(nRetInacAC + nMonRetAC) from EstadDiaACConsol where destadac between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If
    'sqlM = "Select Numero = Sum(nNumRetAC), Monto = Sum(nRetInacAC + nMonRetAC) from EstadDiaACConsol where destadac between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Set rsM = oConConsol.CargaRecordSet(sqlM)

    If RSVacio(rsM) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        CuentasRetAho = 0
        pnMonto = 0
    Else
        CuentasRetAho = IIf(IsNull(rsM!Numero), 0, rsM!Numero)
        pnMonto = IIf(IsNull(rsM!Monto), 0, Round(rsM!Monto, 2))
    End If

    rsM.Close
    Set rsM = Nothing
End Function

Private Function CuentasRetPF(psFecIni As String, psFecFin As String, psMoneda As String, pnMonto As Currency, Optional pbConsol As Boolean = False, Optional psCodAge As String = "") As Long
    Dim sqlCTS As String
    Dim rsCTS As New ADODB.Recordset
    Dim lnTotCantidad As Long

    If pbConsol Then
        sqlCTS = "Select Numero = Sum(nNumRetCTS), Monto = Sum(nMonRetCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Else
        sqlCTS = "Select Numero = Sum(nNumRetCTS), Monto = Sum(nMonRetCTS) from EstadDiaCTS where destadcts >= '" & psFecIni & "' and destadcts < '" & psFecFin & "' and cMoneda = '" & psMoneda & "' and cCodAge = '" & psCodAge & "'"
    End If
    'sqlCTS = "Select Numero = Sum(nNumRetCTS), Monto = Sum(nMonRetCTS) from EstadDiaCTS where destadcts between '" & psFecIni & "' and '" & psFecFin & "' and cMoneda = '" & psMoneda & "'"
    Set rsCTS = oConConsol.CargaRecordSet(sqlCTS)

    If RSVacio(rsCTS) Then
        MsgBox "No existe información sobre el numero cuentas vigentes para este día.", vbInformation, "Aviso"
        lnTotCantidad = 0
    Else
        CuentasRetPF = IIf(IsNull(rsCTS!Numero), 0, rsCTS!Numero)
        pnMonto = IIf(IsNull(rsCTS!Monto), 0, Round(rsCTS!Monto, 2))
    End If

    rsCTS.Close
    Set rsCTS = Nothing
End Function

Private Function ValidaM() As Boolean
Dim sFecha As String
sFecha = txtFecha.Text

    If sFecha = "__/__/____" Then
        MsgBox "Debe colocar una fecha válida para el reporte.", vbInformation, "Aviso"
        txtFecha.SetFocus
        ValidaM = True
    ElseIf Not IsDate(sFecha) Then
        MsgBox "Debe ingresar una fecha válida.", vbInformation, "Aviso"
        txtFecha.SetFocus
        ValidaM = True
    Else
        ValidaM = False
    End If

End Function

Private Function ValidaD() As Boolean
    If Not IsDate(mskFecIni) Then
        MsgBox "Debe Ingresar una fecha Valida.", vbInformation, "Aviso"
        mskFecIni.SetFocus
        ValidaD = True
    ElseIf Not IsDate(mskFecFin) Then
        MsgBox "Debe Ingresar una fecha Valida.", vbInformation, "Aviso"
        mskFecFin.SetFocus
        ValidaD = True
    Else
        ValidaD = False
    End If

End Function

Private Function GetMovCompraVentaME(pdFecIni As Date, pdFecFin As Date, Optional psCodAge As String = "") As String
    Dim sqlC As String
    Dim sqlV As String
    Dim sqlT As String

    Dim I As Integer

    Dim lnConsolTotal As Currency
    Dim lnConsolNumero As Currency
    Dim lnConsolTipCambio As Currency

    Dim lnTotal As Currency
    Dim lnNumero As Currency
    Dim lnTipCambio As Currency

    Dim lnMonComDol As Currency
    Dim lnMonVenDol As Currency
    Dim lnMonComSol As Currency
    Dim lnMonVenSol As Currency
    Dim lnPromCom As Currency
    Dim lnPromVen As Currency
    Dim lnCotMaxCom As Currency
    Dim lnCotMaxVen As Currency
    Dim lnCotMinCom As Currency
    Dim lnCotMinVen As Currency

    Dim rsRes As New ADODB.Recordset
    Dim lnNumAgeNoRes As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion

    lnNumAgeNoRes = 0
    sqlC = " Select TotalDol = Sum(nMonTran), TotalSol = Sum(nMonTran * nTipCambio), Numero = Count(nTipCambio), TipCambio = Sum(nTipCambio) From CompraVenta Where cCodOpe = '230101'" _
         & " and dFecTran Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' and '" & Format(DateAdd("d", 1, pdFecFin), "mm/dd/yyyy") & "'"
    sqlV = " Select TotalDol = Sum(nMonTran), TotalSol = Sum(nMonTran * nTipCambio), Numero = Count(nTipCambio), TipCambio = Sum(nTipCambio) From CompraVenta Where cCodOpe = '230102'" _
         & " and dFecTran Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' and '" & Format(DateAdd("d", 1, pdFecFin), "mm/dd/yyyy") & "'"
    sqlT = " Select ProCom = AVG(nValComp), ProVen = AVG(nValVent), " _
         & " MaxCom = Max(nValComp), MaxVen = Max(nValVent)," _
         & " MinCom = Max(nValComp), MinVen = Min(nValVent)" _
         & " From TipCambio" _
         & " Where dFecCamb Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' and '" & Format(DateAdd("d", 1, pdFecFin), "mm/dd/yyyy") & "'"

    If psCodAge <> "" Then
        sqlC = sqlC & " And cCodAge = '" & psCodAge & "'"
        sqlV = sqlV & " And cCodAge = '" & psCodAge & "'"

        Set rsRes = oCon.CargaRecordSet(sqlC)
        
        If RSVacio(rsRes) Then
            lnMonComDol = 0
            lnMonComSol = 0
        Else
            lnMonComDol = IIf(IsNull(rsRes!TotalDol), 0, rsRes!TotalDol)
            lnMonComSol = IIf(IsNull(rsRes!TotalSol), 0, rsRes!TotalSol)
        End If

        rsRes.Close
        Set rsRes = Nothing

        Set rsRes = oCon.CargaRecordSet(sqlV)
        
        If RSVacio(rsRes) Then
            lnMonVenDol = 0
            lnMonVenSol = 0
        Else
            lnMonVenDol = IIf(IsNull(rsRes!TotalDol), 0, rsRes!TotalDol)
            lnMonVenSol = IIf(IsNull(rsRes!TotalSol), 0, rsRes!TotalSol)
        End If

        rsRes.Close
        Set rsRes = Nothing

        Set rsRes = oCon.CargaRecordSet(sqlT)
        
        If RSVacio(rsRes) Then
            lnPromCom = 0
            lnCotMaxCom = 0
            lnCotMinCom = 0

            lnPromVen = 0
            lnCotMaxVen = 0
            lnCotMinVen = 0
        Else
            lnPromCom = IIf(IsNull(rsRes!ProCom), 0, rsRes!ProCom)
            lnCotMaxCom = IIf(IsNull(rsRes!MaxCom), 0, rsRes!MaxCom)
            lnCotMinCom = IIf(IsNull(rsRes!MinCom), 0, rsRes!MinCom)

            lnPromVen = IIf(IsNull(rsRes!ProVen), 0, rsRes!ProVen)
            lnCotMaxVen = IIf(IsNull(rsRes!Maxven), 0, rsRes!Maxven)
            lnCotMinVen = IIf(IsNull(rsRes!MinVen), 0, rsRes!MinVen)
        End If

        rsRes.Close
        Set rsRes = Nothing

        'CierraConeccion
    Else
        Set rsRes = oCon.CargaRecordSet(sqlC)
        
        If RSVacio(rsRes) Then
            lnMonComDol = lnMonComDol + 0
            lnMonComSol = lnMonComSol + 0
        Else
            lnMonComDol = lnMonComDol + IIf(IsNull(rsRes!TotalDol), 0, rsRes!TotalDol)
            lnMonComSol = lnMonComSol + IIf(IsNull(rsRes!TotalSol), 0, rsRes!TotalSol)
        End If

            rsRes.Close
            Set rsRes = Nothing

            Set rsRes = oCon.CargaRecordSet(sqlV)
            
            If RSVacio(rsRes) Then
                lnMonVenDol = lnMonVenDol + 0
                lnMonVenSol = lnMonVenSol + 0
            Else
                lnMonVenDol = lnMonVenDol + IIf(IsNull(rsRes!TotalDol), 0, rsRes!TotalDol)
                lnMonVenSol = lnMonVenSol + IIf(IsNull(rsRes!TotalSol), 0, rsRes!TotalSol)
            End If

            rsRes.Close
            Set rsRes = Nothing

            Set rsRes = oCon.CargaRecordSet(sqlT)
            
            If RSVacio(rsRes) Then
                lnPromCom = lnPromCom + 0
                lnCotMaxCom = lnCotMaxCom + 0
                lnCotMinCom = lnCotMinCom + 0

                lnPromVen = lnPromVen + 0
                lnCotMaxVen = lnCotMaxVen + 0
                lnCotMinVen = lnCotMinVen + 0
            Else
                lnPromCom = lnPromCom + IIf(IsNull(rsRes!ProCom), 0, rsRes!ProCom)
                lnCotMaxCom = lnCotMaxCom + IIf(IsNull(rsRes!MaxCom), 0, rsRes!MaxCom)
                lnCotMinCom = lnCotMinCom + IIf(IsNull(rsRes!MinCom), 0, rsRes!MinCom)

                lnPromVen = lnPromVen + IIf(IsNull(rsRes!ProVen), 0, rsRes!ProVen)
                lnCotMaxVen = lnCotMaxVen + IIf(IsNull(rsRes!Maxven), 0, rsRes!Maxven)
                lnCotMinVen = lnCotMinVen + IIf(IsNull(rsRes!MinVen), 0, rsRes!MinVen)
            End If

            rsRes.Close
            Set rsRes = Nothing

        lnPromCom = lnPromCom
        lnCotMaxCom = lnCotMaxCom
        lnCotMinCom = lnCotMinCom

        lnPromVen = lnPromVen
        lnCotMaxVen = lnCotMaxVen
        lnCotMinVen = lnCotMinVen

    End If

    GetMovCompraVentaME = DibujaCompraVenta(lnMonComDol, lnMonVenDol, lnMonComSol, lnMonVenSol, lnPromCom, lnPromVen, lnCotMaxCom, lnCotMaxVen, lnCotMinCom, lnCotMinVen, psCodAge, pdFecIni, pdFecFin)
End Function

Private Function GetMovCompraVentaMEConsolida(pdFecIni As Date, pdFecFin As Date, Optional psCodAge As String = "") As String
    Dim sqlC As String
    Dim sqlRes As String
    Dim rsRes As New ADODB.Recordset
    Dim lsCadena As String
    Dim lnNumAgeNoRes As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim I As Integer
    oCon.AbreConexion
    
    sqlRes = " Select dFecTran,nNumTran,cCodUsu,cCodOpe,nMonTran,nTipCambio,cCodAge From TransAho" _
           & " where dFecTran Between '" & Format(pdFecIni, "mm/dd/yyyy hh:mm:ss ") & "' And '" & Format(pdFecFin, "mm/dd/yyyy hh:mm:ss ") & "' and cCodOpe in ('230101','230102') and cFlag is Null"

    If psCodAge <> "" Then
         If oConConsol.AbreConexion Then 'Remota(Right(psCodAge, 2), , False)
            Set rsRes = oConConsol.CargaRecordSet(sqlRes)
         Else
            Exit Function
         End If

        If Not RSVacio(rsRes) Then
            oCon.BeginTrans
                sqlC = "Delete CompraVenta where dFecTran Between '" & Format(pdFecIni, "mm/dd/yyyy hh:mm:ss ") & "' And '" & Format(pdFecFin, "mm/dd/yyyy hh:mm:ss ") & "' And cCodAge = '" & psCodAge & "'"
                oCon.Ejecutar sqlC
                While Not rsRes.EOF
                    sqlC = " Insert CompraVenta (dFecTran,nNumTran,cCodUsu,cCodOpe,nMonTran,nTipCambio,cCodAge)" _
                         & " Values('" & Format(rsRes!DFECTRAN, "mm/dd/yyyy hh:mm:ss ") & "'," & rsRes!nnumTran & ",'" & rsRes!cCodUsu & "'," _
                         & " '" & rsRes!cCodOpe & "'," & rsRes!nMonTran & "," & rsRes!nTipCambio & ",'" & rsRes!cCodAge & "')"
                    oCon.Ejecutar sqlC
                    rsRes.MoveNext
                Wend
            oCon.CommitTrans
        Else
            lsCadena = lsCadena & psCodAge & Space(10) & "- No tiene Compra/Venta." & oImpresora.gPrnSaltoLinea
        End If

        rsRes.Close
        Set rsRes = Nothing

        oConConsol.CierraConexion
    Else
        For I = 0 To lstAge.ListCount - 2
            If oConConsol.AbreConexion Then 'Remota(Right(psCodAge, 2), True, False)
                Set rsRes = oConConsol.CargaRecordSet(sqlRes)
                
                If Not RSVacio(rsRes) Then
                    oConConsol.BeginTrans
                        sqlC = "Delete CompraVenta where dFecTran Between '" & Format(pdFecIni, "mm/dd/yyyy hh:mm:ss ") & "' And '" & Format(pdFecFin, "mm/dd/yyyy hh:mm:ss ") & "' and cCodAge = '" & Right(Me.lstAge.List(I), 5) & "'"
                        oConConsol.Ejecutar sqlC
                        While Not rsRes.EOF
                            sqlC = " Insert CompraVenta (dFecTran,nNumTran,cCodUsu,cCodOpe,nMonTran,nTipCambio,cCodAge)" _
                                 & " Values('" & Format(rsRes!DFECTRAN, "mm/dd/yyyy hh:mm:ss ") & "'," & rsRes!nnumTran & ",'" & rsRes!cCodUsu & "'," _
                                 & " '" & rsRes!cCodOpe & "'," & rsRes!nMonTran & "," & rsRes!nTipCambio & ",'" & rsRes!cCodAge & "')"
                            oConConsol.Ejecutar sqlC
                            rsRes.MoveNext
                        Wend
                    oConConsol.CommitTrans
                Else
                    lsCadena = lsCadena & Left(lstAge.List(I), 20) & Space(10) & "- No tiene Compra/Venta." & oImpresora.gPrnSaltoLinea
                End If

                rsRes.Close
                Set rsRes = Nothing

                oConConsol.CierraConexion
            Else
                lnNumAgeNoRes = lnNumAgeNoRes + 1
            End If
        Next I
    End If

    GetMovCompraVentaMEConsolida = lsCadena
End Function

Private Function DibujaCompraVenta(pnMonComDol As Currency, pnMonVenDol As Currency, _
                                   pnMonComSol As Currency, pnMonVenSol As Currency, _
                                   pnPromCom As Currency, pnPromVen As Currency, _
                                   pnCotMaxCom As Currency, pnCotMaxVen As Currency, _
                                   pnCotMinCom As Currency, pnCotMinVen As Currency, _
                                   psCodAge As String, pdFecIni As Date, pdFecFin As Date)

    Dim lsCadena As String
    Dim lnMargen As Integer
    Dim lsNombreAg As String

    If psCodAge <> "" Then
        lsNombreAg = psCodAge
    Else
        lsNombreAg = "CONSOLIDADO"
    End If

    lnMargen = 5
    lsCadena = ""

    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("Listado de Compra Venta de Dolares - " & lsNombreAg, 140) & oImpresora.gPrnSaltoLinea
    If DateDiff("d", pdFecIni, pdFecFin) = 0 Then
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("Dia: " & Format(pdFecIni, "dd/mm/yyyy"), 140) & oImpresora.gPrnSaltoLinea
    Else
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("Rango: " & Format(pdFecIni, "dd/mm/yyyy") & " - " & Format(pdFecFin, "dd/mm/yyyy"), 140) & oImpresora.gPrnSaltoLinea
    End If
    lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea

    lsCadena = lsCadena & Space(lnMargen) & "+" & String(6, "-") & "+" & String(9, "-") & "+" & String(18, "-") & "+" & String(19, "-") & "+" & String(18, "-") & "+" & String(18, "-") & "+" & String(10, "-") & "+" & String(15, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & "      " & "¦" & "         " & "¦" & "                  " & "¦" & "                   " & "¦" & "                  " & "¦" & " Monto en Moneda  " & "¦" & "          " & "¦" & "               " & "¦" & "               " & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & " Fila " & "¦" & " Monedas " & "¦" & " Codigo de Moneda " & "¦" & " Tipo de Operacion " & "¦" & " Monto en Dolares " & "¦" & "     Nacional     " & "¦" & " Promedio " & "¦" & " Cotiz. Maxima " & "¦" & " Cotiz. Minima " & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(6, "-") & "+" & String(9, "-") & "+" & String(18, "-") & "+" & String(19, "-") & "+" & String(18, "-") & "+" & String(18, "-") & "+" & String(10, "-") & "+" & String(15, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & "  100 " & "¦" & " Datos   " & "¦" & " Dolar de N.A.    " & "¦" & " Compra            " & "¦" & Space(18 - Len(Format(pnMonComDol, "#,##0.00"))) & Format(pnMonComDol, "#,##0.00") & "¦" & Space(18 - Len(Format(pnMonComSol, "#,##0.00"))) & Format(pnMonComSol, "#,##0.00") & "¦" & Space(10 - Len(Format(pnPromCom, "#,##0.000000"))) & Format(pnPromCom, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMaxCom, "#,##0.000000"))) & Format(pnCotMaxCom, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMinCom, "#,##0.000000"))) & Format(pnCotMinCom, "#,##0.000000") & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & "  101 " & "¦" & "         " & "¦" & " Dolar de N.A.    " & "¦" & " Venta             " & "¦" & Space(18 - Len(Format(pnMonVenDol, "#,##0.00"))) & Format(pnMonVenDol, "#,##0.00") & "¦" & Space(18 - Len(Format(pnMonVenSol, "#,##0.00"))) & Format(pnMonVenSol, "#,##0.00") & "¦" & Space(10 - Len(Format(pnPromVen, "#,##0.000000"))) & Format(pnPromVen, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMaxVen, "#,##0.000000"))) & Format(pnCotMaxVen, "#,##0.000000") & "¦" & Space(15 - Len(Format(pnCotMinVen, "#,##0.000000"))) & Format(pnCotMinVen, "#,##0.000000") & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(6, "-") & "+" & String(9, "-") & "+" & String(18, "-") & "+" & String(19, "-") & "+" & String(18, "-") & "+" & String(18, "-") & "+" & String(10, "-") & "+" & String(15, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea

    DibujaCompraVenta = lsCadena
End Function

Public Function SaldoCaptaciones(ByVal dFecha As Date) As String
    Dim m As Integer, P As Integer, I As Byte, J As Byte
    Dim rsEstad As ADODB.Recordset
    Dim rsCOM As ADODB.Recordset
    Dim lnRegCod As Integer
    Dim lsValor As String, lsNumPag As String
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsTipMon As String, lsTipPer As String
    Dim lnCarLin As Integer, lnCntPag As Integer
    Dim lnAge As Integer
    Dim lsAge As String * 13
    Dim lsDefAho(1 To 9) As String
    Dim lsDefPlazo(1 To 11) As String
    Dim lsDefInt(1 To 4) As String

    Dim lnTotal(1 To 2, 1 To 46) As Double
    Dim lnTotalT(1 To 2, 1 To 23) As Double
    Dim lnInteres(1 To 2, 1 To 23) As Double

    Dim lsTotal(1 To 3) As String
    Dim lsConn As String
    Dim lsRTFRfm As String
    Dim sql As String
    Dim lsMoneda As String

    lsDefAho(1) = "Pers. Naturales           "
    lsDefAho(2) = "Pers. Jur. Sin F. Lucro   "
    lsDefAho(3) = "Pers. Jur. Con F. Lucro   "
    lsDefAho(4) = "FONCODES                  "
    lsDefAho(5) = "CMACs                     "
    lsDefAho(6) = "CRACs                     "
    lsDefAho(7) = "Inactivas Pers. Nat.      "
    lsDefAho(8) = "Inactivas Pers. Jur.S.F.L."
    lsDefAho(9) = "Inactivas Pers. Jur.C.F.L."

    lsDefPlazo(1) = "PF Pers. Naturales        "
    lsDefPlazo(2) = "PF Pers. Jur. Sin F. Lucro"
    lsDefPlazo(3) = "PF Pers. Jur. Con F. Lucro"
    lsDefPlazo(4) = "CMACs                     "
    lsDefPlazo(5) = "CRACs                     "
    lsDefPlazo(6) = "CTS Ordinarios            "
    lsDefPlazo(7) = "CTS Convenio              "

    lsDefInt(1) = "Int. Dev. Nat. y J.S.F.L. "
    lsDefInt(2) = "Int. Dev. Otras Pers. Jur."
    lsDefInt(3) = "Int. Dev. CMACs           "
    lsDefInt(4) = "Int. Dev. CRACs           "

    lsTotal(1) = "TOTAL DEPOSITOS DE AHORROS"
    lsTotal(2) = "TOTAL DEPOSITOS A PLAZO   "
    lsTotal(3) = "TOTAL INTERES DEVENGADO   "

    lsRTFRfm = ""
    
    oConConsol.AbreConexion 'Remota "07", , , "03"

    Set rsEstad = New ADODB.Recordset
    rsEstad.CursorLocation = adUseClient
    For m = 1 To 2
        For I = 1 To 20
            lnTotal(m, I) = 0
            lnInteres(m, I) = 0
        Next I

        sql = "Select E.cCodAge, TC.cNomTab, E.nSaldPN, E.nSaldPSFL, E.nSaldPCFL, E.nSaldCMAC, E.nSaldCRAC, E.nSaldFoncodes, E.nSaldInPN, E.nSaldInPSFL, E.nSaldInPCFL " _
            & "FROM EstadMensAho E INNER JOIN TablaCod TC ON E.cCodAge = TC.cValor WHERE DateDiff(dd,E.dEstadMens,'" & Format$(dFecha, "mm/dd/yyyy") & "') = 0 " _
            & "AND E.cMoneda = '" & m & "' And TC.cCodTab like '47__' ORDER BY cCodAge"
        If rsEstad.State = adStateOpen Then rsEstad.Close

        Set rsEstad = oConConsol.CargaRecordSet(sql)
        Set rsEstad.ActiveConnection = Nothing
        lnCarLin = 206 + (15 * (rsEstad.Fields.Count - 9))

        lsTipMon = Trim(Str(m))
        lnCntPag = m
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        lsMoneda = IIf(m = 1, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
        lsTitRp1 = "S I T U A C I O N     D E    C A P T A C I O N E S    P O R    T I P O    D E    P E R S O N A"
        lsTitRp2 = "AL " & Format$(dFecha, "DD MMM YYYY")
        lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lnAge = 0
        If rsEstad.EOF And rsEstad.BOF Then
            MsgBox "No se encontraron registros para el rango de fechas especificados", vbInformation, "Aviso"
            SaldoCaptaciones = ""
            Exit Function
        End If
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & Space(28)
        Do While Not rsEstad.EOF
            lsAge = Trim(rsEstad!cNomtab)
            lsRTFRfm = lsRTFRfm & lsAge & Space(2)
            lnAge = lnAge + 1
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & "    TOTAL    " & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea

        'DEPOSITOS DE AHORROS
        I = 0: J = 1
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(1) & Space(2) 'Pers. Nat.
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldPN), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldPN
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldPN
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 2
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(2) & Space(2) 'Pers.Jur.Sin Fines Lucro
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldPSFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldPSFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldPSFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea
        I = 0: J = 3
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(3) & Space(2) 'Pers.Jur.Con Fines de Lucro
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldPCFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldPCFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldPCFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea
        I = 0: J = 4
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(4) & Space(2) 'FONCODES
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldFoncodes), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldFoncodes
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldFoncodes
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 5
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(5) & Space(2) 'CMACs
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldCMAC), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldCMAC
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldCMAC
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 6
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(6) & Space(2) 'CRACs
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldCRAC), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldCRAC
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldCRAC
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 7
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(7) & Space(2) 'Inactivas Pers. Nat.
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldInPN), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldInPN
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldInPN
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 8
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(8) & Space(2) 'Inactivas Pers.Jur.Sin Fines Lucro
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldInPSFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldInPSFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldInPSFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea
        I = 0: J = 9
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefAho(9) & Space(2) 'Inactivas Pers.Jur.Con Fines Lucro
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldInPCFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldInPCFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldInPCFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea
        rsEstad.Close

        J = 10 'Imprime Total Depositos Ahorros
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & lsTotal(1) & Space(2)
        For I = 1 To lnAge
            lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotal(m, I)), 12, True, 9, 2) & Space(3)
            lnTotalT(m, J) = lnTotalT(m, J) + lnTotal(m, I)
        Next I

        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea

        'DEPOSITOS A PLAZO
        sql = "Select E.nSaldPN, E.nSaldPSFL, E.nSaldPCFL, E.nSaldCMAC, E.nSaldCRAC FROM EstadMensPF E WHERE " _
            & "DateDiff(dd,E.dEstadMens,'" & Format$(dFecha, "mm/dd/yyyy") & "') = 0 " _
            & "AND E.cMoneda = '" & m & "' ORDER BY E.cCodAge "

        Set rsEstad = oConConsol.CargaRecordSet(sql)
        Set rsEstad.ActiveConnection = Nothing

        I = lnAge: J = 11
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(1) & Space(2) 'Pers. Nat.
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldPN), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldPN
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldPN
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = lnAge: J = 12
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(2) & Space(2) 'Pers.Jur.Sin Fines Lucro
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldPSFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldPSFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldPSFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = lnAge: J = 13
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(3) & Space(2) 'Pers.Jur.Con Fines Lucro
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldPCFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldPCFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldPCFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = lnAge: J = 14
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(4) & Space(2) 'CMACs
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldCMAC), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldCMAC
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldCMAC
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = lnAge: J = 15
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(5) & Space(2) 'CRACs
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldCRAC), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldCRAC
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldCRAC
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        rsEstad.Close

        sql = "Select E.cCodAge, TC.cNomTab, E.nSaldOrd, E.nSaldConv, E.nIntDev " _
            & "FROM EstadMensCTS E INNER JOIN TablaCod TC ON E.cCodAge = TC.cValor WHERE " _
            & "DateDiff(dd,E.dEstadMens,'" & Format$(dFecha, "mm/dd/yyyy") & "') = 0 " _
            & "AND E.cMoneda = '" & m & "' And TC.cCodTab like '47__' ORDER BY cCodAge"

        Set rsEstad = oConConsol.CargaRecordSet(sql)
        Set rsEstad.ActiveConnection = Nothing

        I = lnAge: J = 16
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(6) & Space(2) 'CTS Ordinarios
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldOrd), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldOrd
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldOrd
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = lnAge: J = 17
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefPlazo(6) & Space(2) 'CTS Convenios
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nSaldConv), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnTotal(m, I) = lnTotal(m, I) + rsEstad!nSaldConv
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nSaldConv
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea
        rsEstad.Close

        J = 18 'Imprime Total depósitos a Plazo
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & lsTotal(2) & Space(2)
        For I = lnAge + 1 To lnAge * 2
            lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotal(m, I)), 12, True, 9, 2) & Space(3)
            lnTotalT(m, J) = lnTotalT(m, J) + lnTotal(m, I)
        Next I

        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea


        'INTERESES DEVENGADOS
        sql = "Select E.nIntDevPNJSFL, E.nIntDevOtros, E.nIntDevCMAC, E.nIntDevCRAC FROM EstadMensPF E WHERE " _
            & "DateDiff(dd,E.dEstadMens,'" & Format$(dFecha, "mm/dd/yyyy") & "') = 0 " _
            & "AND E.cMoneda = '" & m & "' ORDER BY E.cCodAge"

        Set rsEstad = oConConsol.CargaRecordSet(sql)
        Set rsEstad.ActiveConnection = Nothing

        'Calculo de los Intereses Devengados de Plazo Fijo
        I = 0: J = 19 'Persona Natural y JSFL
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefInt(1) & Space(2)
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nIntDevPNJSFL), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnInteres(m, I) = lnInteres(m, I) + rsEstad!nIntDevPNJSFL
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nIntDevPNJSFL
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 20 'Otras Personas Juridicas
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefInt(2) & Space(2)
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nIntDevOtros), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnInteres(m, I) = lnInteres(m, I) + rsEstad!nIntDevOtros
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nIntDevOtros
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea

        I = 0: J = 21 'CMACs
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefInt(3) & Space(2)
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nIntDevCMAC), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnInteres(m, I) = lnInteres(m, I) + rsEstad!nIntDevCMAC
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nIntDevCMAC
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea


        I = 0: J = 22 'CRACs
        rsEstad.MoveFirst
        lsRTFRfm = lsRTFRfm & lsDefInt(3) & Space(2)
        Do While Not rsEstad.EOF
            lsRTFRfm = lsRTFRfm & JDNum(Str(rsEstad!nIntDevCRAC), 12, False, 9, 2) & Space(3)
            I = I + 1
            lnInteres(m, I) = lnInteres(m, I) + rsEstad!nIntDevCRAC
            lnTotalT(m, J) = lnTotalT(m, J) + rsEstad!nIntDevCRAC
            rsEstad.MoveNext
        Loop
        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, False, 9, 2) & oImpresora.gPrnSaltoLinea
        rsEstad.Close

        J = 23 'Imprime Total Interes Devengado
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & lsTotal(3) & Space(2)
        For I = 1 To lnAge
            lsRTFRfm = lsRTFRfm & JDNum(Str(lnInteres(m, I)), 12, True, 9, 2) & Space(3)
            lnTotalT(m, J) = lnTotalT(m, J) + lnInteres(m, I)
        Next I

        lsRTFRfm = lsRTFRfm & JDNum(Str(lnTotalT(m, J)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea

        lsRTFRfm = lsRTFRfm & oImpresora.gPrnSaltoPagina
    Next m
    Set rsEstad = Nothing
    SaldoCaptaciones = lsRTFRfm
End Function

Private Sub TxtTipCamFijAnt_GotFocus()
    TxtTipCamFijAnt.SelStart = 0
    TxtTipCamFijAnt.SelLength = 50
End Sub

Private Sub TxtTipCamFijAnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtTipCamFij.SetFocus
    End If
End Sub

Private Sub txtTipCamVentaSBS_GotFocus()
    txtTipCamVentaSBS.SelStart = 0
    txtTipCamVentaSBS.SelLength = 50
End Sub

Private Sub txtTipCamVentaSBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstAge.SetFocus
    Else
        KeyAscii = NumerosDecimales(txtTipCamVentaSBS, KeyAscii, 8, 3)
    End If
End Sub


Private Function GetGanCompVenta(pdFecha As Date, pnTCF As Single, pnTCC As Single, pnTCV As Single) As String
Dim lsConn As String
Dim rsCOM As New ADODB.Recordset
Dim ldIni As Date, ldFin As Date
Dim ln1915D As Currency, ln1915H As Currency
Dim ln2825D As Currency, ln2825H As Currency
Dim lnCompra As Currency, lnVenta As Currency
Dim lnAjuste As Currency
Dim ls1915D As String, ls1915H As String
Dim ls2825D As String, ls2825H As String
Dim lsCompra As String, lsVenta As String
Dim lsSaldo19 As String, lsSaldo28 As String
Dim nTCC As Single, nTCV As Single, nTCF As Single
Dim lnCarLin As Integer
Dim lsTitRp1 As String
Dim lsTitRp2 As String
Dim lsMoneda As String
Dim lsNumPag As String
Dim lsFecha As String
Dim lnLinPag As Integer
Dim lnCntPag As Integer
Dim lsRTFExt As String
Dim bPrint As Boolean
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta
Dim gsCVComME As String
Dim gsCVVenME As String

oCon.AbreConexion

ldIni = CDate("01" & "/" & Month(pdFecha) & "/" & Year(pdFecha))
ldFin = DateAdd("m", 1, ldIni)

nTCC = pnTCC
nTCV = pnTCV
nTCF = pnTCF

gsCVComME = "230101"
gsCVVenME = "230102"

sql = "SELECT CONVERT(Varchar(10),dFecTran,103) cFecha, cCodOpe, SUM(ABS(nMonTran)*nTipCambio) nMontoCV, " _
    & "SUM(ABS(nMonTran)* " & nTCF & ") nMontoF, SUM(ABS(nMonTran)) nMonto from CompraVenta where " _
    & "DateDiff(mm,dFecTran,'" & Format$(ldIni, "mm/dd/yyyy") & "') = 0 " _
    & "AND cCodOpe in ('" & gsCVComME & "','" & gsCVVenME & "') " _
    & "GROUP BY CONVERT(Varchar(10),dFecTran,103), cCodOpe ORDER BY cFecha, cCodOpe"
If rsCOM.State = adStateOpen Then rsCOM.Close
Set rsCOM = oCon.CargaRecordSet(sql)

lsNumPag = ""
lnLinPag = 0
lnCarLin = 130
lsTitRp1 = "MOVIMIENTO CUENTAS 191503 Y 282503 POR COMPRA VENTA DE DOLARES US."
lsTitRp2 = "MES DE " & UCase(Format$(pdFecha, "mmmm yyyy"))
lnCntPag = 1
lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
lsRTFExt = lsRTFExt & CabeRepo("", "", lnCarLin, "CONTABILIDAD", lsTitRp1, lsTitRp2, "MONEDA EXTRANJERA", lsNumPag) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & "                                191503                                      282503" & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & "FECHA               DEBE         HABER         SALDO  |          DEBE         HABER         SALDO  |        COMPRA         VENTA" & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
ln1915D = 0: ln1915H = 0: ln2825D = 0: ln2825H = 0: lnCompra = 0: lnVenta = 0
If rsCOM.BOF And rsCOM.EOF Then
    MsgBox "No se encontraron registros para este reporte.", vbInformation, "Aviso"
    rsCOM.Close
    Set rsCOM = Nothing
    GetGanCompVenta = ""
    Exit Function
End If
With rsCOM
    lsFecha = ""
    Do While Not .EOF
        If lsFecha <> !cFecha Then
            If lsFecha <> "" Then
                lsRTFExt = lsRTFExt & lsFecha & "  "
                lsRTFExt = lsRTFExt & ls1915D & "  " & ls1915H & "  " & JDNum(Str(CCur(ls1915D) - CCur(ls1915H)), 12, True, 9, 2) & "  |  " & ls2825D & "  " & ls2825H & "  " & JDNum(Str(CCur(ls2825H) - CCur(ls2825D)), 12, True, 9, 2) & "  |  " & lsCompra & "  " & lsVenta
                lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
            End If
            ls1915H = "0.00"
            ls2825H = "0.00"
            ls1915D = "0.00"
            ls2825D = "0.00"
            lsFecha = !cFecha
            Select Case !cCodOpe
                Case gsCVComME
                    ln1915D = ln1915D + !nMontoCV
                    ln2825H = ln2825H + !nMontoF
                    lnCompra = lnCompra + !nMonto
                    ls1915D = JDNum(Str(!nMontoCV), 12, True, 9, 2)
                    ls2825H = JDNum(Str(!nMontoF), 12, True, 9, 2)
                    lsCompra = JDNum(Str(!nMonto), 12, True, 9, 2)
                Case gsCVVenME
                    ln1915H = ln1915H + !nMontoCV
                    ln2825D = ln2825D + !nMontoF
                    lnVenta = lnVenta + !nMonto
                    ls1915H = JDNum(Str(!nMontoCV), 12, True, 9, 2)
                    ls2825D = JDNum(Str(!nMontoF), 12, True, 9, 2)
                    lsVenta = JDNum(Str(!nMonto), 12, True, 9, 2)
            End Select
        Else
            Select Case !cCodOpe
                Case gsCVComME
                    ln1915D = ln1915D + !nMontoCV
                    ln2825H = ln2825H + !nMontoF
                    lnCompra = lnCompra + !nMonto
                    ls1915D = JDNum(Str(!nMontoCV), 12, True, 9, 2)
                    ls2825H = JDNum(Str(!nMontoF), 12, True, 9, 2)
                    lsCompra = JDNum(Str(!nMonto), 12, True, 9, 2)
                Case gsCVVenME
                    ln1915H = ln1915H + !nMontoCV
                    ln2825D = ln2825D + !nMontoF
                    lnVenta = lnVenta + !nMonto
                    ls1915H = JDNum(Str(!nMontoCV), 12, True, 9, 2)
                    ls2825D = JDNum(Str(!nMontoF), 12, True, 9, 2)
                    lsVenta = JDNum(Str(!nMonto), 12, True, 9, 2)
            End Select

        End If
        .MoveNext
    Loop
End With
rsCOM.Close
Set rsCOM = Nothing
lsRTFExt = lsRTFExt & lsFecha & "  "
lsRTFExt = lsRTFExt & ls1915D & "  " & ls1915H & "  " & JDNum(Str(CCur(ls1915D) - CCur(ls1915H)), 12, True, 9, 2) & "  |  " & ls2825D & "  " & ls2825H & "  " & JDNum(Str(CCur(ls2825H) - CCur(ls2825D)), 12, True, 9, 2) & "  |  " & lsCompra & "  " & lsVenta
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & "TOTAL       "
lsRTFExt = lsRTFExt & JDNum(Str(ln1915D), 12, True, 9, 2) & "  " & JDNum(Str(ln1915H), 12, True, 9, 2) & "  " & JDNum(Str(ln1915D - ln1915H), 12, True, 9, 2) & "  |  " & JDNum(Str(ln2825D), 12, True, 9, 2) & "  " & JDNum(Str(ln2825H), 12, True, 9, 2) & "  " & JDNum(Str(ln2825H - ln2825D), 12, True, 9, 2) & "  |  " & JDNum(Str(lnCompra), 12, True, 9, 2) & "  " & JDNum(Str(lnVenta), 12, True, 9, 2)
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(50) & "282503      " & JDNum(Str(ln2825H - ln2825D), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(50) & "191503      " & JDNum(Str(ln1915D - ln1915H), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(62) & JDNum(Str((ln2825H - ln2825D) - (ln1915D - ln1915H)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea

lsRTFExt = lsRTFExt & Space(30) & "NIVELACION DE CAMBIO" & oImpresora.gPrnSaltoLinea
If ln2825H > ln2825D Then
    lsRTFExt = lsRTFExt & Space(30) & "Tipo de Cambio Compra        " & JDNum(Str(nTCC), 12, True, 8, 3) & oImpresora.gPrnSaltoLinea
    lnAjuste = Abs(ln2825H - ln2825D) * nTCC / nTCF
Else
    lsRTFExt = lsRTFExt & Space(30) & "Tipo de Cambio Venta         " & JDNum(Str(nTCV), 12, True, 8, 3) & oImpresora.gPrnSaltoLinea
    lnAjuste = Abs(ln2825H - ln2825D) * nTCV / nTCF
End If
lsRTFExt = lsRTFExt & Space(30) & "Tipo de Cambio Fijo          " & JDNum(Str(nTCF), 12, True, 8, 3) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(30) & "APLICANDO TIPOS DE CAMBIO" & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(30) & "Saldo de Cuenta 282503 $         " & JDNum(Str((ln2825H - ln2825D) / nTCF), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(30) & "Saldo de Cuenta 191503 $         " & JDNum(Str(lnAjuste), 12, True, 9, 2) & "  (b)" & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(30) & "Ganancia (b)-(a)                 " & JDNum(Str(lnAjuste - (ln1915D - ln1915H)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & oImpresora.gPrnSaltoLinea
lsRTFExt = lsRTFExt & Space(30) & "ASIENTO CONTABLE" & oImpresora.gPrnSaltoLinea
If (lnAjuste - (ln1915D - ln1915H)) > 0 Then
    lsRTFExt = lsRTFExt & Space(40) & "191503       " & JDNum(Str(lnAjuste - (ln1915D - ln1915H)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & Space(40) & "5321                     " & JDNum(Str(lnAjuste - (ln1915D - ln1915H)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
Else
    lsRTFExt = lsRTFExt & Space(40) & "432109       " & JDNum(Str(lnAjuste - (ln1915D - ln1915H)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
    lsRTFExt = lsRTFExt & Space(40) & "191503                   " & JDNum(Str(lnAjuste - (ln1915D - ln1915H)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
End If
GetGanCompVenta = lsRTFExt
End Function

Private Function DepositosEscalaMontos(ByVal nTCF As Double) As String
Dim rsCta As ADODB.Recordset
Dim sCad As String
Dim lnCarLin As Integer
Dim lsTitRp1 As String, lsTitRp2 As String
Dim lsMoneda As String, lsNumPag As String
Dim lsFecha As String
Dim lnLinPag As Integer
Dim lnCntPag As Integer
Dim nFSD As Double
Dim sEscala(1 To 14) As String * 50
Dim sProd(1 To 4) As String * 50
Dim I As Integer, P As Integer, J As Integer
Dim nIni As Double, nFin As Double
Dim sRan As String, sTipo As String
Dim nNum01 As Long, nNum02 As Long, nNum03 As Long
Dim nSald01 As Double, nSald02 As Double, nSald03 As Double
Dim nNumSubTot01 As Long, nNumSubTot02 As Long, nNumSubTot03 As Long
Dim nSubTot01 As Double, nSubTot02 As Double, nSubTot03 As Double
Dim nNumTot01 As Long, nNumTot02 As Long, nNumTot03 As Long
Dim nTot01 As Double, nTot02 As Double, nTot03 As Double
Dim dFecha As Date
Dim oGen As DGeneral
Set oGen = New DGeneral

Dim sql As String
Dim oConAdd As DConecta
Set oConAdd = New DConecta
oConAdd.AbreConexion 'Remota "07", , , "01"
oConConsol.AbreConexion 'Remota "07", , , "03"


sEscala(1) = "Hasta el 10% del FSD": sEscala(2) = "Más del 10% al 25% del FSD"
sEscala(3) = "Más del 25% al 50% del FSD": sEscala(4) = "Más del 50% al 100% del FSD"
sEscala(5) = "Más de 1 vez a 2 Veces del FSD": sEscala(6) = "Más de 2 veces a 4 Veces del FSD"
sEscala(7) = "Más de 4 veces a 6 Veces del FSD": sEscala(8) = "Más de 6 veces a 8 Veces del FSD"
sEscala(9) = "Más de 8 veces a 10 Veces del FSD": sEscala(10) = "Más de 10 veces a 15 Veces del FSD"
sEscala(11) = "Más de 15 veces a 20 Veces del FSD": sEscala(12) = "Más de 20 veces a 50 Veces del FSD"
sEscala(13) = "Más de 50 veces a 100 Veces del FSD": sEscala(14) = "Más de 100 Veces del FSD"

sProd(1) = "Sub Total AHORROS": sProd(2) = "Sub Total PLAZO FIJO"
sProd(3) = "Sub Total CTS": sProd(4) = "T O T A L "

nFSD = oGen.GetParametro(2000, 2019)    'Obtiene el Monto del FSD de la tabla de parámetros


Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
'Determina el servido consolidado de captaciones
sql = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont')) " _
    & "OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
Set rsCta = oConAdd.CargaRecordSet(sql)


Set rsCta.ActiveConnection = Nothing
Do While Not rsCta.EOF
    If Trim(rsCta!cNomVar) = "dFecCierreMes" Then
        dFecha = CDate(Trim(rsCta!cValorVar))
    End If
    rsCta.MoveNext
Loop
rsCta.Close

'Imprime la cabecera del reporte
sCad = ""
lnCarLin = 138
lnLinPag = 65
lsMoneda = String(7, " ")
lsTitRp1 = "D E P O S I T O S   S E G Ú N   E S C A L A   D E   M O N T O S"
lsTitRp2 = "AL " & UCase(ArmaFecha(dFecha)) & "  ( EN NUEVOS SOLES )"
lsNumPag = ""
lnCntPag = 0
lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
sCad = sCad & CabeRepo("", "", lnCarLin, "CONTABILIDAD", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
sCad = sCad & String(80, " ") & "Monto FSD: " & JDNum(Trim(Str(nFSD)), 12, True, 9, 2) & String(4, " ") & "Tipo de Cambio Fijo:  " & JDNum(Trim(Str(nTCF)), 7, False, 3, 4) & oImpresora.gPrnSaltoLinea
sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
sCad = sCad & String(81, " ") & "PERSONAS JURIDICAS" & oImpresora.gPrnSaltoLinea
sCad = sCad & String(11, " ") + "ESCALA EN NUEVOS SOLES" & String(18, " ")
sCad = sCad & "  PERSONAS NATURALES "
sCad = sCad & "   SIN FINES DE LUCRO "
sCad = sCad & "   CON FINES DE LUCRO "
sCad = sCad & "  T  O  T  A  L  E  S" & oImpresora.gPrnSaltoLinea
sCad = sCad & String(51, " ")
sCad = sCad & " NUMERO " & "   M O N T O "
sCad = sCad & "  NUMERO " & "    M O N T O "
sCad = sCad & " NUMERO " & "    M O N T O "
sCad = sCad & " NUMERO " & "    M O N T O" & oImpresora.gPrnSaltoLinea
sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea

nNumTot01 = 0: nNumTot02 = 0: nNumTot03 = 0
nTot01 = 0: nTot02 = 0: nTot03 = 0
For P = 1 To 3
    nNumSubTot01 = 0: nNumSubTot02 = 0: nNumSubTot03 = 0
    nSubTot01 = 0: nSubTot02 = 0: nSubTot03 = 0
    Select Case P
        Case 1
            sql = "SELECT TB.Tipo, TB.cTipPers, ISNULL(Count(*),0) nNum, Round(ISNULL(SUM(nSaldo),0),2) Saldo FROM ( "
            sql = sql & "SELECT P.cCodPers, P.cTipPers, Tipo = CASE "
            sql = sql & "WHEN SUM(TA.nSaldo) <= " & nFSD & "*0.10 THEN '01. Hasta el 10% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.10 And SUM(TA.nSaldo) <= " & nFSD & "*0.25 THEN '02. Hasta el 25% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.25 And SUM(TA.nSaldo) <= " & nFSD & "*0.50 THEN '03. Hasta el 50% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.50 And SUM(TA.nSaldo) <= " & nFSD & "*1.00 THEN '04. Hasta el 100% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*1 And SUM(TA.nSaldo) <= " & nFSD & "*2 THEN '05. 01 a 02 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*2 And SUM(TA.nSaldo) <= " & nFSD & "*4 THEN '06. 02 a 04 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*4 And SUM(TA.nSaldo) <= " & nFSD & "*6 THEN '07. 04 a 06 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*6 And SUM(TA.nSaldo) <= " & nFSD & "*8 THEN '08. 06 a 08 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*8 And SUM(TA.nSaldo) <= " & nFSD & "*10 THEN '09. 08 a 10 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*10 And SUM(TA.nSaldo) <= " & nFSD & "*15 THEN '10. 10 a 15 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*15 And SUM(TA.nSaldo) <= " & nFSD & "*20 THEN '11. 15 a 20 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*20 And SUM(TA.nSaldo) <= " & nFSD & "*50 THEN '12. 20 a 50 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*50 And SUM(TA.nSaldo) <= " & nFSD & "*100 THEN '13. 50 a 100 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*100 THEN '14. Más de 100 Veces' "
            sql = sql & "END, ISNULL(SUM(TA.nSaldo),0) nSaldo FROM ( "
            sql = sql & "SELECT T.cCodCta, ISNULL(nSaldo/Count(*),0) nSaldo FROM ( "
            sql = sql & "SELECT PC.cCodCta, PC.cCodPers, nSaldo = CASE "
            sql = sql & "WHEN SUBSTRING(PC.cCodCta,6,1) = '1' THEN A.nSaldCntAC "
            sql = sql & "WHEN SUBSTRING(PC.cCodCta,6,1) = '2' THEN A.nSaldCntAC*" & nTCF & " END "
            sql = sql & "FROM AhorroCConsol A INNER JOIN PersCuentaConsol PC ON A.cCodCta = PC.cCodCta "
            sql = sql & "WHERE A.cEstCtaAC NOT IN ('C', 'U') AND PC.cRelaCta = 'TI' "
            sql = sql & ") T Group by T.cCodCta, T.nSaldo "
            sql = sql & ") TA INNER JOIN PersCuentaConsol PC INNER JOIN DBPersona..Persona P ON PC.cCodPers = P.cCodPers ON "
            sql = sql & "TA.cCodCta = PC.cCodCta WHERE PC.cRelaCta = 'TI' Group by P.cCodPers, P.cTipPers "
            sql = sql & ") TB Group by TB.Tipo, TB.cTipPers Order by TB.Tipo, TB.cTipPers"
        Case 2
            sql = "SELECT TB.Tipo, TB.cTipPers, ISNULL(Count(*),0) nNum, Round(ISNULL(SUM(nSaldo),0),2) Saldo FROM ( "
            sql = sql & "SELECT P.cCodPers, P.cTipPers, Tipo = CASE "
            sql = sql & "WHEN SUM(TA.nSaldo) <= " & nFSD & "*0.10 THEN '01. Hasta el 10% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.10 And SUM(TA.nSaldo) <= " & nFSD & "*0.25 THEN '02. Hasta el 25% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.25 And SUM(TA.nSaldo) <= " & nFSD & "*0.50 THEN '03. Hasta el 50% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.50 And SUM(TA.nSaldo) <= " & nFSD & "*1.00 THEN '04. Hasta el 100% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*1 And SUM(TA.nSaldo) <= " & nFSD & "*2 THEN '05. 01 a 02 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*2 And SUM(TA.nSaldo) <= " & nFSD & "*4 THEN '06. 02 a 04 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*4 And SUM(TA.nSaldo) <= " & nFSD & "*6 THEN '07. 04 a 06 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*6 And SUM(TA.nSaldo) <= " & nFSD & "*8 THEN '08. 06 a 08 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*8 And SUM(TA.nSaldo) <= " & nFSD & "*10 THEN '09. 08 a 10 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*10 And SUM(TA.nSaldo) <= " & nFSD & "*15 THEN '10. 10 a 15 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*15 And SUM(TA.nSaldo) <= " & nFSD & "*20 THEN '11. 15 a 20 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*20 And SUM(TA.nSaldo) <= " & nFSD & "*50 THEN '12. 20 a 50 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*50 And SUM(TA.nSaldo) <= " & nFSD & "*100 THEN '13. 50 a 100 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*100 THEN '14. Más de 100 Veces' "
            sql = sql & "END, ISNULL(SUM(TA.nSaldo),0) nSaldo FROM ( "
            sql = sql & "SELECT T.cCodCta, ISNULL(nSaldo/Count(*),0) nSaldo FROM ( "
            sql = sql & "SELECT PC.cCodCta, PC.cCodPers, nSaldo = CASE "
            sql = sql & "WHEN SUBSTRING(PC.cCodCta,6,1) = '1' THEN A.nSaldCntPF "
            sql = sql & "WHEN SUBSTRING(PC.cCodCta,6,1) = '2' THEN A.nSaldCntPF*" & nTCF & " END "
            sql = sql & "FROM plazofijoconsol A INNER JOIN PersCuentaConsol PC ON A.cCodCta = PC.cCodCta "
            sql = sql & "WHERE A.cEstCtaPF NOT IN ('C', 'U') AND PC.cRelaCta = 'TI' "
            sql = sql & ") T Group by T.cCodCta, T.nSaldo "
            sql = sql & ") TA INNER JOIN PersCuentaConsol PC INNER JOIN DBPersona..Persona P ON PC.cCodPers = P.cCodPers ON "
            sql = sql & "TA.cCodCta = PC.cCodCta WHERE PC.cRelaCta = 'TI' Group by P.cCodPers, P.cTipPers "
            sql = sql & ") TB Group by TB.Tipo, TB.cTipPers Order by TB.Tipo, TB.cTipPers"
        Case 3
            sql = "SELECT TB.Tipo, TB.cTipPers, ISNULL(Count(*),0) nNum, Round(ISNULL(SUM(nSaldo),0),2) Saldo FROM ( "
            sql = sql & "SELECT P.cCodPers, P.cTipPers, Tipo = CASE "
            sql = sql & "WHEN SUM(TA.nSaldo) <= " & nFSD & "*0.10 THEN '01. Hasta el 10% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.10 And SUM(TA.nSaldo) <= " & nFSD & "*0.25 THEN '02. Hasta el 25% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.25 And SUM(TA.nSaldo) <= " & nFSD & "*0.50 THEN '03. Hasta el 50% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*0.50 And SUM(TA.nSaldo) <= " & nFSD & "*1.00 THEN '04. Hasta el 100% del FSD' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*1 And SUM(TA.nSaldo) <= " & nFSD & "*2 THEN '05. 01 a 02 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*2 And SUM(TA.nSaldo) <= " & nFSD & "*4 THEN '06. 02 a 04 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*4 And SUM(TA.nSaldo) <= " & nFSD & "*6 THEN '07. 04 a 06 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*6 And SUM(TA.nSaldo) <= " & nFSD & "*8 THEN '08. 06 a 08 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*8 And SUM(TA.nSaldo) <= " & nFSD & "*10 THEN '09. 08 a 10 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*10 And SUM(TA.nSaldo) <= " & nFSD & "*15 THEN '10. 10 a 15 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*15 And SUM(TA.nSaldo) <= " & nFSD & "*20 THEN '11. 15 a 20 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*20 And SUM(TA.nSaldo) <= " & nFSD & "*50 THEN '12. 20 a 50 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*50 And SUM(TA.nSaldo) <= " & nFSD & "*100 THEN '13. 50 a 100 Veces' "
            sql = sql & "WHEN SUM(TA.nSaldo) > " & nFSD & "*100 THEN '14. Más de 100 Veces' "
            sql = sql & "END, ISNULL(SUM(TA.nSaldo),0) nSaldo FROM ( "
            sql = sql & "SELECT T.cCodCta, ISNULL(nSaldo/Count(*),0) nSaldo FROM ( "
            sql = sql & "SELECT PC.cCodCta, PC.cCodPers, nSaldo = CASE "
            sql = sql & "WHEN SUBSTRING(PC.cCodCta,6,1) = '1' THEN A.nSaldCntCTS "
            sql = sql & "WHEN SUBSTRING(PC.cCodCta,6,1) = '2' THEN A.nSaldCntCTS*" & nTCF & " END "
            sql = sql & "FROM CTSConsol A INNER JOIN PersCuentaConsol PC ON A.cCodCta = PC.cCodCta "
            sql = sql & "WHERE A.cEstCtaCTS NOT IN ('C', 'U') AND PC.cRelaCta = 'TI' "
            sql = sql & ") T Group by T.cCodCta, T.nSaldo "
            sql = sql & ") TA INNER JOIN PersCuentaConsol PC INNER JOIN DBPersona..Persona P ON PC.cCodPers = P.cCodPers ON "
            sql = sql & "TA.cCodCta = PC.cCodCta WHERE PC.cRelaCta = 'TI' Group by P.cCodPers, P.cTipPers "
            sql = sql & ") TB Group by TB.Tipo, TB.cTipPers Order by TB.Tipo, TB.cTipPers"
    End Select
    If rsCta.State = adStateOpen Then rsCta.Close
    Set rsCta = oConConsol.CargaRecordSet(sql)
    Set rsCta.ActiveConnection = Nothing

    nNum01 = 0: nNum02 = 0: nNum03 = 0
    nSald01 = 0: nSald02 = 0: nSald03 = 0
    sRan = "": I = 0
    If rsCta.EOF And rsCta.BOF Then
    Else
        Do While Not rsCta.EOF
            sTipo = Left(rsCta!Tipo, 2)
            If sRan <> sTipo Then
                If sRan <> "" Then
                    sCad = sCad & sEscala(I) & " "
                    sCad = sCad & JDNum(Str(nNum01), 7, True, 7, 0) & " "
                    sCad = sCad & JDNum(Str(nSald01), 13, True, 10, 2) & " "
                    sCad = sCad & JDNum(Str(nNum02), 7, True, 7, 0) & " "
                    sCad = sCad & JDNum(Str(nSald02), 13, True, 10, 2) & " "
                    sCad = sCad & JDNum(Str(nNum03), 7, True, 7, 0) & " "
                    sCad = sCad & JDNum(Str(nSald03), 13, True, 10, 2) & " "
                    sCad = sCad & JDNum(Str(nNum01 + nNum02 + nNum03), 7, True, 7, 0) & " "
                    sCad = sCad & JDNum(Str(nSald01 + nSald02 + nSald03), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
                    nNumSubTot01 = nNumSubTot01 + nNum01
                    nNumSubTot02 = nNumSubTot02 + nNum02
                    nNumSubTot03 = nNumSubTot03 + nNum03
                    nSubTot01 = nSubTot01 + nSald01
                    nSubTot02 = nSubTot02 + nSald02
                    nSubTot03 = nSubTot03 + nSald03
                    If CInt(sTipo) - 1 > I Then
                        For J = I + 1 To CInt(sTipo) - 1
                            sCad = sCad & sEscala(J) & " "
                            sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
                            sCad = sCad & JDNum("0", 13, True, 10, 2) & " "
                            sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
                            sCad = sCad & JDNum("0", 13, True, 10, 2) & " "
                            sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
                            sCad = sCad & JDNum("0", 13, True, 10, 2) & " "
                            sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
                            sCad = sCad & JDNum("0", 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
                        Next J
                    End If
                End If
                nNum01 = 0: nNum02 = 0: nNum03 = 0
                nSald01 = 0: nSald02 = 0: nSald03 = 0
                I = CInt(sTipo)
                sRan = sTipo
            End If
            Select Case rsCta!cTipPers
                Case "1"
                    nNum01 = nNum01 + rsCta!nNum
                    nSald01 = nSald01 + rsCta!Saldo
                Case "2"
                    nNum02 = nNum02 + rsCta!nNum
                    nSald02 = nSald02 + rsCta!Saldo
                Case Else
                    nNum03 = nNum03 + rsCta!nNum
                    nSald03 = nSald03 + rsCta!Saldo
            End Select
            rsCta.MoveNext
        Loop
    End If
    sCad = sCad & sEscala(I) & " "
    sCad = sCad & JDNum(Str(nNum01), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSald01), 13, True, 10, 2) & " "
    sCad = sCad & JDNum(Str(nNum02), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSald02), 13, True, 10, 2) & " "
    sCad = sCad & JDNum(Str(nNum03), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSald03), 13, True, 10, 2) & " "
    sCad = sCad & JDNum(Str(nNum01 + nNum02 + nNum03), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSald01 + nSald02 + nSald03), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    nNumSubTot01 = nNumSubTot01 + nNum01
    nNumSubTot02 = nNumSubTot02 + nNum02
    nNumSubTot03 = nNumSubTot03 + nNum03
    nSubTot01 = nSubTot01 + nSald01
    nSubTot02 = nSubTot02 + nSald02
    nSubTot03 = nSubTot03 + nSald03
    For J = I + 1 To 14
        sCad = sCad & sEscala(J) & " "
        sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
        sCad = sCad & JDNum("0", 13, True, 10, 2) & " "
        sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
        sCad = sCad & JDNum("0", 13, True, 10, 2) & " "
        sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
        sCad = sCad & JDNum("0", 13, True, 10, 2) & " "
        sCad = sCad & JDNum("0", 7, True, 7, 0) & " "
        sCad = sCad & JDNum("0", 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    Next J
    sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    sCad = sCad & sProd(P) & " "
    sCad = sCad & JDNum(Str(nNumSubTot01), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSubTot01), 13, True, 10, 2) & " "
    sCad = sCad & JDNum(Str(nNumSubTot02), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSubTot02), 13, True, 10, 2) & " "
    sCad = sCad & JDNum(Str(nNumSubTot03), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSubTot03), 13, True, 10, 2) & " "
    sCad = sCad & JDNum(Str(nNumSubTot01 + nNumSubTot02 + nNumSubTot03), 7, True, 7, 0) & " "
    sCad = sCad & JDNum(Str(nSubTot01 + nSubTot02 + nSubTot03), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
    sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
    nNumTot01 = nNumTot01 + nNumSubTot01
    nNumTot02 = nNumTot02 + nNumSubTot02
    nNumTot03 = nNumTot03 + nNumSubTot03
    nTot01 = nTot01 + nSubTot01
    nTot02 = nTot02 + nSubTot02
    nTot03 = nTot03 + nSubTot03
Next P
sCad = sCad & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea
sCad = sCad & sProd(P) & " "
sCad = sCad & JDNum(Str(nNumTot01), 7, True, 7, 0) & " "
sCad = sCad & JDNum(Str(nTot01), 13, True, 10, 2) & " "
sCad = sCad & JDNum(Str(nNumTot02), 7, True, 7, 0) & " "
sCad = sCad & JDNum(Str(nTot02), 13, True, 10, 2) & " "
sCad = sCad & JDNum(Str(nNumTot03), 7, True, 7, 0) & " "
sCad = sCad & JDNum(Str(nTot03), 13, True, 10, 2) & " "
sCad = sCad & JDNum(Str(nNumTot01 + nNumTot02 + nNumTot03), 7, True, 7, 0) & " "
sCad = sCad & JDNum(Str(nTot01 + nTot02 + nTot03), 13, True, 10, 2) & oImpresora.gPrnSaltoLinea
sCad = sCad & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea
DepositosEscalaMontos = sCad
End Function

Public Function GeneraFSD(ByVal nFSD As Double, ByVal nTCF As Double) As String
    Dim sCad As String
    Dim rsQryFSD As ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Dim lnCarLin As Integer, lsNumPag As String
    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lnLinPag As Integer, lnCntPag As Integer
    Dim lnCntCta As Integer, lnMonFSD As Double
    Dim rsQrySdo As ADODB.Recordset
    Dim lsCodCta As String
    Dim dFecha As Date
    Dim sPers As String, sNombre As String * 60
    Dim lsMoneda As String
    Dim sMoneda As String
    sCad = ""
    Set rsCta = New ADODB.Recordset
    Dim oConAdd As DConecta
    Set oConAdd = New DConecta
    Dim sql As String
    Dim sColumna(1 To 4) As String * 30
    Dim sNumero As String * 12
    Dim sMontoSaldo As String * 15, sMontoInt As String * 15
    
    oConAdd.AbreConexion 'Remota "07", , , "01"
    oConConsol.AbreConexion 'Remota "07", , , "03"
        
    rsCta.CursorLocation = adUseClient
    'Determina el servidor consolidado de captaciones
    sql = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont')) " _
        & "OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
    Set rsCta = oConAdd.CargaRecordSet(sql)
    Set rsCta.ActiveConnection = Nothing
    
    Do While Not rsCta.EOF
        If Trim(rsCta!cNomVar) = "dFecCierreMes" Then
            dFecha = CDate(Trim(rsCta!cValorVar))
        End If
        rsCta.MoveNext
    Loop
    rsCta.Close

'    sql = "SELECT F.cCodPers, F.cNomPers, (F.nDisPers+F.nIntPers) nSaldo, FD.cCodCta, FD.bInactiva, " _
'        & "FD.nSdoDis, FD.nIntMes FROM FSDVIEW F INNER JOIN FSDVIEWD FD ON F.cCodPers = FD.cCodPers " _
'        & "WHERE (F.nDisPers+F.nIntPers) > " & nFSD & " And F.cCodPers NOT IN (Select cCodPers FROM " _
'        & "vFuncionarios Group by cCodPers) ORDER BY nSaldo DESC, F.cCodPers, FD.cCodCta"
    sql = "SELECT F.cCodPers, F.cNomPers, (F.nDisPers+F.nIntPers) nSaldo, FD.cCodCta, FD.bInactiva, " _
        & "FD.nSdoDis, FD.nIntMes FROM (Select cCodPers From FSDFuncionario Group by cCodPers) U RIGHT JOIN " _
        & "FSDPersona F INNER JOIN FSDPersonaDet FD ON F.cCodPers = FD.cCodPers ON U.cCodPers = F.cCodpers " _
        & "WHERE (F.nDisPers+F.nIntPers) > " & Trim(nFSD) & " And U.cCodPers IS NULL ORDER BY nSaldo DESC, F.cCodPers, FD.cCodCta"
    Set rsCta = oConConsol.CargaRecordSet(sql)
    Set rsCta.ActiveConnection = Nothing
    
    If rsCta.EOF And rsCta.BOF Then
        lnCntPag = lnCntPag + 1
        lsNumPag = ""
        lnLinPag = 65
        lnCarLin = 89
        lsTitRp1 = "L I S T A D O   P A R A   F O N D O   D E   S E G U R O   D E   D E P O S I T O"
        lsTitRp2 = "( AHORROS + PLAZO FIJO + CTS )"
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        sCad = sCad & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(30, " ") & "AL " & Format$(dFecha, "dd mmm yyyy") & String(30, " ") & oImpresora.gPrnSaltoLinea
        sCad = sCad & "TIPO DE CAMBIO FIJO S/. " & JDNum(Trim(Str(nTCF)), 6, False, 1, 4) & String(28, " ")
        sCad = sCad & "FSD A LA FECHA S/. " & JDNum(Trim(Str(nFSD)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(12, " ")
        sCad = sCad & String(20, " ") & "NOMBRE / RAZON SOCIAL" & String(20, " ")
        sCad = sCad & String(16, " ") & oImpresora.gPrnSaltoLinea
        sCad = sCad & "  CODIGO  " & "  "
        sCad = sCad & String(25, " ") & "DEL CLIENTE" & String(25, " ")
        sCad = sCad & "SALDO DISPONIBLE" & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea         ' hasta 160 caracteres x línea
        sCad = sCad & String(36, " ") & "   CUENTA     SALDO/CAPITAL         INTERESES" & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea         ' hasta 160 caracteres x línea
        sCad = sCad & "NO EXISTEN CUENTAS QUE INFORMAR" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina
        GeneraFSD = sCad
    Else
        lsNumPag = ""
        lnLinPag = 65
        lnCarLin = 89
        lsTitRp1 = "L I S T A D O   P A R A   F O N D O   D E   S E G U R O   D E   D E P O S I T O"
        lsTitRp2 = "( AHORROS + PLAZO FIJO + CTS )"
        lnCntPag = 0
        sPers = ""
        Do While Not rsCta.EOF
            If lnLinPag > 55 Then
                If lsNumPag <> "" Then
                    sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina
                End If
                lnLinPag = 0
                lnCntPag = lnCntPag + 1
                lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
                sCad = sCad & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, "SOLES", lsNumPag) & oImpresora.gPrnSaltoLinea
                sCad = sCad & String(30, " ") & "AL " & Format$(dFecha, "dd mmm yyyy") & String(30, " ") & oImpresora.gPrnSaltoLinea
                sCad = sCad & "TIPO DE CAMBIO FIJO S/. " & JDNum(Trim(Str(nTCF)), 6, False, 1, 4) & String(28, " ")
                sCad = sCad & "FSD A LA FECHA S/. " & JDNum(Trim(Str(nFSD)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
                sCad = sCad & String(12, " ")
                sCad = sCad & String(20, " ") & "NOMBRE / RAZON SOCIAL" & String(20, " ")
                sCad = sCad & String(16, " ") & oImpresora.gPrnSaltoLinea
                sCad = sCad & "  CODIGO  " & "  "
                sCad = sCad & String(25, " ") & "DEL CLIENTE" & String(25, " ")
                sCad = sCad & "SALDO DISPONIBLE" & oImpresora.gPrnSaltoLinea
                sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea         ' hasta 160 caracteres x línea
                sCad = sCad & String(36, " ") & "   CUENTA     SALDO/CAPITAL      INTERESES" & oImpresora.gPrnSaltoLinea
                sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea         ' hasta 160 caracteres x línea
                lnLinPag = 14
            End If
            If sPers <> rsCta!cCodPers Then
                sPers = rsCta!cCodPers
                sCad = sCad & rsCta!cCodPers & "  "
                sNombre = PstaNombre(rsCta!cNomPers, False)
                sCad = sCad & sNombre & "  "
                sCad = sCad & JDNum(Trim(Str(rsCta!nSaldo)), 15, True, 12, 2) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                If rsCta("bInactiva") = 0 Then
                    sCad = sCad & String(36, " ") & rsCta!cCodCta & "   " & JDNum(Trim(Str(rsCta!nSdoDis)), 12, True, 9, 2) & "   " & JDNum(Trim(Str(rsCta!nIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                Else
                    sCad = sCad & String(34, " ") & "* " & rsCta!cCodCta & "   " & JDNum(Trim(Str(rsCta!nSdoDis)), 12, True, 9, 2) & "   " & JDNum(Trim(Str(rsCta!nIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                End If
                lnLinPag = lnLinPag + 3
            Else
                If rsCta("bInactiva") = 0 Then
                    sCad = sCad & String(36, " ") & rsCta!cCodCta & "   " & JDNum(Trim(Str(rsCta!nSdoDis)), 12, True, 9, 2) & "   " & JDNum(Trim(Str(rsCta!nIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                Else
                    sCad = sCad & String(34, " ") & "* " & rsCta!cCodCta & "   " & JDNum(Trim(Str(rsCta!nSdoDis)), 12, True, 9, 2) & "   " & JDNum(Trim(Str(rsCta!nIntMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
                End If
                lnLinPag = lnLinPag + 1
            End If
            rsCta.MoveNext
        Loop
    End If
    sCad = sCad & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
    
    sql = "SELECT SUBSTRING(FD.cCodCta,3,3) Prod, SUBSTRING(FD.cCodCta,6,1) Mon, FD.bInactiva, SUM(ROUND(FD.nSdoDis,2)) nSaldo, SUM(ROUND(FD.nIntMes,2)) nInteres, COUNT(*) nNum " _
        & "FROM FSDPersonaDet FD INNER JOIN FSDPersona F LEFT JOIN (Select cCodPers From FSDFuncionario Group by cCodPers) U ON F.cCodPers = U.cCodPers " _
        & "ON FD.cCodPers = F.cCodPers WHERE (F.nDisPers+F.nIntPers) > " & nFSD & " And U.cCodPers IS NULL GROUP BY SUBSTRING(FD.cCodCta,3,3), SUBSTRING(FD.cCodCta,6,1), " _
        & "FD.bInactiva ORDER BY Mon, Prod, bInactiva"
    Set rsCta = oConConsol.CargaRecordSet(sql)
    Set rsCta.ActiveConnection = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        lnCntPag = lnCntPag + 1
        lsTitRp1 = "L I S T A D O   P A R A   F O N D O   D E   S E G U R O   D E   D E P O S I T O"
        lsTitRp2 = "( AHORROS + PLAZO FIJO + CTS )"
        lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
        sCad = sCad & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(30, " ") & "AL " & Format$(dFecha, "dd mmm yyyy") & String(30, " ") & oImpresora.gPrnSaltoLinea
        sCad = sCad & "TIPO DE CAMBIO FIJO S/. " & JDNum(Trim(Str(nTCF)), 6, False, 1, 4) & String(28, " ")
        sCad = sCad & "FSD A LA FECHA S/. " & JDNum(Trim(Str(nFSD)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
        sCad = sCad & oImpresora.gPrnSaltoLinea
        
        sMoneda = ""
        Dim I As Integer
        sColumna(1) = "Depósito de Ahorros Activas"
        sColumna(2) = "Depósito de Ahorros Inactivas"
        sColumna(3) = "Depósito Plazo Fijo"
        sColumna(4) = "Depósito CTS"
        
        Do While Not rsCta.EOF
            If sMoneda <> rsCta("Mon") Then
                sCad = sCad & oImpresora.gPrnSaltoLinea
                If rsCta("Mon") = "1" Then
                    sCad = sCad & "MONEDA NACIONAL" & Space(17) & "      NUMERO" & Space(2) & "     SDO. DISP." & Space(2) & "        INTERES" & oImpresora.gPrnSaltoLinea
                Else
                    sCad = sCad & "MONEDA EXTRANJERA" & Space(14) & "      NUMERO" & Space(2) & "     SDO. DISP." & Space(2) & "         INTERES" & oImpresora.gPrnSaltoLinea
                End If
                I = 1
                sMoneda = rsCta("Mon")
            End If
            RSet sNumero = "0"
            RSet sMontoSaldo = "0.00"
            RSet sMontoInt = "0.00"
            
            Select Case I
                Case 1
                    If rsCta("Prod") = Producto.gCapAhorros And rsCta("bInactiva") = "0" Then
                        RSet sNumero = Format$(rsCta("nNum"), "#,##0")
                        RSet sMontoSaldo = Format$(rsCta("nSaldo"), "#,##0.00")
                        RSet sMontoInt = Format$(rsCta("nInteres"), "#,##0.00")
                        rsCta.MoveNext
                    End If
                Case 2
                    If rsCta("Prod") = Producto.gCapAhorros And rsCta("bInactiva") = "1" Then
                        RSet sNumero = Format$(rsCta("nNum"), "#,##0")
                        RSet sMontoSaldo = Format$(rsCta("nSaldo"), "#,##0.00")
                        RSet sMontoInt = Format$(rsCta("nInteres"), "#,##0.00")
                        rsCta.MoveNext
                    End If
                Case 3
                    If rsCta("Prod") = Producto.gCapPlazoFijo Then
                        RSet sNumero = Format$(rsCta("nNum"), "#,##0")
                        RSet sMontoSaldo = Format$(rsCta("nSaldo"), "#,##0.00")
                        RSet sMontoInt = Format$(rsCta("nInteres"), "#,##0.00")
                        rsCta.MoveNext
                    End If
                Case 4
                    If rsCta("Prod") = Producto.gCapCTS Then
                        RSet sNumero = Format$(rsCta("nNum"), "#,##0")
                        RSet sMontoSaldo = Format$(rsCta("nSaldo"), "#,##0.00")
                        RSet sMontoInt = Format$(rsCta("nInteres"), "#,##0.00")
                        rsCta.MoveNext
                    End If
            End Select
            sCad = sCad & sColumna(I) & Space(2) & sNumero & Space(2) & sMontoSaldo & Space(2) & sMontoInt & oImpresora.gPrnSaltoLinea
            I = I + 1
        Loop
        sCad = sCad & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(lnCarLin, "=") & oImpresora.gPrnSaltoLinea
    End If
    GeneraFSD = sCad
    Set rsCta = Nothing
End Function

Public Sub GeneraFSDExcel(ByVal nFSD As Double, nTCF As Double)
    Dim rsProd As ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Dim lsFile As New FileSystemObject
    Dim lsMensaj As String
    Dim lnDiaMes As Integer
    Dim lnDiaDif As Integer
    Dim ldIni As Date
    Dim I As Integer, J As Integer, K As Integer
    Dim nDias As Integer
    Dim bExiste As Boolean
    Dim xlsApp As Excel.Application
    Dim xlsBok As Excel.Workbook
    Dim xlsSht As Excel.Worksheet
    Dim posx As Integer, posy As Integer
    Dim sProd As String
    Dim nSaldo As Currency
    Dim sArchivo As String
    Dim lbPrimerDia As Boolean
    Dim dFecha As Date
    On Error GoTo ErrExcel
    Screen.MousePointer = 11
    Set xlsApp = New Excel.Application
    Set xlsBok = xlsApp.Workbooks.Add
    
    Dim sql As String
    Dim oConAdd As DConecta
    Set oConAdd = New DConecta

    oConAdd.AbreConexion 'Remota "07", , , "01"
    oConConsol.AbreConexion 'Remota "07", , , "03"
    
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    'Determina el servidor consolidado de captaciones
    sql = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont')) " _
        & "OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
    Set rsCta = oConAdd.CargaRecordSet(sql)
    Set rsCta.ActiveConnection = Nothing
    Do While Not rsCta.EOF
        If Trim(rsCta!cNomVar) = "dFecCierreMes" Then
            dFecha = CDate(Trim(rsCta!cValorVar))
        End If
        rsCta.MoveNext
    Loop
    rsCta.Close
    
    For I = 1 To 2
        Set xlsSht = xlsBok.Worksheets.Add
        xlsSht.Name = IIf(I = 1, "Moneda Nacional", "Moneda Extranjera")
        ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
        xlsSht.Cells(1, 1) = "CMACT - Agencia " & gsNomAge
        xlsSht.Cells(2, 1) = gdFecSis
        xlsSht.Cells(2, 2) = Format$(Time, "hh:mm:ss AMPM")
        K = 5
        Do While Month(ldIni) = Month(dFecha)
            xlsSht.Cells(K, 1) = ldIni
            xlsSht.Cells(K + 35, 1) = ldIni
            xlsSht.Cells(K + 70, 1) = ldIni
            K = K + 1
            ldIni = DateAdd("d", 1, ldIni)
        Loop
        ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
        
        Set rsProd = New ADODB.Recordset
        rsProd.CursorLocation = adUseClient
        sql = "SELECT SUBSTRING(B.cCodCta,3,3) Prod, B.cCodCta FROM FSDVIEW A INNER JOIN FSDVIEWD B ON  " _
            & "A.cCodPers = B.cCodPers WHERE (A.nDisPers+A.nIntPers) > " & nFSD & " " _
            & "AND SUBSTRING(B.cCodCta,6,1) = '" & I & "' ORDER BY Prod, B.cCodCta"
        If rsProd.State = adStateOpen Then rsProd.Close
        Set rsProd = oConConsol.CargaRecordSet(sql)
        Set rsProd.ActiveConnection = Nothing
        sProd = ""
        Do While Not rsProd.EOF
            If sProd <> rsProd!Prod Then
                posx = 2
                sProd = rsProd!Prod
            End If
            ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
            sql = "SELECT TA.cCodCta, TA.dFecTran, TA.nSaldCnt from Transaho TA " _
                & "WHERE cCodCta = '" & rsProd!cCodCta & "' AND " _
                & "DATEDIFF(mm,TA.dFecTran,'" & Format$(ldIni, "mm/dd/yyyy") & "')= 0 " _
                & "AND nNumTran IN (SELECT MAX(nNumTran) FROM TransAho TA1 WHERE " _
                & "TA1.cCodCta = TA.cCodCta AND CONVERT(VarChar(10), TA1.dFectran, 103) = " _
                & "CONVERT(VarChar(10), TA.dFectran, 103) group by ccodcta) " _
                & "And TA.cCodOpe NOT IN ('260105') ORDER BY TA.dFecTran"
            If rsCta.State = adStateOpen Then rsCta.Close
            Set rsCta = oConConsol.CargaRecordSet(sql)
            Set rsCta.ActiveConnection = Nothing
            Select Case rsProd!Prod
'                Case gsCodProAC
'                    posy = 5
'                    xlsSht.Cells(4, posx) = rsProd!cCodCta
'                Case gsCodProPF
'                    posy = 40
'                    xlsSht.Cells(39, posx) = rsProd!cCodCta
'                Case gsCodProCTS
'                    posy = 75
'                    xlsSht.Cells(74, posx) = rsProd!cCodCta
            End Select
            nSaldo = GetSaldoIniFSD(rsProd!cCodCta, ldIni, oConConsol)
            lbPrimerDia = False
            Do While Not rsCta.EOF
                nDias = DateDiff("d", ldIni, rsCta!DFECTRAN)
                If Day(rsCta!DFECTRAN) = 1 Then lbPrimerDia = True
                If nDias > 0 Then
                    If Day(ldIni) = 1 And Not lbPrimerDia Then
                        If nSaldo > 0 Then xlsSht.Cells(posy, posx) = Format$(nSaldo, "#,##0.00")
                        posy = posy + 1
                    End If
                    For J = Day(ldIni) + 1 To nDias + Day(ldIni) - 1
                        If nSaldo > 0 Then xlsSht.Cells(posy, posx) = Format$(nSaldo, "#,##0.00")
                        posy = posy + 1
                    Next J
                End If
                If rsCta!nSaldCnt > 0 And Not IsNull(rsCta!nSaldCnt) Then xlsSht.Cells(posy, posx) = Format$(rsCta!nSaldCnt, "#,##0.00")
                posy = posy + 1
                ldIni = rsCta!DFECTRAN
                nSaldo = IIf(IsNull(rsCta!nSaldCnt), 0, rsCta!nSaldCnt)
                rsCta.MoveNext
            Loop
            rsCta.Close
            rsProd.MoveNext
            posx = posx + 1
        Loop
        sArchivo = App.path & "\Spooler\LFSD" & Mid(gsCodAge, 4, 2) & "_" & Format$(dFecha, "yyyymmdd") & ".XLS"
        If I = 1 Then xlsBok.SaveAs sArchivo
        Set xlsSht = Nothing
    Next I
    rsProd.Close
    Set rsProd = Nothing
    Set rsCta = Nothing
    xlsBok.Save
    xlsBok.Close
    Set xlsBok = Nothing
    Set xlsApp = Nothing
    Screen.MousePointer = 0
    MsgBox "Proceso finalizado con éxito. Archivo generado : " & sArchivo, vbInformation, "Aviso"
    Exit Sub
ErrExcel:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbExclamation, "Error"
    Set rsProd = Nothing
    Set rsCta = Nothing
    Set xlsBok = Nothing
    Set xlsApp = Nothing
End Sub

Private Function GetSaldoIniFSD(ByVal psCodCta As String, ByVal pdFecha As Date, ByVal sConn As DConecta) As Double
    Dim rsSald As ADODB.Recordset
    Dim sql As String
    sql = "SELECT TA.cCodCta, TA.dFecTran, ISNULL(TA.nSaldCnt,0) nSaldo from TransahoConsol TA " _
        & "WHERE cCodCta = '" & psCodCta & "' AND " _
        & "DATEDIFF(dd,TA.dFecTran,'" & Format$(DateAdd("d", -1, pdFecha), "mm/dd/yyyy") & "')= 0 " _
        & "AND nNumtran IN (SELECT MAX(nNumtran) FROM TransAhoConsol TA1 WHERE " _
        & "TA1.cCodCta = TA.cCodCta AND DateDiff(dd,TA1.dFectran,TA.dFectran) = 0 " _
        & "Group by ccodcta) And TA.cCodOpe NOT IN ('260105','260106') ORDER BY TA.dFecTran"
    Set rsSald = New ADODB.Recordset
    rsSald.CursorLocation = adUseClient
    Set rsSald = oConConsol.CargaRecordSet(sql)
    rsSald.ActiveConnection = Nothing
    If rsSald.EOF And rsSald.BOF Then
        GetSaldoIniFSD = 0
    Else
        GetSaldoIniFSD = rsSald!nSaldo
    End If
    rsSald.Close
    Set rsSald = Nothing
End Function


Public Sub GeneraFSDExcelFuncionarios(ByVal nFSD As Double, ByVal nTCF As Double)
    Dim rsProd As ADODB.Recordset
    Dim rsCta As ADODB.Recordset
    Dim lsFile As New FileSystemObject
    Dim lsMensaj As String
    Dim lnDiaMes As Integer
    Dim lnDiaDif As Integer
    Dim ldIni As Date
    Dim I As Integer, J As Integer, K As Integer
    Dim nDias As Integer
    Dim bExiste As Boolean
    Dim xlsApp As Excel.Application
    Dim xlsBok As Excel.Workbook
    Dim xlsSht As Excel.Worksheet
    Dim posx As Integer, posy As Integer
    Dim sProd As String
    Dim nSaldo As Currency
    Dim sArchivo As String
    Dim lbPrimerDia As Boolean
    Dim dFecha As Date
    On Error GoTo ErrExcel
    Screen.MousePointer = 11
    Set xlsApp = New Excel.Application
    Set xlsBok = xlsApp.Workbooks.Add
    Dim sql As String
    Dim oConAdd As DConecta
    Set oConAdd = New DConecta
    oConAdd.AbreConexion 'Remota "07", , , "01"
    oConConsol.AbreConexion 'Remota "07", , , "03"
    
    Set rsCta = New ADODB.Recordset
    rsCta.CursorLocation = adUseClient
    'Determina el servidor consolidado de captaciones
    sql = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont')) " _
        & "OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
    Set rsCta = oConAdd.CargaRecordSet(sql)
    Set rsCta.ActiveConnection = Nothing
    
    Do While Not rsCta.EOF
        If Trim(rsCta!cNomVar) = "dFecCierreMes" Then
            dFecha = CDate(Trim(rsCta!cValorVar))
        End If
        rsCta.MoveNext
    Loop
    rsCta.Close


    For I = 1 To 2
        Set xlsSht = xlsBok.Worksheets.Add
        xlsSht.Name = IIf(I = 1, "Moneda Nacional", "Moneda Extranjera")
        ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
        xlsSht.Cells(1, 1) = "CMACT - Agencia " & gsNomAge
        xlsSht.Cells(2, 1) = gdFecSis
        xlsSht.Cells(2, 2) = Format$(Time, "hh:mm:ss AMPM")
        K = 5
        Do While Month(ldIni) = Month(dFecha)
            xlsSht.Cells(K, 1) = ldIni
            xlsSht.Cells(K + 35, 1) = ldIni
            xlsSht.Cells(K + 70, 1) = ldIni
            K = K + 1
            ldIni = DateAdd("d", 1, ldIni)
        Loop
        ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
        Set rsProd = New ADODB.Recordset
        rsProd.CursorLocation = adUseClient
        sql = "SELECT SUBSTRING(cCodCta,3,3) Prod, cCodCta FROM fsdFuncionario  " _
            & "WHERE cCodCta LIKE '_____" & I & "%' ORDER BY Prod, cCodCta"
        If rsProd.State = adStateOpen Then rsProd.Close
        
        Set rsProd = oConConsol.CargaRecordSet(sql)
        
        Set rsProd.ActiveConnection = Nothing
        sProd = ""
        Do While Not rsProd.EOF
            If sProd <> rsProd!Prod Then
                posx = 2
                sProd = rsProd!Prod
            End If
            ldIni = CDate("01" & "/" & Month(dFecha) & "/" & Year(dFecha))
            sql = "SELECT TA.cCodCta, TA.dFecTran, TA.nSaldCnt from TransahoConsol TA " _
                & "WHERE cCodCta = '" & rsProd!cCodCta & "' AND " _
                & "DATEDIFF(mm,TA.dFecTran,'" & Format$(ldIni, "mm/dd/yyyy") & "')= 0 " _
                & "AND nNumTran IN (SELECT MAX(nNumTran) FROM TransAhoConsol TA1 WHERE " _
                & "TA1.cCodCta = TA.cCodCta And DateDiff(dd,TA1.dFectran,TA.dFectran) = 0 " _
                & "And TA1.cCodOpe NOT IN ('260105','260106') Group by cCodCta) ORDER BY TA.dFecTran"
            If rsCta.State = adStateOpen Then rsCta.Close
            Set rsCta = oConConsol.CargaRecordSet(sql)
            Set rsCta.ActiveConnection = Nothing
            Select Case rsProd!Prod
                Case Producto.gCapAhorros
                    posy = 5
                    xlsSht.Cells(4, posx) = rsProd("cCodCta")
                Case Producto.gCapPlazoFijo
                    posy = 40
                    xlsSht.Cells(39, posx) = rsProd("cCodCta")
                Case Producto.gCapCTS
                    posy = 75
                    xlsSht.Cells(74, posx) = rsProd("cCodCta")
            End Select
            nSaldo = GetSaldoIniFSD(rsProd!cCodCta, ldIni, oConConsol)
            lbPrimerDia = False
            Do While Not rsCta.EOF
                nDias = DateDiff("d", ldIni, rsCta("dFecTran"))
                If Day(rsCta("dFecTran")) = 1 Then lbPrimerDia = True
                If nDias > 0 Then
                    If Day(ldIni) = 1 And Not lbPrimerDia Then
                        If nSaldo > 0 Then xlsSht.Cells(posy, posx) = Format$(nSaldo, "#,##0.00")
                        posy = posy + 1
                    End If
                    For J = Day(ldIni) + 1 To nDias + Day(ldIni) - 1
                        If nSaldo > 0 Then xlsSht.Cells(posy, posx) = Format$(nSaldo, "#,##0.00")
                        posy = posy + 1
                    Next J
                End If
                If rsCta!nSaldCnt > 0 And Not IsNull(rsCta!nSaldCnt) Then xlsSht.Cells(posy, posx) = Format$(rsCta!nSaldCnt, "#,##0.00")
                posy = posy + 1
                ldIni = rsCta!DFECTRAN
                nSaldo = IIf(IsNull(rsCta!nSaldCnt), 0, rsCta!nSaldCnt)
                rsCta.MoveNext
            Loop
            rsCta.Close
            rsProd.MoveNext
            posx = posx + 1
        Loop
        sArchivo = App.path & "\Spooler\LFSDFun" & Mid(gsCodAge, 4, 2) & "_" & Format$(dFecha, "yyyymmdd") & ".XLS"
        If I = 1 Then xlsBok.SaveAs sArchivo
        Set xlsSht = Nothing
    Next I
    rsProd.Close
    Set rsProd = Nothing
    Set rsCta = Nothing
    xlsBok.Save
    xlsBok.Close
    Set xlsBok = Nothing
    Set xlsApp = Nothing
    Screen.MousePointer = 0
    MsgBox "Proceso finalizado con éxito. Archivo generado : " & sArchivo, vbInformation, "Aviso"
    Exit Sub
ErrExcel:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbExclamation, "Error"
    Set rsProd = Nothing
    Set rsCta = Nothing
    Set xlsBok = Nothing
    Set xlsApp = Nothing
End Sub


'Public Function GenSdosFM(psTipAho As String, psTipPer As String, psTipMon As String, _
'            ByVal psCodAge As String, ByVal psFchPro As String, Optional pbCtaIna As Byte = 0, _
'            Optional psCodIns As String = "") As String
'Dim lsQrySdo As String, lsTipAho As String
'Dim rsQrySdo As ADODB.Recordset, rsCta As ADODB.Recordset
'Dim lsInstit As String, lsPerson As String, lsMoneda As String
'Dim dbAux As ADODB.Connection
'Dim dFecha As Date
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
''Determina el servido consolidado de captaciones
'Set dbAux = New ADODB.Connection
'sql = "Select cNomVar, cValorVar From VarSistema Where (cCodProd = 'AHO' And cNomVar IN ('cDBAhoCont','cServAhoCont')) " _
'    & "OR (cCodProd = 'ADM' And cNomVar IN ('dFecCierreMes'))"
'rsCta.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'Set rsCta.ActiveConnection = Nothing
'dbAux.Provider = "SQLOLEDB"
'Do While Not rsCta.EOF
'    If Trim(rsCta!cNomVar) = "cDBAhoCont" Then
'        dbAux.Properties("Initial Catalog") = Trim(rsCta!cValorVar)
'    ElseIf Trim(rsCta!cNomVar) = "cServAhoCont" Then
'        dbAux.Properties("Data Source") = Trim(rsCta!cValorVar)
'    ElseIf Trim(rsCta!cNomVar) = "dFecCierreMes" Then
'        dFecha = CDate(Trim(rsCta!cValorVar))
'    End If
'    rsCta.MoveNext
'Loop
'rsCta.Close
'Set rsCta = Nothing
'dbAux.Properties("User ID") = "DBAccess"
'dbAux.Properties("Password") = "cmact"
'dbAux.Open
'
'lsRTFRfm = ""
'
'lsMoneda = IIf(psTipMon = "1", "SOLES", "DOLARES")
'
' Select Case psTipPer
'     Case "1"
'          lsPerson = "PERSONA(S)  NATURAL(ES)" & IIf(psTipAho = gsCodProAC, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
'     Case "2"
'          lsPerson = "PERSONA(S)  JURIDICA(S)  SIN  FINES  DE  LUCRO" & IIf(psTipAho = gsCodProAC, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
'     Case "3"
'          lsPerson = "PERSONA(S)  JURIDICA(S)  CON  FINES  DE  LUCRO" & IIf(psTipAho = gsCodProAC, IIf(pbCtaIna = 0, "  -  ACTIVA", "  -  INACTIVA"), "")
' End Select
'
' lsQrySdo = ""
' Select Case psTipAho
'     Case gsCodProAC
'
'          lsTipAho = "A H O R R O S"
'          lsQrySdo = "SELECT A.cCodCta, A.cOrdPag, A.nSaldDispAC nSdoDis, A.nSaldCntAC nSdoCnt, " _
'                   & "A.nInteres nIntMes, A.nSaldAntAC nSdoAnt, A.dUltCntAC dUltMovAC, P.cNomPers, " _
'                   & "P.cCodPers FROM AhorroC A INNER JOIN PersCuenta PC ON A.cCodCta = PC.cCodCta " _
'                   & "INNER JOIN Persona P ON PC.cCodPers = P.cCodPers " _
'                   & "WHERE A.cEstCtaAC NOT IN ('C','U') AND A.cPersoneria = '" & psTipPer & "' " _
'                   & "AND PC.cRelaCta = 'TI' AND A.bInactiva = " & pbCtaIna & " " _
'                   & "AND A.cCodCta LIKE '" & psCodAge & gsCodProAC & psTipMon & "%' " _
'                   & "ORDER BY A.cOrdPag, A.cCodCta"
'     Case gsCodProPF
'
'          lsTipAho = "P L A Z O   F I J O"
'          lsQrySdo = "SELECT PF.dAperPF, DateAdd(day, PF.nPlazo, PF.dAperPF) AS dFchVct, PF.nPlazo, PF.cCodCta, PE.cNomPers, " _
'                   & "PF.nCapital AS nCapital, PF.nIntPag, PF.nIntDev, PF.nDepoIni, PF.nTasaIntPF, PF.dUltCntPF AS dUltMovPF, PC.cCodPers " _
'                   & "FROM PlazoFijo PF  " _
'                   & "INNER JOIN PersCuenta PC ON PF.cCodCta = PC.cCodCta  " _
'                   & "INNER JOIN Persona PE ON PC.cCodPers = PE.cCodPers " _
'                   & "WHERE PF.cEstCtaPF NOT IN ('C','U') " _
'                   & "AND PF.cPersoneria = '" & psTipPer & "'  AND PC.cRelaCta = 'TI' " _
'                   & "AND SUBSTRING(PF.cCodCta,6,1) = '" & psTipMon & "'  " _
'                   & "ORDER BY PF.cCodCta"
'     Case gsCodProCTS
'          lsTipAho = "C T S"
'          Select Case psCodIns
'             Case gsInstCmac
'                  lsPerson = "E M P L E A D O S    C M A C T"
'                  lsInstit = "LTRIM(RTRIM(CTS.cCodInst)) = '" & gsInstCmac & "' "
'             Case Else
'                  lsPerson = "C L I E N T E S    C M A C T"
'                  lsInstit = "LTRIM(RTRIM(CTS.cCodInst)) <> '" & gsInstCmac & "' "
'          End Select
'          lsQrySdo = "SELECT CTS.dAperCTS, CTS.cCodCta, CTS.nSaldDispCTS As nSdoDis, CTS.nSaldCntCTS AS nSdoCnt, CTS.dUltCntCTS AS dUltMovCTS, " _
'                   & "CTS.nUltDepCTS As nUltDep, CTS.nCapitalCTS As nCapital, CTS.nInteres AS nIntMes, PE.cNomPers, PC.cCodPers " _
'                   & "FROM CTS CTS  " _
'                   & "INNER JOIN PersCuenta PC ON CTS.cCodCta = PC.cCodCta  " _
'                   & "INNER JOIN Persona PE ON PC.cCodPers = PE.cCodPers  " _
'                   & "WHERE CTS.cEstCtaCTS NOT IN ('C','U') AND " & lsInstit & "" _
'                   & "AND PE.cTipPers = '" & psTipPer & "'  AND PC.cRelaCta = 'TI'  " _
'                   & "AND SUBSTRING(CTS.cCodCta,6,1) = '" & psTipMon & "'  " _
'                   & "ORDER BY CTS.cCodCta"
'
' End Select
'
' If lsQrySdo = "" Then
'    Exit Function
' End If
'
' Set rsQrySdo = New ADODB.Recordset
' rsQrySdo.CursorLocation = adUseClient
' dbAux.CommandTimeout = 360
' rsQrySdo.Open lsQrySdo, dbAux, adOpenStatic, adLockReadOnly, adCmdText
' rsQrySdo.ActiveConnection = Nothing
' If rsQrySdo.EOF And rsQrySdo.BOF Then
'    lsRTFRfm = ""
'    lsReporte = ""
'    GenSdosFM = lsReporte
' Else
'    lsRTFRfm = ""
'    lsReporte = ""
'    PrtSdosFM rsQrySdo
'    GenSdosFM = lsReporte
' End If
'rsQrySdo.Close
'Set rsQrySdo = Nothing
'dbAux.Close
'Set dbAux = Nothing
'End Function
'
'Private Sub PrtSdosFM(ByVal prQryDat As ADODB.Recordset)
'
'    Dim lnCarLin As Integer, i As Integer
'    Dim lsTitRp1 As String, lsTitRp2 As String
'    Dim lnTotDep As Currency, lnTotRet As Currency
'    Dim lsCodPro As String, lsNumPag As String
'    Dim lnLinPag As Integer, lnCntPag As Integer
'    Dim lsSubTit As String, lnCntCta As Integer
'    Dim lnSdoCnt As Currency, lnSdoDis As Currency
'    Dim lnSumInt As Currency, lnTotInt As Currency
'    Dim lnCapita As Currency, lsCodCta As String
'    Dim lnCapiFi As Currency, lnSdDiFi As Currency, lnSdCnFi As Currency
'    Dim lnCtaPag As Integer, lnTotCta As Integer
'    Dim lnCntRel As Integer, lbCtaIde As Boolean
'    Dim lnCalInt As Currency, lnNumDia As Integer
'    Dim lnApertu As Currency       '   , lnCapita As Currency
'    Dim lnIntDev As Currency, lnIntCal As Currency
'    Dim lnTotalP As Currency, lnTotApe As Currency
'    Dim lnTotCap As Currency, lnToInDv As Currency
'    Dim lnToInCa As Currency, lnTotalF As Currency
'    Dim lbCieMes As Boolean, lnIntCap As Currency
'
'    If CierreRealizado(2) Then
'       lbCieMes = True
'    Else
'       lbCieMes = False
'    End If
'
'    lsNumPag = ""
'    lnLinPag = 65
'    lnCarLin = 112
'    lsTitRp1 = "L I S T A D O   D E   C U E N T A S   D E   " & lsTipAho
'    lsTitRp2 = ""
'    lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
'    lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
'    lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
'    lnTotApe = 0: lnTotCap = 0: lnToInDv = 0: lnToInCa = 0: lnTotalF = 0
'
'    lnCntPag = 0
'    lbCtaIde = False
'    With prQryDat
'         Do While Not .EOF
'            Select Case Mid$(!cCodCta, 3, 3)
'               Case gsCodProAC
'                    lsTitRp2 = IIf(!cOrdPag = "N", "", " CON ORDEN DE PAGO  -  ") & lsPerson
'               Case gsCodProPF
'                    lsTitRp2 = lsPerson
'                    lnCarLin = 167
'               Case gsCodProCTS
'                    lsTitRp2 = lsPerson
'                    lnCarLin = 173
'            End Select
'
'            If lsTitRp2 <> lsCodPro Then
'               If lsNumPag <> "" Then
'                  If Mid$(!cCodCta, 3, 3) = gsCodProAC Then
'                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'                     lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea             ' hasta 160 caracteres x línea
'                     lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(24, " ")
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
'                     lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
'                     lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
'                  Else
'                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'                     lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea             ' hasta 160 caracteres x línea
'                     lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(53, " ")
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotApe)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCap)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInDv)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInCa)), 12, True, 9, 2) & " "
'                     lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalF)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'                     lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
'                     lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
'                     lnTotApe = 0: lnTotCap = 0: lnToInDv = 0: lnToInCa = 0: lnTotalF = 0
'                  End If
'               End If
'               lsCodPro = lsTitRp2
'               lnLinPag = 65
'            End If
'
'            If lnLinPag > 58 Then
'               If lsNumPag <> "" Then
'                  If lnSdoDis <> 0 Or lnCapita <> 0 Then
'                     If Mid$(!cCodCta, 3, 3) = gsCodProAC Then
'                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'                        lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
'                        lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
'                     Else
'                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'                        lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
'                        lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'                        lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
'                        lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
'                     End If
'                  End If
'               End If
'
'               lnCntPag = lnCntPag + 1
'               lsNumPag = FillNum(Trim(Str(lnCntPag)), 4, " ")
'               lsRTFRfm = lsRTFRfm & CabeRepo("", "", lnCarLin, "SECCION AHORROS", lsTitRp1, lsTitRp2, lsMoneda, lsNumPag) & oImpresora.gPrnSaltoLinea
'               lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea
'
'               If Mid$(!cCodCta, 3, 3) = gsCodProAC Then
'                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
'                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
''                  lsRTFRfm = lsRTFRfm & "  CUENTA  " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "     INTERES  " & " "
'                  lsRTFRfm = lsRTFRfm & "  SALDO    " & " "
'                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
'                  lsRTFRfm = lsRTFRfm & " ULTIMO  " & oImpresora.gPrnSaltoLinea
'                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
'                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
''                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "      DEL CLIENTE     " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "    DEL MES  " & " "
'                  lsRTFRfm = lsRTFRfm & "DISPONIBLE " & " "
'                  lsRTFRfm = lsRTFRfm & "   CONTABLE  " & " "
'                  lsRTFRfm = lsRTFRfm & "MOVIMIENTO" & oImpresora.gPrnSaltoLinea
'               End If
'               If Mid$(!cCodCta, 3, 3) = gsCodProPF Then
''                  lnCarLin = 153
'                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
'                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
'                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
'                  lsRTFRfm = lsRTFRfm & "PLAZOS" & " "
'                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
''                  lsRTFRfm = lsRTFRfm & " CUENTA  " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "      DEPOSITO  " & " "
'                  lsRTFRfm = lsRTFRfm & "            " & " "
'                  lsRTFRfm = lsRTFRfm & "  INTERES   " & " "
'                  lsRTFRfm = lsRTFRfm & "  INTERES   " & " "
'                  lsRTFRfm = lsRTFRfm & "            " & " "
'                  lsRTFRfm = lsRTFRfm & " ULTIMO" & oImpresora.gPrnSaltoLinea
'                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
'                  lsRTFRfm = lsRTFRfm & " APERTURA " & " "
'                  lsRTFRfm = lsRTFRfm & "  VENCIM  " & " "
'                  lsRTFRfm = lsRTFRfm & "(DIAS)" & " "
'                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
''                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "     DEL CLIENTE     " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "    APERTURA  " & " "
'                  lsRTFRfm = lsRTFRfm & "  CAPITAL   " & " "
'                  lsRTFRfm = lsRTFRfm & " DEVENGADO  " & " "
'                  lsRTFRfm = lsRTFRfm & " CALCULADO  " & " "
'                  lsRTFRfm = lsRTFRfm & "  T O T A L " & " "
'                  lsRTFRfm = lsRTFRfm & "CONTACTO" & oImpresora.gPrnSaltoLinea
'               End If
'               If Mid$(!cCodCta, 3, 3) = gsCodProCTS Then
'                  lsRTFRfm = lsRTFRfm & "CORRE-" & " "
'                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
'                  lsRTFRfm = lsRTFRfm & "  FECHAS  " & " "
'                  lsRTFRfm = lsRTFRfm & "PLAZOS" & " "
'                  lsRTFRfm = lsRTFRfm & "   NUMERO    " & String(10, " ")
''                  lsRTFRfm = lsRTFRfm & " CUENTA  " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "NOMBRE / RAZON SOCIAL" & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "   ULTIMO   " & " "
'                  lsRTFRfm = lsRTFRfm & "            " & " "
'                  lsRTFRfm = lsRTFRfm & "   INTERES  " & " "
'                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
'                  lsRTFRfm = lsRTFRfm & "   SALDO    " & " "
'                  lsRTFRfm = lsRTFRfm & "    ULTIMO" & oImpresora.gPrnSaltoLinea
'                  lsRTFRfm = lsRTFRfm & "LATIVO" & " "
'                  lsRTFRfm = lsRTFRfm & " APERTURA " & " "
'                  lsRTFRfm = lsRTFRfm & "  VENCIM  " & " "
'                  lsRTFRfm = lsRTFRfm & "(DIAS)" & " "
'                  lsRTFRfm = lsRTFRfm & "   CUENTA    " & String(10, " ")
''                  lsRTFRfm = lsRTFRfm & "ANTERIOR " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "     DEL CLIENTE     " & String(10, " ")
'                  lsRTFRfm = lsRTFRfm & "  DEPOSITO  " & " "
'                  lsRTFRfm = lsRTFRfm & "  CAPITAL   " & " "
'                  lsRTFRfm = lsRTFRfm & "  DEL MES   " & " "
'                  lsRTFRfm = lsRTFRfm & "  DISPONIBLE" & " "
'                  lsRTFRfm = lsRTFRfm & "  CONTABLE  " & " "
'                  lsRTFRfm = lsRTFRfm & "  CONTACTO" & oImpresora.gPrnSaltoLinea
'               End If
'
'               lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea         ' hasta 160 caracteres x línea
'               lnLinPag = 9
'            End If
'            lnCntRel = 0
'            If !cCodCta = lsCodCta Then
'                If lbCtaIde = False Then
'                    Select Case Mid$(lsCodCta, 3, 3)
'                        Case gsCodProAC
'                            lsRTFRfm = lsRTFRfm & String(20, " ")
'                            lnCntRel = lnCntRel + 1
'                        Case gsCodProPF
'                            lsRTFRfm = lsRTFRfm & String(49, " ")
'                            lnCntRel = lnCntRel + 1
'                        Case gsCodProCTS
'                            lsRTFRfm = lsRTFRfm & String(49, " ")
'                            lnCntRel = lnCntRel + 1
'                    End Select
'                    lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & oImpresora.gPrnSaltoLinea
'                    lbCtaIde = True
'               End If
'            Else
'               lsCodCta = !cCodCta
'               lbCtaIde = False
'               lnCtaPag = lnCtaPag + 1: lnTotCta = lnTotCta + 1
'
'               Select Case Mid$(!cCodCta, 3, 3)
'                  Case gsCodProAC
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCta)), 6, False, 6, 0) & " "
'                       lsRTFRfm = lsRTFRfm & !cCodCta & " "
'                       If Len(Trim(PstaNombre(!cNomPers, False))) < 40 Then
'                          lsRTFRfm = lsRTFRfm & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40), 40, " ") & " "
'                       Else
'                          lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & " "
'                       End If
'                       lnIntCap = 0
'                       If lbCieMes Then
'                          lnIntCap = !nSdoDis - !nSdoAnt
'                       Else
'                          lnIntCap = !nIntMes
'                       End If
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCap)), 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoDis, 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoCnt, 12, True, 9, 2) & "   "
'                       lsRTFRfm = lsRTFRfm & Format(!dUltMovAC, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
'
'                       lnCapita = lnCapita + lnIntCap: lnCapiFi = lnCapiFi + lnIntCap
'                       lnSdoCnt = lnSdoCnt + !nSdoCnt
'                       lnSdoDis = lnSdoDis + !nSdoDis
'                       lnSdCnFi = lnSdCnFi + !nSdoCnt
'                       lnSdDiFi = lnSdDiFi + !nSdoDis
'                       lnCntRel = RelCta(!cCodCta, !cCodPers)
'                       lbCtaIde = True
'                  Case gsCodProPF
'                       lnCalInt = 0
'                       lnNumDia = 0
'                       lnNumDia = DateDiff("d", !dUltMovPF, !dFchVct) - 1
'                       lnCalInt = GetInteres(!nTasaIntPF, !nIntDev + (!nIntDev + !nDepoIni), lnNumDia, False)
'
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCta)), 6, False, 6, 0) & " "
'                       lsRTFRfm = lsRTFRfm & Format(!dAperPF, "dd/mm/yyyy") & " "
'                       lsRTFRfm = lsRTFRfm & Format(!dFchVct, "dd/mm/yyyy") & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nPlazo)), 6, False, 6, 0) & " "
'                       lsRTFRfm = lsRTFRfm & !cCodCta & " "
'                       If Len(Trim(PstaNombre(!cNomPers, False))) < 40 Then
'                          lsRTFRfm = lsRTFRfm & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40), 40, " ") & " "
'                       Else
'                          lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & " "
'                       End If
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nDepoIni)), 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nCapital)), 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nIntDev)), 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCalInt)), 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(!nCapital + lnCalInt)), 12, True, 9, 2) & "   "
'                       lsRTFRfm = lsRTFRfm & Format(!dUltMovPF, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
'
'                       lnApertu = lnApertu + !nDepoIni: lnTotApe = lnTotApe + !nDepoIni
'                       lnCapita = lnCapita + !nCapital: lnTotCap = lnTotCap + !nCapital
'                       lnIntDev = lnIntDev + !nIntDev: lnToInDv = lnToInDv + !nIntDev
'                       lnIntCal = lnIntCal + lnCalInt: lnToInCa = lnToInCa + lnCalInt
'                       lnTotalP = lnTotalP + (!nCapital + lnCalInt)
'                       lnTotalF = lnTotalF + (!nCapital + lnCalInt)
'
'                       lnCntRel = RelCta(!cCodCta, !cCodPers)
'                       lbCtaIde = True
'                  Case gsCodProCTS
'                       lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCta)), 6, False, 6, 0) & " "
'                       lsRTFRfm = lsRTFRfm & Format(!dAperCTS, "dd/mm/yyyy") & " "
'                       lsRTFRfm = lsRTFRfm & String(10, " ") & " "
'                       lsRTFRfm = lsRTFRfm & String(6, " ") & " "
'                       lsRTFRfm = lsRTFRfm & !cCodCta & " "
'                       If Len(Trim(PstaNombre(!cNomPers, False))) < 40 Then
'                          lsRTFRfm = lsRTFRfm & FillText(Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40), 40, " ") & " "
'                       Else
'                          lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & " "
'                       End If
'                       lsRTFRfm = lsRTFRfm & JDNum(!nUltDep, 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(!nCapital, 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(!nIntMes, 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoDis, 12, True, 9, 2) & " "
'                       lsRTFRfm = lsRTFRfm & JDNum(!nSdoCnt, 12, True, 9, 2) & "   "
'                       lsRTFRfm = lsRTFRfm & Format(!dUltMovCTS, "dd/mm/yyyy") & oImpresora.gPrnSaltoLinea
'
'                       lnCapita = lnCapita + !nCapital
'                       lnCapiFi = lnCapiFi + !nCapital
'                       lnSdoCnt = lnSdoCnt + !nSdoCnt
'                       lnSdoDis = lnSdoDis + !nSdoDis
'                       lnSdCnFi = lnSdCnFi + !nSdoCnt
'                       lnSdDiFi = lnSdDiFi + !nSdoDis
'
'                       lnApertu = lnApertu + !nUltDep: lnTotApe = lnTotApe + !nUltDep
'                       lnCapita = lnCapita + !nCapital: lnTotCap = lnTotCap + !nCapital
'                       lnIntDev = lnIntDev + !nIntMes: lnToInDv = lnToInDv + !nIntMes
'                       lnIntCal = lnIntCal + !nSdoDis: lnToInCa = lnToInCa + !nSdoDis
'                       lnTotalP = lnTotalP + !nSdoCnt: lnTotalF = lnTotalF + !nSdoCnt
'
'                       lnCntRel = RelCta(!cCodCta, !cCodPers)
'                       lbCtaIde = True
'               End Select
'            End If
'            lnCntCta = lnCntCta + 1
'            lnLinPag = lnLinPag + lnCntRel + 1
'            If lnCtaPag Mod 4 = 0 Then
'                lsReporte = lsReporte & lsRTFRfm
'                lsRTFRfm = ""
'            End If
'            .MoveNext
'         Loop
'    End With
'    If prQryDat.EOF Then
'       If Mid$(lsCodCta, 3, 3) = gsCodProAC Then
'          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'          lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(21, " ")
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoDis)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdoCnt)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea             ' hasta 160 caracteres x línea
'          lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(24, " ")
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapiFi)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdDiFi)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnSdCnFi)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'          lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
'          lnCapita = 0: lnSdoCnt = 0: lnSdoDis = 0: lnCtaPag = 0
'          lnCapiFi = 0: lnSdCnFi = 0: lnSdDiFi = 0: lnTotCta = 0
'       Else
'          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea              ' hasta 160 caracteres x línea
'          lsRTFRfm = lsRTFRfm & "Registros x Pag. " & FillNum(Trim(Str(lnCtaPag)), 5, " ") & "   Total Pag. " & lsNumPag & String(50, " ")
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnApertu)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnCapita)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntDev)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnIntCal)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalP)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'          lsRTFRfm = lsRTFRfm & String(lnCarLin, "-") & oImpresora.gPrnSaltoLinea             ' hasta 160 caracteres x línea
'          lsRTFRfm = lsRTFRfm & "Total Registros  " & FillNum(Trim(Str(lnTotCta)), 5, " ") & "   Acum. Total " & String(53, " ")
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotApe)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotCap)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInDv)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnToInCa)), 12, True, 9, 2) & " "
'          lsRTFRfm = lsRTFRfm & JDNum(Trim(Str(lnTotalF)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
'          lsRTFRfm = lsRTFRfm & String(lnCarLin, "=") & oImpresora.gPrnSaltoPagina             ' hasta 160 caracteres x línea
'          lnApertu = 0: lnCapita = 0: lnIntDev = 0: lnIntCal = 0: lnTotalP = 0
'          lnTotApe = 0: lnTotCap = 0: lnToInDv = 0: lnToInCa = 0: lnTotalF = 0
'       End If
'    End If
'    lsReporte = lsReporte & lsRTFRfm
'End Sub


'Private Function RelCta(psCodCta As String, psCodPer As String) As Integer
'Dim lsQrySd1 As String
'Dim rsQrySd1 As New ADODB.Recordset
'Dim lnCntRel As Integer
'
'lsQrySd1 = ""
'lsQrySd1 = "SELECT PC.cCodCta, PC.cRelaCta, PE.cNomPers " _
'    & "FROM PersCuenta PC INNER JOIN Persona PE ON PC.cCodPers = PE.cCodPers " _
'    & "WHERE PC.cCodCta = '" & psCodCta & "' AND PC.cCodPers <> '" & psCodPer & "' " _
'    & "AND PC.cRelaCta = 'TI'"
'
'    If lsQrySd1 = "" Then
'       Exit Function
'    End If
'    If rsQrySd1.State = adStateOpen Then rsQrySd1.Close
'    rsQrySd1.Open lsQrySd1, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RSVacio(rsQrySd1) Then
'       rsQrySd1.Close
'       Set rsQrySd1 = Nothing
'       RelCta = 0
'       Exit Function
'    Else
'       lnCntRel = 0
'       With rsQrySd1
'            Do While Not .EOF
'               Select Case Mid$(psCodCta, 3, 3)
'                    Case gsCodProAC
'                         lsRTFRfm = lsRTFRfm & String(20, " ")
'                         lnCntRel = lnCntRel + 1
'                    Case gsCodProPF
'                         lsRTFRfm = lsRTFRfm & String(49, " ")
'                         lnCntRel = lnCntRel + 1
'                    Case gsCodProCTS
'                         lsRTFRfm = lsRTFRfm & String(49, " ")
'                         lnCntRel = lnCntRel + 1
'               End Select
'               lsRTFRfm = lsRTFRfm & Left(PstaNombre(ImpreCarEsp(!cNomPers), False), 40) & oImpresora.gPrnSaltoLinea
'               .MoveNext
'            Loop
'       End With
'       rsQrySd1.Close
'       Set rsQrySd1 = Nothing
'       RelCta = lnCntRel
'    End If
'
'End Function

'   ------------------------------------------------------------
'   Función     :   JDNum
'   Propósito   :   Justifica a la derecha un campo numérico en
'                   un reporte
'   Uso         :   Reportes en general
'   Parámetro(s):   pnCampos -> Importe
'                   pnLongit -> Longitud del Campo
'                   pbComass -> Número con separador de miles
'                   pnDigEnt -> Cantidad de digitos enteros
'                   pnDigDec -> Cantidad de digitos decimales
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'
Private Function JDNum(pnCampos As String, pnLongit As Integer, _
                      pbComass As Boolean, pnDigEnt As Integer, _
                      pnDigDec As Integer)

    Dim Formato As String, I As Integer, lnPosDig As Integer

    If pnCampos = "0.00" Then
       JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))
       Exit Function
    End If

    If pbComass Then
       lnPosDig = 0
       For I = 1 To pnDigEnt
           lnPosDig = lnPosDig + 1
           Select Case lnPosDig
              Case 1
                   Formato = "0" & Formato
              Case 4, 7, 10
                   Formato = "#," & Formato
              Case Else
                   Formato = "#" & Formato
           End Select
       Next I
       If pnDigDec > 0 Then
          Formato = Formato & "." & String(pnDigDec, "0")
       End If
    Else
       For I = 1 To pnDigEnt
           Formato = IIf(I = 1, "0", "#") & Formato
       Next I
       If pnDigDec > 0 Then
          Formato = Formato & "." & String(pnDigDec, "0")
       End If
    End If
    pnCampos = Format(pnCampos, Formato)

    JDNum = Format(Trim(pnCampos), String(pnLongit, "@"))

End Function

'   ------------------------------------------------------------
'   Función     :   CabeRepo
'   Propósito   :   Define cabecera del reporte
'   Parámetro(s):   pnCarLin -> Caracteres x linea del reporte
'                   psSeccio -> Area responsable
'                   psTitRp1 -> Titulo 1 del reporte
'                   psTitRp2 -> Titulo 2 del reporte
'                   psMoneda -> Tipo de Moneda
'   Creado      :   02/07/1999  -   FAOS
'   Modificado  :   02/07/1999  -   FAOS
'   ------------------------------------------------------------
'   Formato     :   CabeRepo()
Private Function CabeRepo(psCabe01 As String, psCabe02 As String, _
                         pnCarLin As Integer, psSeccio As String, _
                         psTitRp1 As String, psTitRp2 As String, _
                         psMoneda As String, psNumPag As String) As String

    Dim lsTitRp1 As String, lsTitRp2 As String
    Dim lsMoneda As String

    lsTitRp1 = "": lsTitRp2 = ""
    CabeRepo = ""
    lsMoneda = ""
    lsMoneda = IIf(psMoneda = "", String(10, " "), " - " & psMoneda)
'    psCabe01 = ""
'    psCabe02 = ""

    '   Definición de Cabecera 1

    psCabe01 = FillText(UCase(Trim(gsNomCmac) & " - " & Trim(gsNomAge)) & lsMoneda, 36, " ")
    psCabe01 = psCabe01 & Space((pnCarLin - 36) - (Len(psCabe01) - 2))
    psCabe01 = psCabe01 & "PAGINA: " & psNumPag
    psCabe01 = psCabe01 & Space(5) & "FECHA: " & Format(gdFecSis, "dd/mm/yyyy")
    '   Definición de Cabecera 2
    psCabe02 = FillText(psSeccio, 19, " ")
    psCabe02 = psCabe02 & Space((pnCarLin - 19) - (Len(psCabe02) - 2))
    psCabe02 = psCabe02 & "HORA :   " & Format(Now(), "hh:mm:ss")
    '   Definición del Titulo del Reporte
    lsTitRp1 = String(Int((pnCarLin - Len(psTitRp1)) / 2), " ") & BON & psTitRp1 & BOFF
    lsTitRp2 = String(Int((pnCarLin - Len(psTitRp2)) / 2), " ") & BON & psTitRp2 & BOFF

    CabeRepo = CabeRepo & psCabe01 & oImpresora.gPrnSaltoLinea
    CabeRepo = CabeRepo & psCabe02 & oImpresora.gPrnSaltoLinea
    CabeRepo = CabeRepo & lsTitRp1 & oImpresora.gPrnSaltoLinea
    CabeRepo = CabeRepo & lsTitRp2

End Function



