VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCapRepCargosAI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emision de Cargos para Cartas de Actividades Ilicitas"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmCapRepCargosAI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   2610
      Width           =   1110
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2610
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro de Impresión"
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
      Height          =   2325
      Left            =   195
      TabIndex        =   6
      Top             =   165
      Width           =   6300
      Begin VB.CheckBox chkPro3 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3825
         TabIndex        =   16
         Top             =   1725
         Width           =   870
      End
      Begin VB.CheckBox ChkPro2 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2775
         TabIndex        =   15
         Top             =   1725
         Width           =   870
      End
      Begin VB.CheckBox chkPro1 
         Caption         =   "Ica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1845
         TabIndex        =   14
         Top             =   1725
         Width           =   735
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Por Fecha"
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
         Left            =   150
         TabIndex        =   10
         Top             =   900
         Width           =   1245
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Por Nro Carta"
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
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1620
      End
      Begin MSMask.MaskEdBox txtDesde2 
         Height          =   300
         Left            =   2400
         TabIndex        =   2
         Top             =   1050
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtHasta2 
         Height          =   300
         Left            =   4770
         TabIndex        =   3
         Top             =   1095
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtDesde1 
         Height          =   300
         Left            =   2385
         TabIndex        =   0
         Top             =   420
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "##########"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHasta1 
         Height          =   300
         Left            =   4770
         TabIndex        =   1
         Top             =   420
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         ClipMode        =   1
         MousePointer    =   99
         Appearance      =   0
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "##########"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Procedencia"
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
         Height          =   225
         Left            =   435
         TabIndex        =   13
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
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
         Left            =   1800
         TabIndex        =   12
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Left            =   4125
         TabIndex        =   11
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   435
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
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
         Left            =   4125
         TabIndex        =   7
         Top             =   435
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCapRepCargosAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkPro1_Click()
  If chkPro1.value = vbChecked Then
        ChkPro2.value = vbUnchecked
        chkPro3.value = vbUnchecked
  End If
  
End Sub

Private Sub ChkPro2_Click()
 If ChkPro2.value = vbChecked Then
        chkPro1.value = vbUnchecked
        chkPro3.value = vbUnchecked
 End If
  
End Sub

Private Sub chkPro3_Click()
 If chkPro3.value = vbChecked Then
        chkPro1.value = vbUnchecked
        ChkPro2.value = vbUnchecked
 End If
End Sub

Private Sub cmdImprimir_Click()
Dim bResult As Boolean, sUbica As String
Dim clsServ As COMDCaptaServicios.DCOMCaptaServicios
Set clsServ = New COMDCaptaServicios.DCOMCaptaServicios

  sUbica = ""
  sUbica = IIf(chkPro1.value = vbChecked, "11", IIf(ChkPro2.value = vbChecked, "<>11", ""))

  If Opt1.value = True Then
         bResult = clsServ.GetCargosCartasAI(1, Trim(TxtDesde1.Text), Trim(txtHasta1.Text), sUbica, gsCodUser)
         
  Else
        Dim dfechaini As String, dfechafin As String
        dfechaini = Format(CDate(txtDesde2.Text), "yyyymmdd")
        dfechafin = Format(CDate(TxtHasta2.Text), "yyyymmdd")
        bResult = clsServ.GetCargosCartasAI(2, dfechaini, dfechafin, sUbica, gsCodUser)
         
  End If
  
  If bResult = False Then
     MsgBox "No hay información para los parámetros indicados.", vbOKOnly + vbInformation, "AVISO"
  End If
Set clsServ = Nothing

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{Tab}"
 
End Sub

Private Sub opt1_Click()
 TxtDesde1.Enabled = True
 txtHasta1.Enabled = True
 
 txtDesde2.Enabled = False
 TxtHasta2.Enabled = False
End Sub

Private Sub Opt2_Click()
TxtDesde1.Enabled = False
txtHasta1.Enabled = False
 
 txtDesde2.Enabled = True
 TxtHasta2.Enabled = True
End Sub
