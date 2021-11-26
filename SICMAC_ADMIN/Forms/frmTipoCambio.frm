VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTipoCambio 
   Caption         =   "Tipo de Cambio"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form2"
   ScaleHeight     =   3855
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdgrabar 
      Caption         =   "Grabar"
      Height          =   735
      Left            =   480
      TabIndex        =   10
      Top             =   2760
      Width           =   4095
   End
   Begin Sicmact.EditMoney TCFIJODIA 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
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
   Begin Sicmact.EditMoney TCFIJO 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
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
   Begin Sicmact.EditMoney TCCOMPRA 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
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
   Begin Sicmact.EditMoney TCVENTA 
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57278465
      CurrentDate     =   38324
   End
   Begin VB.Label Label5 
      Caption         =   "T.C FIJO DIA"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "T.C FIJO"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label3 
      Caption         =   "T.C. COMPRA"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "T.C VENTA"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FECHA"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   525
   End
End
Attribute VB_Name = "frmTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim clsTC As dTipoCambio

Private Sub Cmdgrabar_Click()
Dim dfecha As Date
dfecha = DTPicker1.value
Dim RESULTADO As Integer

RESULTADO = clsTC.InsertaTipoCambio(dfecha, TCVENTA.value, TCCOMPRA.value, 0, 0, TCFIJO, TCFIJODIA.value, 0, "200405011212481080100ENMC", False)



End Sub



Private Sub Form_Load()

Set clsTC = New dTipoCambio


End Sub

