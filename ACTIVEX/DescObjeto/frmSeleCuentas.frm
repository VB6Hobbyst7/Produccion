VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSeleCuentas 
   Caption         =   "Cuentas Contables : Selección"
   ClientHeight    =   5145
   ClientLeft      =   1995
   ClientTop       =   1815
   ClientWidth     =   6825
   Icon            =   "frmSeleCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6825
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDebe 
      Height          =   1995
      Left            =   90
      TabIndex        =   4
      Top             =   315
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3519
      _Version        =   393216
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483632
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Asigna Cuenta Contable a Asiento"
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Asigna Cuenta Contable a Asiento"
      Top             =   4710
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgHaber 
      Height          =   1995
      Left            =   90
      TabIndex        =   5
      Top             =   2610
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3519
      _Version        =   393216
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483632
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DEBE"
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
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "HABER"
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
      Left            =   90
      TabIndex        =   2
      Top             =   2400
      Width           =   645
   End
End
Attribute VB_Name = "frmSeleCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsD As New ADODB.Recordset
Dim rsH As New ADODB.Recordset
Public sCtaCod As String
Public sCtaDesc As String
Public sDH As String
Dim nUltFoco As Integer
Dim Ok As Boolean
Public Sub inicio(rsDebe As ADODB.Recordset, rsHaber As ADODB.Recordset)
Set rsD = rsDebe
Set rsH = rsHaber
Me.Show 1
End Sub
Private Sub DefineFormato(fg As MSHFlexGrid)
fg.ColWidth(0) = 300
fg.ColWidth(1) = 1700
fg.ColWidth(2) = 4300
fg.TextMatrix(0, 1) = "Cuenta"
fg.TextMatrix(0, 2) = "Descripción"
End Sub
Private Sub cmdAceptar_Click()
Ok = True
If nUltFoco = 1 Then
   sDH = "D"
   sCtaCod = fgDebe.TextMatrix(fgDebe.Row, 1)
   sCtaDesc = fgDebe.TextMatrix(fgDebe.Row, 2)
Else
   sDH = "H"
   sCtaCod = fgHaber.TextMatrix(fgHaber.Row, 1)
   sCtaDesc = fgHaber.TextMatrix(fgHaber.Row, 2)
End If
Unload Me
End Sub
Private Sub cmdCancelar_Click()
Ok = False
Unload Me
End Sub
Private Sub fgDebe_GotFocus()
nUltFoco = 1
End Sub
Private Sub fgHaber_GotFocus()
nUltFoco = 2
End Sub
Private Sub Form_Load()
CentraForm Me
Ok = False
Set fgDebe.DataSource = rsD
Set fgHaber.DataSource = rsH
DefineFormato fgDebe
DefineFormato fgHaber

End Sub
Public Property Get lOk() As Boolean
lOk = Ok
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
Ok = vNewValue
End Property
Public Property Get psCtaCod() As String
psCtaCod = sCtaCod
End Property

Public Property Let psCtaCod(ByVal vNewValue As String)
sCtaCod = vNewValue
End Property

Public Property Get psCtaDesc() As String
psCtaDesc = sCtaDesc
End Property

Public Property Let psCtaDesc(ByVal vNewValue As String)
sCtaDesc = vNewValue
End Property

Public Property Get psDH() As String
psDH = sDH
End Property

Public Property Let psDH(ByVal vNewValue As String)
sDH = vNewValue
End Property
