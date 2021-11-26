VERSION 5.00
Begin VB.Form frmRepBaseFormulaOpc 
   Caption         =   "Reportes en Base a Fórmulas: Opciones de Impresión"
   ClientHeight    =   2640
   ClientLeft      =   2055
   ClientTop       =   1950
   ClientWidth     =   6870
   Icon            =   "frmRepBaseFormulaOpc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Columnas"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6675
      Begin VB.CheckBox chkDiferencia 
         Caption         =   "Diferencia"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkMEDAnterior 
         Caption         =   "Moneda Extranjera Exp. en ME del día anterior al corte"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox cMED 
         Caption         =   "Moneda Extranjera Exp. en ME"
         Height          =   345
         Left            =   3960
         TabIndex        =   8
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.CheckBox cMES 
         Caption         =   "Moneda Extranjera Exp. en MN"
         Height          =   345
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox cCOAj 
         Caption         =   "Consolidado Ajustado"
         Height          =   345
         Left            =   3960
         TabIndex        =   6
         Top             =   300
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CheckBox cMNHist 
         Caption         =   "Moneda Nacional Histórico"
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   750
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.CheckBox cMNAj 
         Caption         =   "Moneda Nacional Ajustado"
         Height          =   345
         Left            =   3960
         TabIndex        =   4
         Top             =   750
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox cCOHist 
         Caption         =   "Consolidado Histórico"
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Value           =   1  'Checked
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3300
      TabIndex        =   1
      Top             =   2250
      Width           =   1320
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   1920
      TabIndex        =   0
      Top             =   2250
      Width           =   1290
   End
End
Attribute VB_Name = "frmRepBaseFormulaOpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lCOAj As Integer
Dim lCOHi As Integer
Dim lMNAj As Integer
Dim lMNHi As Integer
Dim lMES  As Integer
Dim lMED  As Integer
Dim TlCOAj As Integer
Dim TlCOHi As Integer
Dim TlMNAj As Integer
Dim TlMNHi As Integer
Dim TlMES  As Integer
Dim TlMED  As Integer
'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
Dim fnMEDAnterior As Integer
Dim fnMEDAnterior2 As Integer
Dim fnDiferencia As Integer
Dim fnDiferencia2 As Integer
Dim fsCodigoReporte As String
'***Fin Agregado por ELRO*************************************

Public Property Get plCOAj() As Integer
plCOAj = lCOAj
End Property

Public Property Let plCOAj(ByVal vNewValue As Integer)
lCOAj = vNewValue
End Property


Public Property Get plCOHist() As Integer
plCOHist = lCOHi
End Property

Public Property Let plCOHist(ByVal vNewValue As Integer)
lCOHi = vNewValue
End Property

Public Property Get plMNAj() As Integer
plMNAj = lMNAj
End Property

Public Property Let plMNAj(ByVal vNewValue As Integer)
lMNAj = vNewValue
End Property

Public Property Get plMNHist() As Integer
plMNHist = lMNHi
End Property

Public Property Let plMNHist(ByVal vNewValue As Integer)
lMNHi = vNewValue
End Property

Public Property Get plMES() As Integer
plMES = lMES
End Property

Public Property Let plMES(ByVal vNewValue As Integer)
lMES = vNewValue
End Property

Public Property Get plMED() As Integer
plMED = lMED
End Property

Public Property Let plMED(ByVal vNewValue As Integer)
lMED = vNewValue
End Property

'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
Public Property Get pfnMEDAnterior() As Integer
pfnMEDAnterior = fnMEDAnterior
End Property
Public Property Let pfnMEDAnterior(ByVal vNewValue As Integer)
fnMEDAnterior = vNewValue
End Property
Public Property Get pfnDiferencia() As Integer
pfnDiferencia = fnDiferencia
End Property
Public Property Let pfnDiferencia(ByVal vNewValue As Integer)
fnDiferencia = vNewValue
End Property
'***Fin Agregado por ELRO*************************************

Private Sub cmdAceptar_Click()
lCOAj = cCOAj.value
lCOHi = cCOHist.value
lMNAj = cMNAj.value
lMNHi = cMNHist.value
lMES = cMES.value
lMED = cMED.value
'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
pfnMEDAnterior = chkMEDAnterior.value
pfnDiferencia = chkDiferencia.value
fsCodigoReporte = ""
'***Fin Agregado por ELRO*************************************
Unload Me
End Sub

Private Sub cmdCancelar_Click()
lCOAj = TlCOAj
lCOHi = TlCOHi
lMNAj = TlMNAj
lMNHi = TlMNHi
lMES = TlMES
lMED = TlMED
'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
fnMEDAnterior = fnMEDAnterior2
fnDiferencia = fnDiferencia2
fsCodigoReporte = ""
'***Fin Agregado por ELRO*************************************
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
TlCOAj = lCOAj
TlCOHi = lCOHi
TlMNAj = lMNAj
TlMNHi = lMNHi
TlMES = lMES
TlMED = lMED
cCOAj.value = lCOAj
cCOHist.value = lCOHi
cMNAj.value = lMNAj
cMNHist.value = lMNHi
cMES.value = lMES
cMED.value = lMED
'***Agregado por ELRO el 20111222, según Acta N° 352-2011/TI-D
fnMEDAnterior2 = pfnMEDAnterior
fnDiferencia2 = pfnDiferencia
chkMEDAnterior.value = pfnMEDAnterior
chkDiferencia.value = pfnDiferencia
If fsCodigoReporte = "770090" Then
    chkMEDAnterior.Visible = True
    chkDiferencia.Visible = True
End If
'***Fin Agregado por ELRO*************************************
CentraForm Me
End Sub

Public Sub inicio(ByVal psCodigoReporte As String)
    fsCodigoReporte = psCodigoReporte
    Show 1
End Sub
