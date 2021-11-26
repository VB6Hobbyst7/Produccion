VERSION 5.00
Begin VB.Form frmCredListaVigAtraso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Créditos Vencidos"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   Icon            =   "frmCredListaVigAtraso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin SICMACT.FlexEdit FePolizas 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      _extentx        =   16113
      _extenty        =   4921
      cols0           =   6
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Nº de Crédito-Fecha Venc.-Días de atraso-Monto p/Canc.-Monto p/Renov."
      encabezadosanchos=   "400-2000-1400-1400-0-0"
      font            =   "frmCredListaVigAtraso.frx":030A
      font            =   "frmCredListaVigAtraso.frx":0336
      font            =   "frmCredListaVigAtraso.frx":0362
      font            =   "frmCredListaVigAtraso.frx":038E
      font            =   "frmCredListaVigAtraso.frx":03BA
      fontfixed       =   "frmCredListaVigAtraso.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1  'True
      columnasaeditar =   "X-X-X-X-X-X"
      listacontroles  =   "0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-L-C-R-R"
      formatosedit    =   "0-0-0-0-0-5"
      textarray0      =   "#"
      lbpuntero       =   -1  'True
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "CLIENTE TIENE CREDITOS VENCIDOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   4875
   End
End
Attribute VB_Name = "frmCredListaVigAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Dim nTipoOperacion As TipoOperacion

Public sAgencia As String
Public nInmueble As Integer
Public ntna  As Double
Public nPriMinima As Double
Public nDereEmi As Double
Public nalternativa As Integer
Public dfecha  As Date


Public cCodCli As String
Public sPersCodContr As String
Dim nEstadoPoliza As Integer

Public Sub Inicio(ByVal pRs As ADODB.Recordset)
    
FePolizas.Clear
FePolizas.FormaCabecera
FePolizas.rows = 2
FePolizas.rsFlex = pRs
    
Me.Show 1

End Sub


Private Sub cmdBuscar_Click()
Dim oPol As COMDCredito.DCOMCredDoc
Dim rs As ADODB.Recordset
Set oPol = New COMDCredito.DCOMCredDoc

Set rs = oPol.RecuperaCredPigVigVen(cCodCli, gdFecSis)

If rs.EOF Then MsgBox "No se encontraron datos.", vbInformation, "Mensaje"

FePolizas.Clear
FePolizas.FormaCabecera
FePolizas.rows = 2
FePolizas.rsFlex = rs
FePolizas.SetFocus
Set oPol = Nothing
End Sub
'RECO20150421*****************
Private Sub CmdAceptar_Click()
    Unload Me
End Sub

'RECO FIN ********************
Private Sub Form_Load()
    Call CentraForm(Me)
End Sub


