VERSION 5.00
Begin VB.Form frmCtasAhoPersona 
   Caption         =   "Cuentas de la Persona"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmCtasAhoPersona.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2310
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2310
      Width           =   1095
   End
   Begin SICMACT.FlexEdit flxCuentas 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _extentx        =   13573
      _extenty        =   3625
      cols0           =   5
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Cuenta-Titular-Moneda-nTpoPrograma"
      encabezadosanchos=   "400-2000-3500-1500-0"
      font            =   "frmCtasAhoPersona.frx":030A
      font            =   "frmCtasAhoPersona.frx":0336
      font            =   "frmCtasAhoPersona.frx":0362
      font            =   "frmCtasAhoPersona.frx":038E
      font            =   "frmCtasAhoPersona.frx":03BA
      fontfixed       =   "frmCtasAhoPersona.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X"
      listacontroles  =   "----0"
      encabezadosalineacion=   "L-L-L-C-C"
      formatosedit    =   "----0"
      textarray0      =   "#"
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCtasAhoPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lsCodCta As String
Public lsTitular As String
Public lsMoneda As String
Public lnTpoPrograma As Integer 'RIRO20150512 ERS146-2014

Private Sub cmdAceptar_Click()
    If flxCuentas.Rows > 1 Then
        lsCodCta = flxCuentas.TextMatrix(flxCuentas.row, 1)
        lsTitular = flxCuentas.TextMatrix(flxCuentas.row, 2)
        lsMoneda = flxCuentas.TextMatrix(flxCuentas.row, 3)
        lnTpoPrograma = flxCuentas.TextMatrix(flxCuentas.row, 4) 'RIRO20150512 ERS146-2014
    Else
        lsCodCta = ""
        lsTitular = ""
        lsMoneda = ""
        lnTpoPrograma = -1 ' RIRO20150512 ERS146-2014
    End If
    Unload Me
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Public Sub Inicia(ByVal psCodPers As String)
Dim rs As ADODB.Recordset
Dim i As Integer
Dim lsCodPers As String
Dim oComMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Set oComMov = New COMNCaptaGenerales.NCOMCaptaMovimiento

    lsCodPers = psCodPers
    
    Set rs = oComMov.DevuelveCtasPorCodPers(psCodPers)
    
    Set flxCuentas.Recordset = rs
    
    Me.Show (1)
End Sub

Private Sub flxCuentas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

