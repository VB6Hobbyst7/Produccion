VERSION 5.00
Begin VB.Form frmCredSugExonera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exoneraciones"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   Icon            =   "frmCredSugExonera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feTiposExonera 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _extentx        =   7594
      _extenty        =   5530
      cols0           =   5
      highlight       =   1
      encabezadosnombres=   "-cExoneraCod-Exoneracion-Tipo-Solicitar"
      encabezadosanchos=   "0-0-2530-600-800"
      font            =   "frmCredSugExonera.frx":030A
      font            =   "frmCredSugExonera.frx":0336
      font            =   "frmCredSugExonera.frx":0362
      font            =   "frmCredSugExonera.frx":038E
      fontfixed       =   "frmCredSugExonera.frx":03BA
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1  'True
      tipobusqueda    =   3
      columnasaeditar =   "X-X-X-X-4"
      listacontroles  =   "0-0-0-0-4"
      encabezadosalineacion=   "L-L-L-C-C"
      formatosedit    =   "0-1-0-0-0"
      lbeditarflex    =   -1  'True
      rowheight0      =   300
   End
End
Attribute VB_Name = "frmCredSugExonera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim rs As ADODB.Recordset

Private Sub cmdCerrar_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oDNiv.RecuperaTiposExoneraciones()
    Set oDNiv = Nothing
    Call LimpiaFlex(feTiposExonera)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feTiposExonera.AdicionaFila
            lnFila = feTiposExonera.Row
            feTiposExonera.TextMatrix(lnFila, 1) = rs!cExoneraCod
            feTiposExonera.TextMatrix(lnFila, 2) = rs!cExoneraDesc
            feTiposExonera.TextMatrix(lnFila, 3) = Left(rs!cTipoExoneraDesc, 1)
            rs.MoveNext
        Loop
        feTiposExonera.TopRow = 1
        feTiposExonera.Row = 1
    End If
End Sub
