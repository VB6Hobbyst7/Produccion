VERSION 5.00
Begin VB.Form frmArqueoTarjDebBovedaDetFal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "frmArqueoTarjDebBovedaDetFal.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1400
      TabIndex        =   2
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feDetalleFaltante 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _extentx        =   6165
      _extenty        =   8916
      cols0           =   3
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-# Tarjeta-Faltante"
      encabezadosanchos=   "300-1800-1000"
      font            =   "frmArqueoTarjDebBovedaDetFal.frx":030A
      font            =   "frmArqueoTarjDebBovedaDetFal.frx":0336
      font            =   "frmArqueoTarjDebBovedaDetFal.frx":0362
      font            =   "frmArqueoTarjDebBovedaDetFal.frx":038E
      font            =   "frmArqueoTarjDebBovedaDetFal.frx":03BA
      fontfixed       =   "frmArqueoTarjDebBovedaDetFal.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2"
      listacontroles  =   "0-0-4"
      encabezadosalineacion=   "C-C-C"
      formatosedit    =   "0-0-0"
      textarray0      =   "#"
      lbeditarflex    =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   300
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmArqueoTarjDebBovedaDetFal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMatDetFaltante() As Variant
Dim gnCantHabilita As Long
Dim bConforme As Boolean 'GIPO ERS051-2016
Public Function Inicio(ByVal pnMatDetFaltante As Variant, ByVal nCantHabilita As Long, ByVal bConform) As Variant  'GIPO
    nMatDetFaltante = pnMatDetFaltante
    gnCantHabilita = nCantHabilita
    bConforme = bConform 'GIPO
    Me.Show 1
    Inicio = nMatDetFaltante
End Function
Private Sub CmdAceptar_Click()
    Dim b As Boolean
    Dim i As Integer
    Dim X As Integer
    Dim nCantHab As Long
    b = False
    nCantHab = 0
    For i = 1 To feDetalleFaltante.Rows - 1
        If feDetalleFaltante.TextMatrix(i, 2) = "." Then
            nCantHab = nCantHab + 1
            b = True
        End If
    Next
    If Not b Then
        MsgBox "No se ha seleccionado ninguna tarjeta faltante. Verifique", vbInformation
        Exit Sub
    End If
    If nCantHab <> gnCantHabilita Then
        MsgBox "La Cantidad de Tarjetas Seleccionadas no es igual a la Cantidad de Tarjetas Faltantes. Verifique.", vbInformation, "Aviso"
        Exit Sub
    End If
    For i = 1 To feDetalleFaltante.Rows - 1
        For X = 1 To UBound(nMatDetFaltante, 2)
            If nMatDetFaltante(4, X) = feDetalleFaltante.TextMatrix(i, 1) Then
                nMatDetFaltante(5, X) = IIf(feDetalleFaltante.TextMatrix(i, 2) = ".", 1, 0)
            End If
        Next
    Next
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
'GIPO ERS051-2016
    If bConforme Then
        Me.cmdAceptar.Visible = False
        Me.cmdCancelar.Visible = False
        Me.feDetalleFaltante.ColumnasAEditar = "X-X-X"
    End If
'END GIPO
    CargaDatos
End Sub
Private Sub CargaDatos()
Dim i As Integer
Dim row As Integer
    For i = 1 To UBound(nMatDetFaltante, 2)
        feDetalleFaltante.AdicionaFila
        row = feDetalleFaltante.row
        feDetalleFaltante.TextMatrix(row, 1) = nMatDetFaltante(4, i)
        feDetalleFaltante.TextMatrix(row, 2) = IIf(bConforme, 0, nMatDetFaltante(5, i)) 'Modified by GIPO
    Next
End Sub

