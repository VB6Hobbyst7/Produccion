VERSION 5.00
Begin VB.Form frmCapVerReglas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reglas"
   ClientHeight    =   1845
   ClientLeft      =   12150
   ClientTop       =   2535
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   1065
   End
   Begin SICMACT.FlexEdit grdReglas 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2520
      _extentx        =   4445
      _extenty        =   3254
      highlight       =   1
      allowuserresizing=   3
      encabezadosnombres=   "#-Regla"
      encabezadosanchos=   "500-1600"
      font            =   "frmCapVerReglas.frx":0000
      font            =   "frmCapVerReglas.frx":002C
      font            =   "frmCapVerReglas.frx":0058
      font            =   "frmCapVerReglas.frx":0084
      fontfixed       =   "frmCapVerReglas.frx":00B0
      columnasaeditar =   "X-X"
      textstylefixed  =   4
      listacontroles  =   "0-0"
      encabezadosalineacion=   "C-C"
      formatosedit    =   "0-0"
      textarray0      =   "#"
      colwidth0       =   495
      rowheight0      =   300
   End
End
Attribute VB_Name = "frmCapVerReglas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************************
'***Nombre      : frmCapVerReglas
'***Descripción : Formulario creado para visualizar las reglas establecidas para cada cuenta
'***Creación    : RIRO el 20130501, según Proyecto de Ahorros - Definición de Poderes en Cuentas de Productos Pasivos
'***************************************************************************************************

Public Sub inicia(ByVal strReglas As String)
    Dim arrReglas() As String
    Dim v As Variant
    arrReglas = Split(strReglas, "-")
    For Each v In arrReglas
        grdReglas.AdicionaFila
        grdReglas.TextMatrix(grdReglas.row, 1) = v
    Next
    grdReglas.lbEditarFlex = False
    cmdSalir.Visible = True
    Me.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
