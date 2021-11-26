VERSION 5.00
Begin VB.Form frmRepOpePorCajero 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Operaciones por Cajero"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   Icon            =   "frmRepOpePorCajero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   45
      TabIndex        =   4
      Top             =   900
      Width           =   6420
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   360
         Left            =   5040
         TabIndex        =   5
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   6420
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         Height          =   360
         Left            =   1980
         TabIndex        =   3
         Top             =   195
         Width           =   1230
      End
      Begin VB.TextBox TxtFecha 
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Text            =   "01/01/2008"
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   255
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmRepOpePorCajero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub GeneraReporte()
'     Dim D As New ADODB.Recordset
'       Dim lsArchivo As String
'        Dim xlAplicacion As Excel.Application
'        Dim xlLibro As Excel.Workbook
'        Dim xlLibros As Excel.Workbooks
'        Dim lsNomHoja As String
'        Dim xlHoja1 As Excel.Worksheet
'        Dim xlRango As Excel.Range
'        Dim i As Integer, j As Integer
'        Dim D As ComTiposAdmin.clsDataset
'
'        lsArchivo = "ReporteOperacionesCajero_" & Format(Now, "ddmmyyyyhhmmss") & ".XLS"
'
'        xlAplicacion = New Excel.Application
'
'        If File.Exists(Application.StartupPath & "\" & lsArchivo) Then
'            File.Delete (Application.StartupPath & "\" & lsArchivo)
'            xlLibros = xlAplicacion.Workbooks
'            xlLibro = xlLibros.Add
'        Else
'            xlLibros = xlAplicacion.Workbooks
'            xlLibro = xlLibros.Add
'        End If
'
'        xlHoja1 = xlLibro.Worksheets.Add
'
'        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 4)).Font.Bold = True
'
'        xlHoja1.Cells(1, 3) = " REPORTE DE OPERACIONES DE CAJEROS AUTOMATICOS AL : " & Me.TxtFecha.Text
'        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 4)).Merge (True)
'        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 4)).HorizontalAlignment = xlHAlignCenter
'
'        'Carga Data
'
'
'        xlHoja1.Cells(3, 1) = "CODIGO"
'        xlHoja1.Cells(3, 2) = "NOMBRE"
'        xlHoja1.Cells(3, 3) = "CARGO"
'        xlHoja1.Cells(3, 4) = "REMUNERACION BRUTA"
'        xlHoja1.Cells(3, 5) = "ASIGNACION FAM."
'        xlHoja1.Cells(3, 6) = "REMUN. POR VACAC."
'        xlHoja1.Cells(3, 7) = "OTRAS REMUN."
'        xlHoja1.Cells(3, 8) = "TOTAL REMUN."
'        xlHoja1.Cells(3, 9) = "APORTE OBL. AFP"
'        xlHoja1.Cells(3, 10) = "COMISION VAR. AFP"
'        xlHoja1.Cells(3, 11) = "PRIMA SEG. AFP"
'        xlHoja1.Cells(3, 12) = "IMPUESTO DE 5TA CAT."
'        xlHoja1.Cells(3, 13) = "ESSALUD VIDA"
'        xlHoja1.Cells(3, 14) = "OTROS DESC."
'        xlHoja1.Cells(3, 15) = "TORAL DESC."
'        xlHoja1.Cells(3, 16) = "EPS"
'        xlHoja1.Cells(3, 17) = "ESSALUD"
'        xlHoja1.Cells(3, 18) = "NETO A PAGAR"
'
'        xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, 18)).Font.Bold = True
'
'        i = 4
'        D.Inicio()
'        Do While Not D.EOF
'            xlHoja1.Cells(i, 1) = D.ObtenerDato("nCodEmp")
'            xlHoja1.Cells(i, 2) = D.ObtenerDato("cNombre")
'            xlHoja1.Cells(i, 3) = D.ObtenerDato("cCargo")
'            xlHoja1.Cells(i, 4) = D.ObtenerDato("nRemunBruta")
'            xlHoja1.Cells(i, 5) = D.ObtenerDato("nAsigFam")
'            xlHoja1.Cells(i, 6) = D.ObtenerDato("nRemVacac")
'            xlHoja1.Cells(i, 7) = D.ObtenerDato("nOtrRem")
'            xlHoja1.Cells(i, 8) = D.ObtenerDato("nSubTotRem")
'            xlHoja1.Cells(i, 9) = D.ObtenerDato("nAporObliAfp")
'            xlHoja1.Cells(i, 10) = D.ObtenerDato("nComVarAfp")
'            xlHoja1.Cells(i, 11) = D.ObtenerDato("nPrimaSegAfp")
'            xlHoja1.Cells(i, 12) = D.ObtenerDato("nImpQuinta")
'            xlHoja1.Cells(i, 13) = D.ObtenerDato("nESSaludVida")
'            xlHoja1.Cells(i, 14) = D.ObtenerDato("nOtrDesc")
'            xlHoja1.Cells(i, 15) = D.ObtenerDato("nSubTotDesc")
'            xlHoja1.Cells(i, 16) = D.ObtenerDato("nEPS")
'            xlHoja1.Cells(i, 17) = D.ObtenerDato("nESSalud")
'            xlHoja1.Cells(i, 18) = D.ObtenerDato("nNetoAPagar")
'
'
'            xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i, 18)).NumberFormat = "#,0.00"
'            xlHoja1.Range(xlHoja1.Cells(i, 1), xlHoja1.Cells(i, 1)).NumberFormat = "#,0"
'
'            i = i + 1
'            D.Siguiente()
'
'
'        Loop
'
'
'        xlHoja1.SaveAs (Application.StartupPath & "\" & lsArchivo)
'
'        'Elimina totalmente xlHoja1
'        Call EliminaObjeto(xlHoja1)
'        'Cierra el libro de trabajo
'        xlLibro.Close()
'        'xlAplicacion.Workbooks(1).Close(False)
'        Call EliminaObjeto(xlLibro)
'        xlLibros.Close()
'        Call EliminaObjeto(xlLibros)
'        ' Cierra Microsoft Excel con el método Quit.
'        xlAplicacion.Quit()
'
'        Call EliminaObjeto(xlAplicacion)

End Sub



Private Sub CmdSalir_Click()
    Unload Me
End Sub
