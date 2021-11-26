VERSION 5.00
Begin VB.Form frmProyeccionSeguimientoPorAgencia 
   Caption         =   "Seguimiento de Proyecciones Semanales"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   Icon            =   "frmProyeccionSeguimientoPorAgencia.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Mes - Año"
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
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   340
         Width           =   1095
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmProyeccionSeguimientoPorAgencia.frx":030A
         Left            =   120
         List            =   "frmProyeccionSeguimientoPorAgencia.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtAnio 
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   975
   End
   Begin SICMACT.FlexEdit feProyeccion 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8493
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Semana 1-Semana 2-Semana 3-Semana 4-Semana 5-Semana 6-cAgeCod"
      EncabezadosAnchos=   "300-3000-1200-1200-1200-1200-1200-1200-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmProyeccionSeguimientoPorAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmProyeccionSeguimientoPorAgencia
'***     Descripcion:       Permite ver el seguimiento de las proyecciones por agencia
'***     Creado por:        FRHU
'***     Fecha-Tiempo:         24/04/2014 01:00:00 PM
'*****************************************************************************************
Option Explicit
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdMostrar_Click()
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRs As ADODB.Recordset
    Dim Fila As Integer
    Dim Columna As Integer
    Dim pCodAge As String
    Dim pCodAgeAnte As String
    Dim TotalSemana As Integer
    'Validar Año
    If txtAnio.Text = "" Then
        MsgBox "Ingrese un Año", vbInformation, "Advertencia"
        Exit Sub
    End If
    If val(txtAnio.Text) < 1900 Or val(txtAnio.Text) > 9972 Then
        MsgBox "Año no Valido", vbInformation, "Advertencia"
        Exit Sub
    End If
    'Fin Validar Año
    Columna = 2
    Fila = 0
    pCodAgeAnte = "0"
    FormateaFlex feProyeccion
    Set oRs = oCred.ObtenerSeguimientoProyeccionSemanal(txtAnio.Text, cboMes.ListIndex + 1)
    If Not oRs.EOF And Not oRs.BOF Then
        Do While Not oRs.EOF
                pCodAge = oRs!cAgeCod
                If pCodAge <> pCodAgeAnte Then
'                    If fila <> 0 Then
'                        Do While Columna <= TotalSemana
'                            If feProyeccion.TextMatrix(fila, Columna) = "" Then
'                                feProyeccion.TextMatrix(fila, Columna) = "N/A"
'                            End If
'                            Columna = Columna + 1
'                        Loop
'                    End If
                    Fila = Fila + 1
                    feProyeccion.AdicionaFila
                End If
                    feProyeccion.TextMatrix(Fila, 1) = oRs!cAgeDescripcion
                    Select Case oRs!nSemana
                        Case 1: feProyeccion.TextMatrix(Fila, 2) = Space(10) & IIf(oRs!valor = 1, "Si", "No") & Space(50) & oRs!dfechafin & " " & oRs!dfechaini
                        Case 2: feProyeccion.TextMatrix(Fila, 3) = Space(10) & IIf(oRs!valor = 1, "Si", "No") & Space(50) & oRs!dfechafin & " " & oRs!dfechaini
                        Case 3: feProyeccion.TextMatrix(Fila, 4) = Space(10) & IIf(oRs!valor = 1, "Si", "No") & Space(50) & oRs!dfechafin & " " & oRs!dfechaini
                        Case 4: feProyeccion.TextMatrix(Fila, 5) = Space(10) & IIf(oRs!valor = 1, "Si", "No") & Space(50) & oRs!dfechafin & " " & oRs!dfechaini
                        Case 5: feProyeccion.TextMatrix(Fila, 6) = Space(10) & IIf(oRs!valor = 1, "Si", "No") & Space(50) & oRs!dfechafin & " " & oRs!dfechaini
                        Case 6: feProyeccion.TextMatrix(Fila, 7) = Space(10) & IIf(oRs!valor = 1, "Si", "No") & Space(50) & oRs!dfechafin & " " & oRs!dfechaini
                    End Select
                    feProyeccion.TextMatrix(Fila, 8) = oRs!cAgeCod
                pCodAgeAnte = oRs!cAgeCod
                TotalSemana = oRs!nSemana
            oRs.MoveNext
        Loop
    End If
End Sub
Private Sub cmdVer_Click()
    If feProyeccion.Rows > 2 Then
        Call MostrarVentanaModal
    Else
        MsgBox "Debe Ingresar los parametros y hacer click en Mostrar", vbInformation
    End If
End Sub
Private Sub feProyeccion_DblClick()
    If feProyeccion.Rows > 2 Then
        Call MostrarVentanaModal
    Else
        MsgBox "Debe Ingresar los parametros y hacer click en Mostrar", vbInformation
    End If
End Sub
Private Sub MostrarVentanaModal()
    Dim lsAgencia As String
    Dim lsCodAge As String
    Dim lf As Integer
    Dim lc As Integer
    Dim ldFechaIni As Date
    Dim ldFechaFin As Date
    
    lf = feProyeccion.row
    lc = feProyeccion.Col
    If lf <> 0 And lc <> 0 And lc <> 1 Then
        lsAgencia = feProyeccion.TextMatrix(lf, 1)
        ldFechaIni = Right(feProyeccion.TextMatrix(lf, lc), 10)
        ldFechaFin = Left(Right(feProyeccion.TextMatrix(lf, lc), 21), 10)
        lsCodAge = feProyeccion.TextMatrix(lf, 8)
        Call frmProyeccionPorAgencias.MostrarProyeccionSemanal(lsCodAge, lsAgencia, ldFechaIni, ldFechaFin, 2)
    End If
End Sub

Private Sub txtAnio_Change()
    If Len(Me.txtAnio.Text) = 4 Then
        Me.cmdMostrar.SetFocus
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
