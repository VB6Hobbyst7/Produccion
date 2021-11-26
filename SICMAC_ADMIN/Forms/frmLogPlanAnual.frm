VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmLogPlanAnual 
   ClientHeight    =   5835
   ClientLeft      =   330
   ClientTop       =   2265
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   11325
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   5160
      Visible         =   0   'False
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame fraBoton 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   375
      Left            =   60
      TabIndex        =   20
      Top             =   5400
      Width           =   9915
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   2940
         TabIndex        =   27
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   1740
         TabIndex        =   26
         Top             =   0
         Width           =   1155
      End
      Begin VB.CommandButton cmdGenPlan 
         Caption         =   "Generar Plan Anual"
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdAprobacion 
         Caption         =   "Trámite de Aprobación"
         Height          =   375
         Left            =   7980
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4200
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdAge 
         Caption         =   "Agencias"
         Height          =   375
         Left            =   5460
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdAreas 
         Caption         =   "Areas"
         Height          =   375
         Left            =   6720
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSTmp 
      Height          =   615
      Left            =   2640
      TabIndex        =   19
      Top             =   6120
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   1085
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   14942183
      ForeColorSel    =   8388608
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.Frame fraTitulo 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      TabIndex        =   15
      Top             =   20
      Width           =   11175
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   16
         Text            =   "2005"
         Top             =   105
         Width           =   615
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7530
         TabIndex        =   18
         Top             =   120
         Width           =   3570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   525
         Left            =   7440
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label lblPlan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones del Año  "
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   120
         Width           =   6375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   525
         Left            =   0
         Top             =   0
         Width           =   7395
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10020
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlex 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   1800
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColorSel    =   14942183
      ForeColorSel    =   8388608
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   11175
      Begin VB.TextBox txtAprueba 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   5280
         TabIndex        =   14
         Top             =   900
         Width           =   5655
      End
      Begin VB.TextBox txtEjecutora 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   5280
         TabIndex        =   12
         Top             =   600
         Width           =   5655
      End
      Begin VB.TextBox txtPliego 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   960
         TabIndex        =   10
         Top             =   900
         Width           =   1635
      End
      Begin VB.TextBox txtSiglas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtRUC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   300
         Width           =   1755
      End
      Begin VB.TextBox txtEntidad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A86602&
         Height          =   280
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   300
         Width           =   7395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Instrumento - Aprueba o Modifica"
         Height          =   195
         Left            =   2820
         TabIndex        =   13
         Top             =   960
         Width           =   2340
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Ejecutora"
         Height          =   195
         Left            =   3900
         TabIndex        =   11
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pliego"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Siglas"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   660
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8520
         TabIndex        =   5
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmLogPlanAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String, nPlanAnualNro As Integer, bEditable As Boolean
Dim mMatMes(1 To 12) As String

Private Sub cmdAge_Click()
frmLogPlanAgeArea.Inicio 2, nPlanAnualNro, 2006
End Sub

Private Sub cmdAgregar_Click()
frmLogPlanAnualDetalle.Inicio nPlanAnualNro, 0, True
If frmLogPlanAnualDetalle.vpHaGrabado Then
    ListaPlanAnual
End If
End Sub

Private Sub cmdAreas_Click()
frmLogPlanAgeArea.Inicio 1, nPlanAnualNro, 2006
End Sub

Private Sub cmdAprobacion_Click()
Dim oConn As New DConecta
sSQL = ""
If MsgBox("¿ Está seguro de trámite de aprobación del Plan Anual " & txtAnio.Text & " ?" + Space(10), vbQuestion + vbYesNo, "Confirme operación") = vbYes Then

   sSQL = "UPDATE LogPlanAnual SET nPlanAnualEstado = 2 WHERE nPlanAnualNro = " & nPlanAnualNro & " "
   If oConn.AbreConexion Then
      oConn.Ejecutar sSQL
   End If
   
   oConn.CierraConexion
   DatosPlanAnual
   MsgBox "El trámite de aprobación del Plan Anual " & txtAnio.Text & " se ha iniciado con éxito!" + Space(10), vbInformation
End If
End Sub


Private Sub cmdImprimir_Click()
Dim appExcel As New Excel.Application
Dim wbExcel As Excel.Workbook
Dim nFil As Integer, nCol As Integer, cCelda As String
Dim i As Integer, j As Integer, k As Integer, n As Integer
Dim cLetra As String, nSuma As Currency
Dim cUltFil As String, cUltCol As String
Dim nNroFilas As Integer
Dim Kini As Integer

Set wbExcel = appExcel.Workbooks.Add
wbExcel.Worksheets(1).Range("A1").ColumnWidth = 12
wbExcel.Worksheets(1).Range("B1").ColumnWidth = 12
wbExcel.Worksheets(1).Range("C1").ColumnWidth = 8
wbExcel.Worksheets(1).Range("D1").ColumnWidth = 40
wbExcel.Worksheets(1).Range("E1").ColumnWidth = 15
wbExcel.Worksheets(1).Range("F1").ColumnWidth = 10
wbExcel.Worksheets(1).Range("G1").ColumnWidth = 12
wbExcel.Worksheets(1).Range("H1").ColumnWidth = 30
wbExcel.Worksheets(1).Range("I1").ColumnWidth = 40

wbExcel.Worksheets(1).Range("D1").value = "PLAN ANUAL DE ADQUISICIONES Y CONTRATACIONES DEL " + txtAnio.Text
wbExcel.Worksheets(1).Range("D1").Font.Bold = True

nFil = MSFlex.Rows - 1
nCol = MSFlex.Cols - 1
Barra.max = nFil
Barra.Visible = True
cLetra = ""
For i = 0 To nFil
    Barra.value = i
    k = i + 3
    If i = 0 Then
       Kini = k
       cLetra = ExcelColumnaString(nCol - 4)
       wbExcel.Worksheets(1).Range("A" + CStr(k) + ":" + cLetra + CStr(k)).value = MSFlex.TextMatrix(i, j)
       wbExcel.Worksheets(1).Range("A" + CStr(k) + ":" + cLetra + CStr(k)).Font.Bold = True
       wbExcel.Worksheets(1).Range("A" + CStr(k) + ":" + cLetra + CStr(k)).WrapText = True
       wbExcel.Worksheets(1).Range("A" + CStr(k) + ":" + cLetra + CStr(k)).RowHeight = 38
    End If
    For j = 5 To nCol
        cLetra = ExcelColumnaString(j - 4)
        cCelda = cLetra + CStr(k)
        wbExcel.Worksheets(1).Range(cCelda).value = MSFlex.TextMatrix(i, j)
        If i > 0 Then
           wbExcel.Worksheets(1).Range("A" + CStr(k) + ":" + cLetra + CStr(k)).RowHeight = 40
           
        End If
    Next j
    
Next i

cUltFil = cLetra + CStr(nFil + 3)
cUltCol = ExcelColumnaString(nCol - 4)
nNroFilas = i - 1

wbExcel.Worksheets(1).Range("A" + CStr(Kini) + ":" + cLetra + CStr(nFil + 3)).VerticalAlignment = 2
wbExcel.Worksheets(1).Range("A" + CStr(Kini) + ":C" + CStr(nFil + 3)).HorizontalAlignment = 3
wbExcel.Worksheets(1).Range("E" + CStr(Kini) + ":F" + CStr(nFil + 3)).HorizontalAlignment = 3
wbExcel.Worksheets(1).Range("A" + CStr(Kini) + ":" + cLetra + CStr(nFil + 3)).Font.Name = "Tahoma"
wbExcel.Worksheets(1).Range("A" + CStr(Kini) + ":" + cLetra + CStr(nFil + 3)).Font.Size = 8
wbExcel.Worksheets(1).Range("A" + CStr(Kini) + ":" + cLetra + CStr(nFil + 3)).Borders.LineStyle = 1
Barra.Visible = False
appExcel.Application.Visible = True
appExcel.Windows(1).Visible = True
End Sub

Private Sub cmdQuitar_Click()
Dim i As Integer, k As Integer

k = MSFlex.row



End Sub

Private Sub Form_Load()
CentraForm Me
Me.Caption = "Generación del Plan Anual de Adquisiciones y Contrataciones"

mMatMes(1) = "ENERO":     mMatMes(7) = "JULIO"
mMatMes(2) = "FEBRERO":   mMatMes(8) = "AGOSTO"
mMatMes(3) = "MARZO":     mMatMes(9) = "SEPTIEMBRE"
mMatMes(4) = "ABRIL":     mMatMes(10) = "OCTUBRE"
mMatMes(5) = "MAYO":      mMatMes(11) = "NOVIEMBRE"
mMatMes(6) = "JUNIO":     mMatMes(12) = "DICIEMBRE"

bEditable = False
nPlanAnualNro = 0
txtAnio.Text = Year(gdFecSis) + 1
FormaFlex
DatosPlanAnual
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Sub DatosPlanAnual()
Dim rs As New ADODB.Recordset, oConn As New DConecta

sSQL = "SELECT p.nPlanAnualNro,p.nPlanAnualAnio,p.cPlanAnualEntidad,p.cPlanAnualRUC,p.cPlanAnualSiglas, " & _
       "       P.cPlanAnualPliego , P.cPlanAnualEjecutor, P.cPlanAnualAprueba, P.nPlanAnualEstado, e.cEstado " & _
       "  from LogPlanAnual p inner join (select nConsValor as nEstado, cEstado=cConsDescripcion from Constante where nConsCod = 9049 and nConsCod<>nConsValor) e on p.nPlanAnualEstado = e.nEstado " & _
       " where nPlanAnualAnio = 2006 and nPlanAnualEstado >= 1"
       
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      nPlanAnualNro = rs!nPlanAnualNro
      txtEntidad.Text = rs!cPlanAnualEntidad
      txtSiglas.Text = rs!cPlanAnualSiglas
      txtRUC.Text = rs!cPlanAnualRUC
      txtEjecutora.Text = rs!cPlanAnualEjecutor
      txtAprueba.Text = rs!cPlanAnualAprueba
      txtPliego.Text = rs!cPlanAnualPliego
      lblEstado.Caption = rs!cEstado
      Select Case rs!nPlanAnualEstado
          Case 0
               cmdGenPlan.Visible = True
          Case 1
               cmdGenPlan.Visible = True
               cmdAprobacion.Visible = True
               lblEstado.ForeColor = "&H00C00000"
               bEditable = True
          Case 2
               cmdGenPlan.Visible = False
               'cmdConsucode.Visible = False
               cmdAprobacion.Visible = False
               lblEstado.ForeColor = "&H00000080"
               bEditable = False
          Case 3
               cmdGenPlan.Visible = True
               'cmdConsucode.Visible = True
               cmdAprobacion.Visible = False
               lblEstado.ForeColor = "&H00C00000"
               bEditable = False
      End Select
      DoEvents
   Else
      sSQL = "select p.cPersNombre, j.cPersJurSigla " & _
           "  from Persona p inner join PersonaJur j on p.cPersCod = j.cPersCod " & _
           " where p.cPersCod ='1120800013498'"

      If oConn.AbreConexion Then
         Set rs = oConn.CargaRecordSet(sSQL)
         oConn.CierraConexion
         If Not rs.EOF Then
            txtEntidad.Text = rs!cPersNombre
            txtSiglas.Text = rs!cPersJurSigla
         End If
      End If
      txtRUC.Text = "20132243230"
      txtEjecutora.Text = "LOGISTICA"
      txtAprueba.Text = "LOGISTICA"
      lblEstado.Caption = "SIN GENERAR"
      cmdAprobacion.Visible = False
      DoEvents
   End If
End If
ListaPlanAnual
End Sub

Function VerificaGrupos() As Boolean
Dim rs As New ADODB.Recordset
Dim oPlan As New DLogPlanAnual
Dim nAnio As Integer
Dim cArchivo As String, f As Integer, v As Variant

VerificaGrupos = True
nAnio = CInt(VNumero(txtAnio.Text))
Set rs = oPlan.TodosBienesRequeridos(nAnio, True)
If Not rs.EOF Then
   VerificaGrupos = False
   If MsgBox("Existen requerimientos de Bienes/Servicios" + Space(10) + vbCrLf + _
             " que no tienen código de Grupo asignado..." + vbCrLf + vbCrLf + _
             "  ¿ Desea actualizar códigos de grupo ?" + Space(10), vbQuestion + vbYesNo + vbDefaultButton2, "Confirme") = vbYes Then
      oPlan.ActualizaBSGrupos nAnio
   Else
      cArchivo = App.path + "\Bienes.txt"
      f = FreeFile
      Open cArchivo For Output As #f
      Print #f, ""
      Print #f, "RELACION DE BIENES / SERVICIOS QUE NO TIENEN CODIGO DE GRUPO"
      Print #f, "------------------------------------------------------------"
      Do While Not rs.EOF
         Print #f, rs!cProSelBSCod + "  " + rs!cBSDescripcion
         rs.MoveNext
      Loop
      Print #f, "------------------------------------------------------------"
      Print #f, "NOTA:"
      Print #f, "TODOS los Bienes / Servicios deben tener un codigo de grupo"
      Print #f, "antes de generar el Plan Anual"
      Close #f
      v = Shell("notepad.exe " + cArchivo, vbNormalFocus)
      Exit Function
   End If
End If
Set rs = Nothing
      
Set rs = oPlan.TodosBienesRequeridos(nAnio, True)
If Not rs.EOF Then
   MsgBox "No se pudo actualizar el grupo en algunos Bienes/Servicios" + Space(10) + vbCrLf + _
          Space(18) + "Revise la asignación de grupos", vbInformation
          
      cArchivo = App.path + "\Bienes.txt"
      f = FreeFile

      Open cArchivo For Output As #f
      Print #f, ""
      Print #f, "RELACION DE BIENES / SERVICIOS QUE NO TIENEN CODIGO DE GRUPO"
      Print #f, "------------------------------------------------------------"
      Do While Not rs.EOF
         Print #f, rs!cProSelBSCod + "  " + rs!cBSDescripcion
         rs.MoveNext
      Loop
      Print #f, "------------------------------------------------------------"
      Print #f, "NOTA:"
      Print #f, "TODOS los Bienes / Servicios deben tener un codigo de grupo"
      Print #f, "antes de generar el Plan Anual"
      Close #f
      v = Shell("notepad.exe " + cArchivo, vbNormalFocus)
   VerificaGrupos = False
Else
   VerificaGrupos = True
End If
End Function

Private Sub cmdGenPlan_Click()
Dim oConn As New DConecta, rs As New ADODB.Recordset
Dim nPlanAnualNro As Integer, nItem As Integer
Dim cProceso As String, cBSGrupoCod As String
Dim nTpoCod As Integer, nSubTpo As Integer
Dim nAnio As Integer, i As Integer
Dim oPlan As New DLogPlanAnual, sSQL1 As String, sSQL As String
Dim cCIIUCod As String, nObjetoCod As Integer
Dim cBSCod As String, nMoneda As Integer
Dim nPrecioUnitario As Currency, nCantidad As Integer

nAnio = CInt(VNumero(txtAnio.Text))

If VerificaGrupos Then
   fraBoton.Visible = True
Else
   fraBoton.Visible = False
   Exit Sub
End If

'-------------------------------------------------------------------------------

If MsgBox("El Plan Anual se genera sólo con requerimientos aprobados" + Space(10) + vbCrLf & _
          "¿ Generar Plan Anual de Adquisiciones y Contrataciones ?" + Space(10), vbQuestion + vbYesNo + vbDefaultButton2, "Confirme operación") = vbYes Then
   nPlanAnualNro = 0
   nPlanAnualNro = oPlan.GrabaCabeceraPlanAnual(nAnio, txtEntidad.Text, txtRUC.Text, txtSiglas.Text, txtEjecutora.Text, txtAprueba.Text, txtPliego.Text)

   If nPlanAnualNro > 0 Then
      Set rs = New ADODB.Recordset
      If oConn.AbreConexion Then
               
               
         sSQL1 = "delete from LogPlanAnualDetalleBS"
         oConn.Ejecutar sSQL1
         
         sSQL = "select t.nPlanNro,t.cCIIUCod,t.nObjetoCod,t.cBSGrupoCod,g.cBSGrupoDescripcion,t.nMoneda,sum(t.nSubTotal) as nMonto " & _
         " from (select nPlanNro=" & nPlanAnualNro & ", b.cCIIUCod, nObjetoCod=convert(tinyint,substring(b.cProSelBSCod,2,1)), d.cBSGrupoCod,nMoneda, " & _
         "       nSubTotal = nPrecioUnitario * Sum(nMes01 + nMes02 + nMes03 + nMes04 + nMes05 + nMes06 + nMes07 + nMes08 + nMes09 + nMes10 + nMes11 + nMes12) " & _
         "  From LogPlanAnualReqDetalle d inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
         " where nEstado=1 and nPlanReqNro in (select x.nPlanReqNro from " & _
         "      (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
         "       Where r.nAnio = " & nAnio & " and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
         "      (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro  Where x.nNro = Y.nApro) " & _
         " group by b.cCIIUCod, convert(tinyint,substring(b.cProSelBSCod,2,1)), d.cBSGrupoCod, nMoneda, nPrecioUnitario) t " & _
         " inner join BSGrupos g on t.cBSGrupoCod = g.cBSGrupoCod " & _
         " " & _
         " group by t.nPlanNro,t.cCIIUCod,t.nObjetoCod,t.cBSGrupoCod,g.cBSGrupoDescripcion,t.nMoneda  "
 
         'Set Rs = oPlan.GruposPlanAnualDetalleBS
         Set rs = oConn.CargaRecordSet(sSQL)
         If Not rs.EOF Then
            nItem = 0
            Do While Not rs.EOF
               nItem = nItem + 1
               cBSGrupoCod = rs!cBSGrupoCod
               
               cProceso = DeterminaProcesoSeleccion(rs!nObjetoCod, rs!nMonto, nTpoCod, nSubTpo)
               
               sSQL = "insert into LogPlanAnualDetalle (nPlanAnualNro,nPlanAnualItem,nPlanAnualAnio,nProSelTpoCod,nProSelSubTpo,cBSGrupoCod,nObjetoCod,cCIIU,cSintesis,nMoneda,nValorEstimado) " & _
                      " VALUES (" & nPlanAnualNro & "," & nItem & "," & nAnio & "," & nTpoCod & "," & nSubTpo & ",'" & cBSGrupoCod & "','" & rs!nObjetoCod & "','" & rs!cCIIUCod & "','" & rs!cBSGrupoDescripcion & "'," & rs!nMoneda & "," & rs!nMonto & ") "
               oConn.Ejecutar sSQL
            
               sSQL = "insert into LogPlanAnualDetalleBS (nPlanAnualNro, nPlanAnualItem, cCIIUCod, nObjetoCod, cBSCod,  cBSGrupoCod, nMoneda, nPrecioUnitario, nCantidad ) " & _
                      " select nPlanNro=" & nPlanAnualNro & "," & nItem & ", b.cCIIUCod, nObj=convert(tinyint,substring(b.cProSelBSCod,2,1)),b.cProSelBSCod, d.cBSGrupoCod, nMoneda,  " & _
                      "       nPrecioUnitario , nCantidad = Sum(nMes01 + nMes02 + nMes03 + nMes04 + nMes05 + nMes06 + nMes07 + nMes08 + nMes09 + nMes10 + nMes11 + nMes12) " & _
                      "  From LogPlanAnualReqDetalle d inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
                      " where nEstado=1 and nPlanReqNro in (select x.nPlanReqNro from " & _
                      "      (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
                      "       Where r.nAnio = " & nAnio & " and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
                      "      (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro  Where x.nNro = Y.nApro) " & _
                      "       and d.cBSGrupoCod = '" & cBSGrupoCod & "' " & _
                      " group by b.cProSelBSCod, b.cCIIUCod, convert(tinyint,substring(b.cProSelBSCod,2,1)), d.cBSGrupoCod, nMoneda, nPrecioUnitario"
               oConn.Ejecutar sSQL
 
               'sSQL = "UPDATE LogPlanAnualReqDetalle SET nPlanAnualNro = " & nPlanAnualNro & ", nPlanAnualItem = " & nItem & " WHERE cBSGrupoCod = '" & cBSGrupoCod & "' "
               'oConn.Ejecutar sSQL
                  
               'sSQL = "UPDATE LogPlanAnualDetalleBS SET nPlanAnualItem = " & nItem & " " & _
               '       " WHERE cBSGrupoCod = '" & cBSGrupoCod & "'   and " & _
               '       "          cCIIUCod = '" & Rs!cCIIUCod & "'   and " & _
               '       "        nObjetoCod =  " & Rs!nObjetoCod & "  and " & _
               '       "           nMoneda = '" & Rs!nMoneda & "' "
               'oConn.Ejecutar sSQL
               
               rs.MoveNext
            Loop
         End If
         ListaPlanAnual
         MsgBox "Se ha generado el PLAN ANUAL !" + Space(10), vbInformation
      End If
   End If
End If
End Sub



         
      
'         sSQL = " select nPlan=" & nPlanAnualNro & ", b.cCIIUCod,nObj=convert(tinyint,substring(b.cProSelBSCod,2,1)), b.cProSelBSCod, d.cBSGrupoCod,nMoneda,nPrecioUnitario,nCantidad=sum(nMes01+nMes02+nMes03+nMes04+nMes05+nMes06+nMes07+nMes08+nMes09+nMes10+nMes11+nMes12) " & _
'         "   From LogPlanAnualReqDetalle d inner join LogProSelBienesServicios b on d.cBSCod = b.cProSelBSCod " & _
'         "  where nEstado=1 and nPlanReqNro in (select x.nPlanReqNro from " & _
'         "       (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
'         "        Where r.nAnio = " & nAnio & " and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
'         "       (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro  Where x.nNro = Y.nApro) " & _
'         "  group by b.cProSelBSCod, b.cCIIUCod,convert(tinyint,substring(b.cProSelBSCod,2,1)), d.cBSGrupoCod, nMoneda, nPrecioUnitario "
'
'         Set Rs = oConn.CargaRecordSet(sSQL)
'         Do While Not Rs.EOF
'            nItem = nItem + 1
'
'            cCIIUCod = Rs!cCIIUCod
'            nObjetoCod = Rs!nObj
'            cBSCod = Rs!cProSelBSCod
'            cBSGrupoCod = Rs!cBSGrupoCod
'            nMoneda = Rs!nMoneda
'            nPrecioUnitario = Rs!nPrecioUnitario
'            nCantidad = Rs!nCantidad
'
'            sSQL = "insert into LogPlanAnualDetalleBS (nPlanAnualNro,nPlanAnualItem,cCIIUCod,nObjetoCod, cBSCod,cBSGrupoCod,nMoneda,nPrecioUnitario,nCantidad) " & _
'                   " values (" & nPlanAnualNro & "," & nItem & ", '" & cCIIUCod & "'," & nObjetoCod & ",'" & cBSCod & "', '" & cBSGrupoCod & "'," & nMoneda & "," & nPrecioUnitario & "," & nCantidad & ") "
'            oConn.Ejecutar sSQL
'
'            Rs.MoveNext
'         Loop


Sub ListaPlanAnual()
Dim oConn As New DConecta, rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim nAnio As Integer

nAnio = CInt(VNumero(txtAnio.Text))

sSQL = ""
nSuma = 0
FormaFlex

If oConn.AbreConexion Then
  
  sSQL = "SELECT d.nPlanAnualNro,d.nPlanAnualItem,d.nProSelTpoCod,d.nProSelSubTpo,d.nObjetoCod, o.cObjeto,d.nPlanAnualMes, d.cUbigeoCod, " & _
  "       d.cCIIU , d.cSintesis, d.nMoneda, d.nValorEstimado, r.cAbreviatura, f.cFuenteFinanciamiento " & _
  "  from LogPlanAnualDetalle d " & _
  "  left outer join (select nProSelTpoCod,nProSelSubTpo,cAbreviatura from LogProSelTpoRangos) r on r.nProSelTpoCod = d.nProSelTpoCod and r.nProSelSubTpo = d.nProSelSubTpo " & _
  "  left outer join (select nConsValor as nFuenteFinCod, cFuenteFinanciamiento=cConsDescripcion from Constante where nConsCod = 9046 and nConsCod<>nConsValor) f on d.nFuenteFinCod = f.nFuenteFinCod " & _
  "  left outer join (select nConsValor as nObjetoCod, cObjeto=cConsDescripcion from Constante where nConsCod = 9048 and nConsCod<>nConsValor) o on d.nObjetoCod = o.nObjetoCod  " & _
  " Where d.nPlanAnualAnio = " & nAnio & " And d.nPlanAnualEstado = 1"

   If Len(sSQL) = 0 Then Exit Sub
   
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         InsRow MSFlex, i
         MSFlex.RowHeight(i) = 500
         MSFlex.TextMatrix(i, 0) = rs!nPlanAnualNro
         MSFlex.TextMatrix(i, 1) = rs!nPlanAnualItem
         MSFlex.TextMatrix(i, 2) = rs!nProSelTpoCod
         MSFlex.TextMatrix(i, 3) = rs!nProSelSubTpo
         MSFlex.TextMatrix(i, 4) = rs!nObjetoCod
         MSFlex.TextMatrix(i, 5) = IIf(IsNull(rs!cAbreviatura), "", rs!cAbreviatura)
         MSFlex.TextMatrix(i, 6) = IIf(IsNull(rs!cObjeto), "", rs!cObjeto)
         MSFlex.TextMatrix(i, 7) = rs!cCIIU
         MSFlex.TextMatrix(i, 8) = rs!cSintesis
         If rs!nPlanAnualMes > 0 Then
            MSFlex.TextMatrix(i, 9) = UCase(mMatMes(rs!nPlanAnualMes))
         Else
            MSFlex.TextMatrix(i, 9) = ""
         End If
         MSFlex.TextMatrix(i, 10) = IIf(rs!nMoneda = 1, "SOLES", "DOLARES")
         MSFlex.TextMatrix(i, 11) = FNumero(rs!nValorEstimado)
         MSFlex.TextMatrix(i, 12) = rs!cFuenteFinanciamiento
         MSFlex.TextMatrix(i, 13) = GetUbigeoConsucode(rs!cUbigeoCod)
         rs.MoveNext
      Loop
   End If
End If
End Sub

Sub FormaFlex()
MSFlex.Clear
MSFlex.Rows = 2
MSFlex.RowHeight(0) = 420
MSFlex.RowHeight(1) = 8
MSFlex.ColWidth(0) = 0
MSFlex.ColWidth(1) = 400:    MSFlex.TextMatrix(0, 1) = "Nro":     MSFlex.ColAlignment(1) = 4
MSFlex.ColWidth(2) = 0
MSFlex.ColWidth(3) = 0
MSFlex.ColWidth(4) = 0:      MSFlex.TextMatrix(0, 4) = "Nro":     MSFlex.ColAlignment(4) = 4
MSFlex.ColWidth(5) = 600:    MSFlex.TextMatrix(0, 5) = "Tipo de Proceso"
MSFlex.ColWidth(6) = 1000:   MSFlex.TextMatrix(0, 6) = "Objeto"
MSFlex.ColWidth(7) = 600:    MSFlex.TextMatrix(0, 7) = " CIIU":    MSFlex.ColAlignment(7) = 4
MSFlex.ColWidth(8) = 3800:   MSFlex.TextMatrix(0, 8) = "Síntesis de Especificaciones Técnicas"
MSFlex.ColWidth(9) = 1200:   MSFlex.TextMatrix(0, 9) = "Fecha Probable de Convocatoria":    MSFlex.ColAlignment(9) = 4
MSFlex.ColWidth(10) = 900:   MSFlex.TextMatrix(0, 10) = "Moneda":    MSFlex.ColAlignment(10) = 4
MSFlex.ColWidth(11) = 1000:  MSFlex.TextMatrix(0, 11) = "Valor Estimado"
MSFlex.ColWidth(12) = 2500:  MSFlex.TextMatrix(0, 12) = "Fuente de Financiamiento" ' MSFlex.ColAlignment(12) = 4
MSFlex.ColWidth(13) = 3500:  MSFlex.TextMatrix(0, 13) = "Ubicación Geográfica" ' MSFlex.ColAlignment(12) = 4
MSFlex.WordWrap = True
End Sub

Private Sub MSFlex_KeyPress(KeyAscii As Integer)
Dim i As Integer
i = MSFlex.row
If KeyAscii = 13 And bEditable Then
   frmLogPlanAnualDetalle.Inicio MSFlex.TextMatrix(i, 0), MSFlex.TextMatrix(i, 1)
   If frmLogPlanAnualDetalle.vpHaGrabado Then
      DatosPlanAnual
   End If
End If
End Sub

Private Sub txtAnio_Change()
If Len(Trim(txtAnio)) = 4 Then
   ListaPlanAnual
End If
End Sub

Private Sub txtAnio_GotFocus()
SelTexto txtAnio
End Sub

