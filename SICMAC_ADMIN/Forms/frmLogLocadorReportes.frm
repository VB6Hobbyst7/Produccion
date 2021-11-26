VERSION 5.00
Begin VB.Form frmLogLocadorReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Locadores"
   ClientHeight    =   3315
   ClientLeft      =   2625
   ClientTop       =   2445
   ClientWidth     =   6075
   Icon            =   "frmLogLocadorReportes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4740
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar Reporte"
      Height          =   375
      Left            =   2940
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contrato "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5835
      Begin VB.TextBox txtFecFin 
         Height          =   315
         Left            =   3060
         TabIndex        =   9
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox txtFecIni 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Locación del Servicio "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   5835
      Begin VB.ComboBox cboAge 
         Height          =   315
         ItemData        =   "frmLogLocadorReportes.frx":08CA
         Left            =   1020
         List            =   "frmLogLocadorReportes.frx":08D1
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   4575
      End
      Begin VB.ComboBox cboArea 
         Height          =   315
         ItemData        =   "frmLogLocadorReportes.frx":08ED
         Left            =   1020
         List            =   "frmLogLocadorReportes.frx":08F4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   4575
      End
      Begin VB.ComboBox cboFuncion 
         Height          =   315
         ItemData        =   "frmLogLocadorReportes.frx":090D
         Left            =   1020
         List            =   "frmLogLocadorReportes.frx":0914
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Función"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1260
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmLogLocadorReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String

Private Sub cmdGenerar_Click()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta
Dim i As Integer
Dim appExcel As New Excel.Application
Dim wbExcel As Excel.Workbook
Dim cAgeCod As String, cCargoCod As String
Dim cAreaCod As String, cFuncionCod As String
Dim nFuncion As Integer
Dim cConsulta As String

Set wbExcel = appExcel.Workbooks.Add

wbExcel.Worksheets(1).Range("A1:HZ1000").Font.Size = 8
wbExcel.Worksheets(1).Range("A1").ColumnWidth = 12
wbExcel.Worksheets(1).Range("B1").ColumnWidth = 35
wbExcel.Worksheets(1).Range("C1").ColumnWidth = 20
wbExcel.Worksheets(1).Range("D1").ColumnWidth = 8
wbExcel.Worksheets(1).Range("E1").ColumnWidth = 8
wbExcel.Worksheets(1).Range("F1").ColumnWidth = 30
wbExcel.Worksheets(1).Range("G1").ColumnWidth = 30

wbExcel.Worksheets(1).Range("A3").value = "Codigo"
wbExcel.Worksheets(1).Range("B3").value = "Apellidos y Nombres"
wbExcel.Worksheets(1).Range("C3").value = "Nro Contrato"
wbExcel.Worksheets(1).Range("D3").value = "Desde"
wbExcel.Worksheets(1).Range("E3").value = "Hasta"
wbExcel.Worksheets(1).Range("F3").value = "Area"
wbExcel.Worksheets(1).Range("G3").value = "Agencia"
wbExcel.Worksheets(1).Range("A3:H3").Font.Bold = True

wbExcel.Worksheets(1).Range("D1:E1000").HorizontalAlignment = 3

i = 3

cAgeCod = Format(cboAge.ItemData(cboAge.ListIndex), "00")
cAreaCod = Format(cboArea.ItemData(cboArea.ListIndex), "000")
nFuncion = cboFuncion.ItemData(cboFuncion.ListIndex)

cConsulta = ""
If cboAge.ListIndex > 0 Then
   cConsulta = cConsulta + " x.cAgeCod = '" & cAgeCod & "' "
End If

If cboArea.ListIndex > 0 Then
   If Len(cConsulta) > 0 Then
      cConsulta = cConsulta + " AND x.cAreaCod = '" & cAreaCod & "' "
   Else
      cConsulta = cConsulta + " x.cAreaCod = '" & cAreaCod & "' "
   End If
End If

If cboFuncion.ListIndex > 0 Then
   If Len(cConsulta) > 0 Then
      cConsulta = cConsulta + " AND x.nFuncionCod = " & nFuncion & " "
   Else
      cConsulta = cConsulta + " x.nFuncionCod = " & nFuncion & " "
   End If
End If

If oConn.AbreConexion Then

   sSQL = "select x.nRegLocador, x.cPersCod, cPersona=replace(p.cPersNombre,'/',' '),cNroContrato,dFechaInicio,dFechaTermino, " & _
          " aa.cAreaDescripcion as cArea, a.cAgeDescripcion as cAgencia " & _
          " from LogLocadores x inner join Persona p on x.cPersCod=p.cPersCod " & _
          "            inner join Agencias a on x.cAgeCod = a.cAgeCod " & _
          "            inner join Areas aa on x.cAreaCod = aa.cAreaCod  " & _
          " Where x.nEstado = 1 and " & cConsulta & " order by cPersona"

   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         i = i + 1
         wbExcel.Worksheets(1).Range("A" + CStr(i)).value = rs!cPersCod
         wbExcel.Worksheets(1).Range("B" + CStr(i)).value = rs!cPersona
         wbExcel.Worksheets(1).Range("C" + CStr(i)).value = rs!cNroContrato
         wbExcel.Worksheets(1).Range("D" + CStr(i)).value = CStr(rs!dFechaInicio)
         wbExcel.Worksheets(1).Range("E" + CStr(i)).value = CStr(rs!dFechaTermino)
         wbExcel.Worksheets(1).Range("F" + CStr(i)).value = rs!cArea
         wbExcel.Worksheets(1).Range("G" + CStr(i)).value = rs!cAgencia
         rs.MoveNext
      Loop
   End If
End If

appExcel.Application.Visible = True
appExcel.Windows(1).Visible = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
txtFecIni = "01/01/" + CStr(Year(Date))
txtFecFin = Date
CargaCombos
End Sub

Sub CargaCombos()
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

If oConn.AbreConexion Then
   'cboAge.Clear
   sSQL = "Select cAgeCod,cAgeDescripcion from Agencias where nEstado = 1"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboAge.AddItem rs!cAgeDescripcion
         cboAge.ItemData(cboAge.ListCount - 1) = rs!cAgeCod
         rs.MoveNext
      Loop
      cboAge.ListIndex = 0
   End If
   
   'cboArea.Clear
   sSQL = "select cAreaCod, cAreaDescripcion from Areas order by cAreaDescripcion"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboArea.AddItem rs!cAreaDescripcion
         cboArea.ItemData(cboArea.ListCount - 1) = rs!cAreaCod
         rs.MoveNext
      Loop
      cboArea.ListIndex = 0
   End If
   
   'cboFuncion.Clear
   sSQL = "select nConsValor, cConsDescripcion from Constante where nConsCod = 9132 and nConsCod<>nConsValor"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         cboFuncion.AddItem rs!cConsDescripcion
         cboFuncion.ItemData(cboFuncion.ListCount - 1) = rs!nConsValor
         rs.MoveNext
      Loop
      cboFuncion.ListIndex = 0
   End If
   
End If
End Sub
