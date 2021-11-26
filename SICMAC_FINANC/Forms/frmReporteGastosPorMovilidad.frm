VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteGastosPorMovilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gastos por movilidad"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "frmReporteGastosPorMovilidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPeriodo 
      Caption         =   "Periodo"
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
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
      Height          =   1275
      Left            =   2040
      TabIndex        =   16
      Top             =   120
      Width           =   2880
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReporteGastosPorMovilidad.frx":030A
         Left            =   120
         List            =   "frmReporteGastosPorMovilidad.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.OptionButton optRangoFec 
      Caption         =   "Rango de Fechas"
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Frame FRAgencia 
      Caption         =   "Agencia"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   4815
      Begin VB.CheckBox chkTodasAgencias 
         Caption         =   "&Todas las Agencias"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin MSComctlLib.ListView lvAgencia 
         Height          =   2700
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   4763
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6174
         EndProperty
      End
   End
   Begin VB.Frame FRUsuario 
      Caption         =   "Usuario"
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
      TabIndex        =   10
      Top             =   4680
      Width           =   4815
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame FRTipo 
      Caption         =   "Tipo"
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
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   8160
      Visible         =   0   'False
      Width           =   4455
      Begin VB.OptionButton optTpo 
         Caption         =   "Detalle"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optTpo 
         Caption         =   "Resumen"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame fraFechaRango 
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
      Height          =   1275
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1905
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   300
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   795
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   435
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmReporteGastosPorMovilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************
'***Nombre:         frmReporteGastosPorMovilidad
'***Descripción:    Formulario que permite generar el Reporte de Gastos por Movilidad.
'***Creación:       ELRO el 20130117, según OYP-RFC126-2012
'************************************************************************************

Private fbNoTodo As Boolean

Private Function validarCampos() As Boolean
Dim lsMensaje As String
validarCampos = False

lsMensaje = ValidaFecha(txtFechaDel.Text)
If Trim(lsMensaje) <> "" Then
    MsgBox lsMensaje, vbInformation, "!Aviso¡"
    txtFechaDel.SetFocus
    Exit Function
End If

lsMensaje = ValidaFecha(txtFechaAl.Text)
If Trim(lsMensaje) <> "" Then
    MsgBox lsMensaje, vbInformation, "!Aviso¡"
    txtFechaAl.SetFocus
    Exit Function
End If

If CDate(txtFechaDel) > CDate(txtFechaAl) Then
    MsgBox "Fecha final debe ser mayor", vbInformation, "¡Aviso!"
    Exit Function
End If

If cboUsuario = "" Then
    MsgBox "Debe elegir un usuario o todos los usuarios", vbInformation, "¡Aviso!"
    Exit Function
End If

validarCampos = True
End Function

Private Sub CargarAgencias()
    Dim oNArendir As New NARendir
    Set oNArendir = New NARendir
    Dim rsAgencia As ADODB.Recordset
    Set rsAgencia = New ADODB.Recordset
    Dim lvItem As ListItem
    
    Set rsAgencia = oNArendir.devolverAgencias
    
    lvAgencia.ListItems.Clear
    Do While Not rsAgencia.EOF
        Set lvItem = lvAgencia.ListItems.Add
        lvItem.Text = rsAgencia.Fields(0)
        lvItem.SubItems(1) = rsAgencia.Fields(1)
        lvItem.Checked = False
        rsAgencia.MoveNext
    Loop

    Set rsAgencia = Nothing
    Set oNArendir = Nothing
End Sub

Private Sub CargarUsuarios()
    Dim oNArendir As New NARendir
    Set oNArendir = New NARendir
    Dim rsUsuario As ADODB.Recordset
    Set rsUsuario = New ADODB.Recordset
    Dim lvItem As ListItem
    
    Set rsUsuario = oNArendir.devolverUsuarios
     
    RSLlenaCombo rsUsuario, cboUsuario, , , False
    cboUsuario.AddItem "XXXX  TODOS LOS USUARIOS"
    RSClose rsUsuario
    Set oNArendir = Nothing
End Sub

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboUsuario.SetFocus
    End If
End Sub

Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub chkTodasAgencias_Click()
Dim k As Integer

If fbNoTodo = False Then
    If chkTodasAgencias Then
        For k = 1 To CInt(lvAgencia.ListItems.Count)
            DoEvents
            lvAgencia.ListItems(k).Selected = True
            lvAgencia.ListItems(k).Checked = True
            lvAgencia.SelectedItem.EnsureVisible
        Next k
     Else
        For k = 1 To CInt(lvAgencia.ListItems.Count)
                DoEvents
                lvAgencia.ListItems(k).Selected = True
                lvAgencia.ListItems(k).Checked = False
                lvAgencia.SelectedItem.EnsureVisible
        Next k
     End If
 End If
 
End Sub

Private Sub cmdAceptar_Click()
Dim lsUsuario As String
Dim lsAgencia As String
Dim ldFechaIni As String
Dim ldFechaFin As String
Dim nU As Integer
Dim k, l As Integer
Dim lnError As Integer
'*** PEAC 20131121
Dim pdFecha As Date
Dim lcFechaIni As String
Dim lcFechaFin As String
'*** FIN PEAC

ldFechaIni = txtFechaDel.Text
ldFechaFin = txtFechaAl.Text

If validarCampos = False Then Exit Sub


l = CInt(lvAgencia.ListItems.Count)

If chkTodasAgencias.value = False Then
    For k = 1 To l
        If lvAgencia.ListItems(k).Checked Then
            If Trim(lsAgencia) = "" Then
                lsAgencia = lvAgencia.ListItems(k).Text
            Else
                lsAgencia = lsAgencia & "," & Trim(lvAgencia.ListItems(k).Text)
            End If
        End If
        If k = l Then Exit For
     Next k
Else
    lsAgencia = ""
End If

If cboUsuario.ListIndex >= 0 Then
   If Left(cboUsuario, 4) = "XXXX" Then
      lsUsuario = ""
   Else
      lsUsuario = UCase(Left(Me.cboUsuario, 4))
   End If
End If

'*** PEAC 20131121
If Me.optPeriodo.value = True Then
    lcFechaIni = Left(Format(DateAdd("d", -1, DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(Me.txtAnio.Text, "0000"))), "yyyymmdd"), 6) + "01"
    lcFechaFin = Format(DateAdd("d", -1, DateAdd("m", 1, "01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(Me.txtAnio.Text, "0000"))), "yyyymmdd")
ElseIf Me.optRangoFec.value = True Then
    lcFechaIni = Format(ldFechaIni, "yyyyMMdd")
    lcFechaFin = Format(ldFechaFin, "yyyyMMdd")
Else
    MsgBox "Seleccione un filtro para las fechas.", vbOKOnly + vbInformation, "Atención"
    Exit Sub
End If

If Trim(lsAgencia) <> "" Or cboUsuario.ListIndex >= 0 Then
   'lnError = GeneraReporteGastosMovilidad(lsUsuario, lsAgencia, Format(ldFechaIni, "yyyyMMdd"), Format(ldFechaFin, "yyyyMMdd"))
   lnError = GeneraReporteGastosMovilidad(lsUsuario, lsAgencia, lcFechaIni, lcFechaFin) '*** PEAC 20131121
   If lnError = -1 Then Exit Sub
Else
   MsgBox "Seleccione un Agencia/Usuario", vbInformation, "¡Aviso!"
   lvAgencia.SetFocus
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargarAgencias
    CargarUsuarios
    txtFechaDel = gdFecSis
    txtFechaAl = gdFecSis
    
    cboUsuario = cboUsuario.List(cboUsuario.ListCount - 1)
    cboMes = cboMes.List(cboMes.ListIndex + Month(gdFecSis))
    Me.txtAnio = Year(gdFecSis)
    
End Sub

Private Function GeneraReporteGastosMovilidad(ByVal psUsuario As String, ByVal psAgencia As String, _
                                      ByVal pdFechaIni As String, ByVal pdFechaFin As String) As Integer
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo As String
    Dim lsArchivo1, lsArchivo2 As String
    Dim rsMovilidad As New ADODB.Recordset
    Dim oNArendir As NARendir
    Dim lnFilaIni, lnFilaFin As Integer
    Dim lsHoja As String
    Dim fs As scripting.FileSystemObject
    Dim lsNombreAgencia As String
    
    Set fs = New scripting.FileSystemObject
    Set xlAplicacion = New Excel.Application
    
    lsArchivo = "RPGMovilidad"
    lsArchivo1 = Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS")
    lsArchivo2 = "\SPOOLER\RPGMovilidad_" & lsArchivo1 & ".xls"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    
    Else
        MsgBox "No existe la plantilla RPGMovilidad.xls en la carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        GeneraReporteGastosMovilidad = -1
        Exit Function
    End If
    
    If chkTodasAgencias.value = 0 Then
        Set oNArendir = New NARendir
        lsNombreAgencia = oNArendir.devolverNombreAgenciasLineal(psAgencia)
        Set oNArendir = Nothing
    End If
        
    For Each xlHoja1 In xlLibro.Worksheets

    If UCase(xlHoja1.Name) = UCase(optTpo.Item(0).Caption) Then
        
'        xlHoja1.Cells(3, 2) = txtFechaDel & " - " & txtFechaAl
'        xlHoja1.Cells(4, 2) = IIf(Left(cboUsuario, 4) = "XXXX", "Todos", psUsuario)
'        xlHoja1.Cells(5, 2) = IIf(chkTodasAgencias, "Todas", lsNombreAgencia)
        
        xlHoja1.Cells(3, 2) = IIf(Me.optRangoFec.value = True, txtFechaDel & " - " & txtFechaAl, UCase(Me.cboMes.Text) & " - " & Format(Me.txtAnio.Text, "####"))
        
        lnFilaIni = 8
        lnFilaFin = 8
        
        If optTpo.Item(0).value = True Then
            xlHoja1.Activate
        End If

        Set oNArendir = New NARendir
        Set rsMovilidad = oNArendir.devolverGastoMovilidadResumen(pdFechaIni, pdFechaFin, psUsuario, psAgencia)
        Set oNArendir = Nothing

        If Not (rsMovilidad.BOF And rsMovilidad.EOF) Then
            Do While Not rsMovilidad.EOF
                xlHoja1.Cells(lnFilaFin, 1).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 1) = "'" + UCase(rsMovilidad!cFecha)
                xlHoja1.Range("B" & lnFilaFin & ":B" & lnFilaFin).NumberFormat = "#,###0.00"
                xlHoja1.Cells(lnFilaFin, 2).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 2) = UCase(rsMovilidad!nImporte)
                xlHoja1.Range("A" & lnFilaFin & ":B" & lnFilaFin).Font.Size = 8
                rsMovilidad.MoveNext
                lnFilaFin = lnFilaFin + 1
            Loop
                xlHoja1.Range("A" & lnFilaFin & ":B" & lnFilaFin).Font.Size = 8
                xlHoja1.Cells(lnFilaFin, 1).HorizontalAlignment = xlRight
                xlHoja1.Cells(lnFilaFin, 1).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 1).Font.Bold = True
                xlHoja1.Cells(lnFilaFin, 1) = "TOTAL"
                xlHoja1.Cells(lnFilaFin, 2).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 2).Font.Bold = True
                xlHoja1.Cells(lnFilaFin, 2) = "=SUM(B" & lnFilaIni & ":B" & lnFilaFin - 1 & ")"
        End If

    ElseIf UCase(xlHoja1.Name) = UCase(optTpo.Item(1).Caption) Then
    
'        xlHoja1.Cells(3, 2) = txtFechaDel & " - " & txtFechaAl
'        xlHoja1.Cells(4, 2) = IIf(Left(cboUsuario, 4) = "XXXX", "Todos", psUsuario)
'        xlHoja1.Cells(5, 2) = IIf(chkTodasAgencias, "Todas", lsNombreAgencia)

        xlHoja1.Cells(3, 2) = IIf(Me.optRangoFec.value = True, txtFechaDel & " - " & txtFechaAl, UCase(Me.cboMes.Text) & " - " & Format(Me.txtAnio.Text, "####"))

        lnFilaFin = 8
        If optTpo.Item(1).value = True Then
            xlHoja1.Activate
        End If

        Set rsMovilidad = Nothing
        Set oNArendir = New NARendir
        Set rsMovilidad = oNArendir.devolverGastoMovilidadDetalle(pdFechaIni, pdFechaFin, psUsuario, psAgencia)
        Set oNArendir = Nothing

        If Not (rsMovilidad.BOF And rsMovilidad.EOF) Then
            Do While Not rsMovilidad.EOF
                xlHoja1.Cells(lnFilaFin, 1).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 1) = "'" + UCase(rsMovilidad!cFecha)
                
                xlHoja1.Cells(lnFilaFin, 2).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 2) = rsMovilidad!cDocNro
                                
                xlHoja1.Cells(lnFilaFin, 3).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 3) = UCase(rsMovilidad!cPersNombre)
                
                xlHoja1.Cells(lnFilaFin, 4).Borders.LineStyle = 1
                'xlHoja1.Cells(lnFilaFin, 3) = UCase(rsMovilidad!cAgeDescripcion)
                xlHoja1.Cells(lnFilaFin, 4) = "'" + UCase(rsMovilidad!cAgecod)
                
                xlHoja1.Cells(lnFilaFin, 5).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 5) = "'" + UCase(rsMovilidad!cDni)
                
                xlHoja1.Cells(lnFilaFin, 6).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 6) = UCase(rsMovilidad!cLugar)
                
                xlHoja1.Cells(lnFilaFin, 7).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 7) = UCase(rsMovilidad!cMovDesc)
                
'                xlHoja1.Range("A" & lnFilaFin & ":H" & lnFilaFin).Font.Size = 8
'                xlHoja1.Range("A" & lnFilaFin & ":H" & lnFilaFin).RowHeight = 27
                
                xlHoja1.Range("A" & lnFilaFin & ":I" & lnFilaFin).Font.Size = 8
                xlHoja1.Range("A" & lnFilaFin & ":I" & lnFilaFin).RowHeight = 27
                
'                xlHoja1.Range("G" & lnFilaFin & ":G" & lnFilaFin).Borders.LineStyle = 1
'                xlHoja1.Range("G" & lnFilaFin & ":G" & lnFilaFin).NumberFormat = "#,###0.00"
                
                xlHoja1.Range("H" & lnFilaFin & ":H" & lnFilaFin).Borders.LineStyle = 1
                xlHoja1.Range("H" & lnFilaFin & ":H" & lnFilaFin).NumberFormat = "#,###0.00"
                
'                xlHoja1.Cells(lnFilaFin, 7) = UCase(rsMovilidad!nImporte)
'                xlHoja1.Cells(lnFilaFin, 8).Borders.LineStyle = 1
                
                xlHoja1.Cells(lnFilaFin, 8) = UCase(rsMovilidad!nImporte)
                xlHoja1.Cells(lnFilaFin, 9).Borders.LineStyle = 1
                
                rsMovilidad.MoveNext
                lnFilaFin = lnFilaFin + 1
            Loop
                'xlHoja1.Range("A" & lnFilaFin & ":G" & lnFilaFin).Font.Size = 8
                xlHoja1.Range("A" & lnFilaFin & ":H" & lnFilaFin).Font.Size = 8
                
'                xlHoja1.Range("A" & lnFilaFin & ":F" & lnFilaFin).Merge
'                xlHoja1.Range("A" & lnFilaFin & ":F" & lnFilaFin).Borders.LineStyle = 1
'                xlHoja1.Range("A" & lnFilaFin & ":F" & lnFilaFin).Font.Bold = True
'                xlHoja1.Range("A" & lnFilaFin & ":F" & lnFilaFin).HorizontalAlignment = xlRight
'                xlHoja1.Range("A" & lnFilaFin & ":F" & lnFilaFin) = "TOTAL"
                
                xlHoja1.Range("A" & lnFilaFin & ":G" & lnFilaFin).Merge
                xlHoja1.Range("A" & lnFilaFin & ":G" & lnFilaFin).Borders.LineStyle = 1
                xlHoja1.Range("A" & lnFilaFin & ":G" & lnFilaFin).Font.Bold = True
                xlHoja1.Range("A" & lnFilaFin & ":G" & lnFilaFin).HorizontalAlignment = xlRight
                xlHoja1.Range("A" & lnFilaFin & ":G" & lnFilaFin) = "TOTAL"
                
'                xlHoja1.Cells(lnFilaFin, 7).Borders.LineStyle = 1
'                xlHoja1.Cells(lnFilaFin, 7).Font.Bold = True
'                xlHoja1.Cells(lnFilaFin, 7) = "=SUM(G" & lnFilaIni & ":G" & lnFilaFin - 1 & ")"
                
                xlHoja1.Cells(lnFilaFin, 8).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 8).Font.Bold = True
                xlHoja1.Cells(lnFilaFin, 8) = "=SUM(H" & lnFilaIni & ":H" & lnFilaFin - 1 & ")"
                
        End If

    End If

    Next

    Set rsMovilidad = Nothing
    
    xlLibro.SaveAs (App.path & lsArchivo2)
    xlLibro.Close
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
    
    CargaArchivo lsArchivo & "_" & lsArchivo1 & ".xls", App.path & "\SPOOLER\"
    GeneraReporteGastosMovilidad = 1
    Exit Function
ErrImprime:
     MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
          xlLibro.Close
          xlAplicacion.Quit
       Set xlAplicacion = Nothing
       Set xlLibro = Nothing
       Set xlHoja1 = Nothing

End Function



Private Sub lvAgencia_Click()
Dim k, l As Integer
fbNoTodo = True
l = CInt(lvAgencia.ListItems.Count)

For k = 1 To l
    If lvAgencia.ListItems(k).Checked = False Then
       chkTodasAgencias.value = 0
       Exit For
    Else
        If k = l Then
            chkTodasAgencias.value = 1
            Exit For
        End If
    End If
Next k
fbNoTodo = False
End Sub

Private Sub txtAnio_GotFocus()
    fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Not ValidaAnio(txtAnio) Then
      Exit Sub
   End If
End If
End Sub

Private Sub txtAnio_Validate(Cancel As Boolean)
   If Not ValidaAnio(Val(txtAnio)) Then
      Cancel = True
   End If
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim lsMensaje As String
        lsMensaje = ValidaFecha(txtFechaAl.Text)
    
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "!Aviso¡"
            txtFechaAl.SetFocus
            Exit Sub
        ElseIf Trim(lsMensaje) = "" Then
            lvAgencia.SetFocus
        End If
    End If
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim lsMensaje As String
        lsMensaje = ValidaFecha(txtFechaDel.Text)
    
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "!Aviso¡"
            txtFechaDel.SetFocus
            Exit Sub
        ElseIf Trim(lsMensaje) = "" Then
            txtFechaAl.SetFocus
        End If
    End If
End Sub
