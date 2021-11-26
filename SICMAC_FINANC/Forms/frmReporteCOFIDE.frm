VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporteCOFIDE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte COFIDE"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   Icon            =   "frmReporteCOFIDE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame fraTpoDato 
      Caption         =   "Tipo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
      Begin VB.ComboBox cboTpoRep 
         Height          =   315
         ItemData        =   "frmReporteCOFIDE.frx":030A
         Left            =   240
         List            =   "frmReporteCOFIDE.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraRango 
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtAnio 
         Height          =   330
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   6
         Top             =   340
         Width           =   735
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmReporteCOFIDE.frx":0341
         Left            =   240
         List            =   "frmReporteCOFIDE.frx":0343
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblGuion 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmReporteCOFIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmReporteCOFIDE
'*** Descripción : Formulario para generar el Reporte COFIDE.
'*** Creación : MIOL el 20130614, según TI-ERS071-2013
'********************************************************************************
Option Explicit
Dim oRepCtaColumna As DRepCtaColumna

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGenerar_Click()
    If Me.cboMes.Text = "" Or Me.cboTpoRep.Text = "" Or Me.txtAnio.Text = "" Then
        MsgBox "Falta completar los datos; Verificar!"
        Exit Sub
    End If
    If Me.cboTpoRep.Text = "Colocaciones" Then
        generarMensualColocaciones
    ElseIf Me.cboTpoRep.Text = "Captaciones" Then
        generarMensualCaptaciones
    ElseIf Me.cboTpoRep.Text = "Adeudados" Then
        generarReporteAdeudados
    End If
End Sub

Private Sub Form_Load()
Inicializa_cbo
End Sub

Private Sub optMesTrim_Click(Index As Integer)
Inicializa_cbo
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub Inicializa_cbo()
cboMes.Clear
        cboMes.AddItem "Enero" & Space(200) & "01"
        cboMes.AddItem "Febrero" & Space(200) & "02"
        cboMes.AddItem "Marzo" & Space(200) & "03"
        cboMes.AddItem "Abril" & Space(200) & "04"
        cboMes.AddItem "Mayo" & Space(200) & "05"
        cboMes.AddItem "Junio" & Space(200) & "06"
        cboMes.AddItem "Julio" & Space(200) & "07"
        cboMes.AddItem "Agosto" & Space(200) & "08"
        cboMes.AddItem "Setiembre" & Space(200) & "09"
        cboMes.AddItem "Octubre" & Space(200) & "10"
        cboMes.AddItem "Noviembre" & Space(200) & "11"
        cboMes.AddItem "Diciembre" & Space(200) & "12"
End Sub

Private Sub generarMensualColocaciones()
    Me.MousePointer = vbHourglass
        Dim sPathColMor As String
        Dim sMesAnio As String
        Dim fs As New Scripting.FileSystemObject
        Dim obj_excel As Object, Libro As Object, Hoja As Object
        Dim convert As Double
        PB1.Min = 0
        PB1.Max = 46
        PB1.value = 0
        PB1.Visible = True
        On Error GoTo error_sub
        sPathColMor = App.path & "\Spooler\COL_MOR_" + Trim(Right(cboMes.Text, 2)) + "_" + Me.txtAnio.Text + ".xlsx"
        If fs.FileExists(sPathColMor) Then
            If ArchivoEstaAbierto(sPathColMor) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathColMor) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
            fs.DeleteFile sPathColMor, True
        End If
        sPathColMor = App.path & "\FormatoCarta\COL_MOR.xlsx"
        If Len(Dir(sPathColMor)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathColMor, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
        End If
        Set obj_excel = CreateObject("Excel.Application")
        obj_excel.DisplayAlerts = False
        Set Libro = obj_excel.Workbooks.Open(sPathColMor)
        Set Hoja = Libro.ActiveSheet
        Dim celda As Excel.Range
        Set oRepCtaColumna = New DRepCtaColumna
        Dim rsCtaColumna As ADODB.Recordset
        PB1.value = 1
        Set celda = obj_excel.Range("14A!G1") '****************************CARGAR DATOS 14A*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14A!G35")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14A!G68")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2)) 'VIGENTES
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15A(sMesAnio, 1)
        Dim nFilas As Integer
        nFilas = 3
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("14A!B" & nFilas)
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("14A!C" & nFilas)
               celda.value = rsCtaColumna(1)
               Set celda = obj_excel.Range("14A!D" & nFilas)
               celda.value = rsCtaColumna(2)
               Set celda = obj_excel.Range("14A!E" & nFilas)
               celda.value = rsCtaColumna(3)
               Set celda = obj_excel.Range("14A!F" & nFilas)
               celda.value = rsCtaColumna(4)
               Set celda = obj_excel.Range("14A!G" & nFilas)
               celda.value = rsCtaColumna(5)
               Set celda = obj_excel.Range("14A!H" & nFilas) 'NAGL
               celda.value = rsCtaColumna(6)
               Set celda = obj_excel.Range("14A!I" & nFilas) 'NAGL
               celda.value = rsCtaColumna(7)
               Set celda = obj_excel.Range("14A!J" & nFilas) 'NAGL
               celda.value = rsCtaColumna(8)
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 2
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15A(sMesAnio, 2) 'ATRASADOS
        nFilas = 37
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("14A!B" & nFilas)
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("14A!C" & nFilas)
               celda.value = rsCtaColumna(1)
               Set celda = obj_excel.Range("14A!D" & nFilas)
               celda.value = rsCtaColumna(2)
               Set celda = obj_excel.Range("14A!E" & nFilas)
               celda.value = rsCtaColumna(3)
               Set celda = obj_excel.Range("14A!F" & nFilas)
               celda.value = rsCtaColumna(4)
               Set celda = obj_excel.Range("14A!G" & nFilas)
               celda.value = rsCtaColumna(5)
               Set celda = obj_excel.Range("14A!H" & nFilas) 'NAGL
               celda.value = rsCtaColumna(6)
               Set celda = obj_excel.Range("14A!I" & nFilas) 'NAGL
               celda.value = rsCtaColumna(7)
               Set celda = obj_excel.Range("14A!J" & nFilas) 'NAGL
               celda.value = rsCtaColumna(8)
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
         PB1.value = 3
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15A(sMesAnio, 3) 'REFINANCIADOS
        nFilas = 70
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("14A!B" & nFilas)
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("14A!C" & nFilas)
               celda.value = rsCtaColumna(1)
               Set celda = obj_excel.Range("14A!D" & nFilas)
               celda.value = rsCtaColumna(2)
               Set celda = obj_excel.Range("14A!E" & nFilas)
               celda.value = rsCtaColumna(3)
               Set celda = obj_excel.Range("14A!F" & nFilas)
               celda.value = rsCtaColumna(4)
               Set celda = obj_excel.Range("14A!G" & nFilas)
               celda.value = rsCtaColumna(5)
               Set celda = obj_excel.Range("14A!H" & nFilas) 'NAGL
               celda.value = rsCtaColumna(6)
               Set celda = obj_excel.Range("14A!I" & nFilas) 'NAGL
               celda.value = rsCtaColumna(7)
               Set celda = obj_excel.Range("14A!J" & nFilas) 'NAGL
               celda.value = rsCtaColumna(8)
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 4
        cargarDatos15ACuadroConsCarteraBrutaRefAtrasxCredito obj_excel, "0"   'Cuadro Consolidado NAGL 20170417
        PB1.value = 5
        Set celda = obj_excel.Range("14B!C3") '****************************CARGAR DATOS 14B*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14B!C23")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14B!C43")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        PB1.value = 6
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15B(sMesAnio, 1) 'BRUTA
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C4")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C5")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 250
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C6")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C7")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 350
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C8")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C9")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 450
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C10")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C11")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 550
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C12")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C13")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 650
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C14")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C15")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 750
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C16")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C17")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 850
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C18")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C19")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 7
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15B(sMesAnio, 2) 'REFINANCIADOS
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C24")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C25")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 250
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C26")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C27")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 350
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C28")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C29")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 450
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C30")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C31")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 550
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C32")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C33")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 650
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C34")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C35")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 750
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C36")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C37")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 850
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C38")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C39")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
         PB1.value = 8
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15B(sMesAnio, 3) 'ATRASADOS
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C44")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C45")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 250
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C46")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C47")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 350
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C48")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C49")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 450
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C50")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C51")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 550
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C52")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C53")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 650
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C54")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C55")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 750
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C56")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C57")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
                Case 850
                    If rsCtaColumna!cConsDescripcion = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C58")
                        celda.value = rsCtaColumna!nMontoMes
                    Else
                        Set celda = obj_excel.Range("14B!C59")
                        celda.value = rsCtaColumna!nMontoMes
                    End If
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        
        PB1.value = 9
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15B(sMesAnio, 4) 'NRO DE DEUDORES - NAGL
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C65")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C66")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 250
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C67")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C68")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 350
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C69")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C70")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 450
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C71")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C72")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 550
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C73")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C74")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 650
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C75")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C76")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 750
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C77")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C78")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 850
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C79")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C80")
                        celda.value = rsCtaColumna!nCantidad
                    End If
                Case 755
                    If rsCtaColumna!Moneda = "SOLES" Then
                        Set celda = obj_excel.Range("14B!C81")
                        celda.value = rsCtaColumna!nCantidad
                    Else
                        Set celda = obj_excel.Range("14B!C82")
                        celda.value = rsCtaColumna!nCantidad
                    End If
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If 'Fin Comentario NAGL
    
        Set rsCtaColumna = Nothing
        PB1.value = 10
        Set celda = obj_excel.Range("14C!B3") '****************************CARGAR DATOS 14C*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14C!B14")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14C!B25")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        PB1.value = 11
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15C(sMesAnio, 1) 'BRUTO
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!nConsValor
                    Case 1
                        Set celda = obj_excel.Range("14C!B4")
                        celda.value = rsCtaColumna(2)
                    Case 2
                        Set celda = obj_excel.Range("14C!B5")
                        celda.value = rsCtaColumna(2)
                    Case 3
                        Set celda = obj_excel.Range("14C!B6")
                        celda.value = rsCtaColumna(2)
                    Case 4
                        Set celda = obj_excel.Range("14C!B7")
                        celda.value = rsCtaColumna(2)
                    Case 7
                        Set celda = obj_excel.Range("14C!B8")
                        celda.value = rsCtaColumna(2)
                    Case 8
                        Set celda = obj_excel.Range("14C!B9")
                        celda.value = rsCtaColumna(2)
                    Case 10
                        Set celda = obj_excel.Range("14C!B10")
                        celda.value = rsCtaColumna(2)
                End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 12
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15C(sMesAnio, 2) 'REFINANCIADO
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!nConsValor
                    Case 1
                        Set celda = obj_excel.Range("14C!B15")
                        celda.value = rsCtaColumna(2)
                    Case 2
                        Set celda = obj_excel.Range("14C!B16")
                        celda.value = rsCtaColumna(2)
                    Case 3
                        Set celda = obj_excel.Range("14C!B17")
                        celda.value = rsCtaColumna(2)
                    Case 4
                        Set celda = obj_excel.Range("14C!B18")
                        celda.value = rsCtaColumna(2)
                    Case 7
                        Set celda = obj_excel.Range("14C!B19")
                        celda.value = rsCtaColumna(2)
                    Case 8
                        Set celda = obj_excel.Range("14C!B20")
                        celda.value = rsCtaColumna(2)
                    Case 10
                        Set celda = obj_excel.Range("14C!B21")
                        celda.value = rsCtaColumna(2)
                End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 13
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15C(sMesAnio, 3) 'ATRASADO
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!nConsValor
                    Case 1
                        Set celda = obj_excel.Range("14C!B26")
                        celda.value = rsCtaColumna(2)
                    Case 2
                        Set celda = obj_excel.Range("14C!B27")
                        celda.value = rsCtaColumna(2)
                    Case 3
                        Set celda = obj_excel.Range("14C!B28")
                        celda.value = rsCtaColumna(2)
                    Case 4
                        Set celda = obj_excel.Range("14C!B29")
                        celda.value = rsCtaColumna(2)
                    Case 7
                        Set celda = obj_excel.Range("14C!B30")
                        celda.value = rsCtaColumna(2)
                    Case 8
                        Set celda = obj_excel.Range("14C!B31")
                        celda.value = rsCtaColumna(2)
                    Case 10
                        Set celda = obj_excel.Range("14C!B32")
                        celda.value = rsCtaColumna(2)
                End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 14
        Set celda = obj_excel.Range("14D!B3") '****************************CARGAR DATOS 14D*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14D!B24")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14D!B45")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        PB1.value = 15
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15D(sMesAnio)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
               Dim ContadorA As Integer
               Dim ContadorB As Integer
               Dim ContadorC As Integer
               Dim contador As Integer

               ContadorA = 4
               ContadorB = 25
               ContadorC = 46
               contador = 0
           Do While Not rsCtaColumna.EOF
           Set celda = obj_excel.Range("14D!B" & ContadorA)
            celda.value = rsCtaColumna(2)
            Set celda = obj_excel.Range("14D!B" & ContadorB)
            celda.value = rsCtaColumna(3)
            Set celda = obj_excel.Range("14D!B" & ContadorC)
            celda.value = rsCtaColumna(4)
            ContadorA = ContadorA + 1
            ContadorB = ContadorB + 1
            ContadorC = ContadorC + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 16
        Set celda = obj_excel.Range("14E!B4") '****************************CARGAR DATOS 14E*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14E!B38")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14E!B72")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        PB1.value = 17
        'LUCV20170125 -> Agregó y Comentó, según correo: RABE
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15E(sMesAnio)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!nConsValor
                    Case 501
                        Set celda = obj_excel.Range("14E!B5")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B39")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B73")
                        celda.value = rsCtaColumna(4)
                    Case 502
                        Set celda = obj_excel.Range("14E!B6")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B40")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B74")
                        celda.value = rsCtaColumna(4)
                    Case 503
                        Set celda = obj_excel.Range("14E!B7")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B41")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B75")
                        celda.value = rsCtaColumna(4)
                    Case 504
                        Set celda = obj_excel.Range("14E!B8")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B42")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B76")
                        celda.value = rsCtaColumna(4)
                    Case 505
                        Set celda = obj_excel.Range("14E!B9")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B43")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B77")
                        celda.value = rsCtaColumna(4)
                    Case 506
                        Set celda = obj_excel.Range("14E!B10")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B44")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B78")
                        celda.value = rsCtaColumna(4)
                    Case 507
                        Set celda = obj_excel.Range("14E!B11")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B45")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B79")
                        celda.value = rsCtaColumna(4)
                    Case 508
                        Set celda = obj_excel.Range("14E!B12")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B46")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B80")
                        celda.value = rsCtaColumna(4)
                    Case 509
                        Set celda = obj_excel.Range("14E!B13")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B47")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B81")
                        celda.value = rsCtaColumna(4)
                    Case 510
                        Set celda = obj_excel.Range("14E!B14")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B48")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B82")
                        celda.value = rsCtaColumna(4)
                    Case 512
                        Set celda = obj_excel.Range("14E!B15")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B49")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B83")
                        celda.value = rsCtaColumna(4)
                    Case 513
                        Set celda = obj_excel.Range("14E!B16")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B50")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B84")
                        celda.value = rsCtaColumna(4)
                    Case 515
                        Set celda = obj_excel.Range("14E!B17")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B51")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B85")
                        celda.value = rsCtaColumna(4)
                    Case 517
                        Set celda = obj_excel.Range("14E!B18")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B52")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B86")
                        celda.value = rsCtaColumna(4)
                    Case 518 'LUCV20170125, Agregó
                        Set celda = obj_excel.Range("14E!B19")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B53")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B87")
                        celda.value = rsCtaColumna(4) '***Fin LUCV20170125
                    Case 601
                        Set celda = obj_excel.Range("14E!B20")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B54")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B88")
                        celda.value = rsCtaColumna(4)
                    Case 602
                        Set celda = obj_excel.Range("14E!B21")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B55")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B89")
                        celda.value = rsCtaColumna(4)

                    Case 701
                        Set celda = obj_excel.Range("14E!B22")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B56")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B90")
                        celda.value = rsCtaColumna(4)
                    Case 702
                        Set celda = obj_excel.Range("14E!B23")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B57")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B91")
                        celda.value = rsCtaColumna(4)
                    Case 703
                        Set celda = obj_excel.Range("14E!B24")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B58")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B92")
                        celda.value = rsCtaColumna(4)
                    Case 704
                        Set celda = obj_excel.Range("14E!B25")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B59")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B93")
                        celda.value = rsCtaColumna(4)
                    Case 705
                        Set celda = obj_excel.Range("14E!B26")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B60")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B94")
                        celda.value = rsCtaColumna(4)
                    Case 706
                        Set celda = obj_excel.Range("14E!B27")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B61")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B95")
                        celda.value = rsCtaColumna(4)
                    Case 718
                        Set celda = obj_excel.Range("14E!B28")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B62")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B96")
                        celda.value = rsCtaColumna(4)
                    Case 801
                        Set celda = obj_excel.Range("14E!B29")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B63")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B97")
                        celda.value = rsCtaColumna(4)
                    Case 802
                        Set celda = obj_excel.Range("14E!B30")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B64")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B98")
                        celda.value = rsCtaColumna(4)
                    Case 803
                        Set celda = obj_excel.Range("14E!B31")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B65")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B99")
                        celda.value = rsCtaColumna(4)
                    Case 804
                        Set celda = obj_excel.Range("14E!B32")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B66")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B100")
                        celda.value = rsCtaColumna(4)
                   Case 805 '**LUCV20170125->, Agregó según correo: RABE
                        Set celda = obj_excel.Range("14E!B33")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B67")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B101")
                        celda.value = rsCtaColumna(4)
                    Case 806
                        Set celda = obj_excel.Range("14E!B34")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14E!B68")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14E!B102")
                        celda.value = rsCtaColumna(4) '**Fin LUCV20170125<-
               End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
'Fin Comentario LUCV20170125
        PB1.value = 18
        Set celda = obj_excel.Range("14F!B3") '****************************CARGAR DATOS 14F*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14F!B16")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14F!B29")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        PB1.value = 19
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15F(sMesAnio)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!Item
                    Case 1
                        Set celda = obj_excel.Range("14F!B4")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B17")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B30")
                        celda.value = rsCtaColumna(4)
                    Case 2
                        Set celda = obj_excel.Range("14F!B5")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B18")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B31")
                        celda.value = rsCtaColumna(4)
                    Case 3
                        Set celda = obj_excel.Range("14F!B6")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B19")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B32")
                        celda.value = rsCtaColumna(4)
                    Case 4
                        Set celda = obj_excel.Range("14F!B7")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B20")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B33")
                        celda.value = rsCtaColumna(4)
                    Case 5
                        Set celda = obj_excel.Range("14F!B8")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B21")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B34")
                        celda.value = rsCtaColumna(4)
                    Case 6
                        Set celda = obj_excel.Range("14F!B9")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B22")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B35")
                        celda.value = rsCtaColumna(4)
                    Case 7
                        Set celda = obj_excel.Range("14F!B10")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B23")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B36")
                        celda.value = rsCtaColumna(4)
                    Case 8
                        Set celda = obj_excel.Range("14F!B11")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B24")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B37")
                        celda.value = rsCtaColumna(4)
                    Case 9
                        Set celda = obj_excel.Range("14F!B12")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14F!B25")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14F!B38")
                        celda.value = rsCtaColumna(4)
               End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 20
        Set celda = obj_excel.Range("14G!B3") '****************************CARGAR DATOS 14G*******************************
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14G!B17")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14G!B31")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        PB1.value = 21
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15G(sMesAnio)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!Item
                        Case 1
                                Set celda = obj_excel.Range("14G!B4")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B18")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B32")
                                celda.value = rsCtaColumna(4)
                        Case 2
                                Set celda = obj_excel.Range("14G!B5")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B19")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B33")
                                celda.value = rsCtaColumna(4)
                        Case 3
                                Set celda = obj_excel.Range("14G!B6")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B20")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B34")
                                celda.value = rsCtaColumna(4)
                        Case 4
                                Set celda = obj_excel.Range("14G!B7")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B21")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B35")
                                celda.value = rsCtaColumna(4)
                        Case 5
                                Set celda = obj_excel.Range("14G!B8")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B22")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B36")
                                celda.value = rsCtaColumna(4)
                        Case 6
                                Set celda = obj_excel.Range("14G!B9")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B23")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B37")
                                celda.value = rsCtaColumna(4)
                        Case 7
                                Set celda = obj_excel.Range("14G!B10")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B24")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B38")
                                celda.value = rsCtaColumna(4)
                        Case 8
                                Set celda = obj_excel.Range("14G!B11")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B25")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B39")
                                celda.value = rsCtaColumna(4)
                        Case 9
                                Set celda = obj_excel.Range("14G!B12")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B26")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B40")
                                celda.value = rsCtaColumna(4)
                        Case 10
                                Set celda = obj_excel.Range("14G!B13")
                                celda.value = rsCtaColumna(2)
                                Set celda = obj_excel.Range("14G!B27")
                                celda.value = rsCtaColumna(3)
                                Set celda = obj_excel.Range("14G!B41")
                                celda.value = rsCtaColumna(4)
               End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 22
        Set celda = obj_excel.Range("14H!A1") '****************************CARGAR DATOS 14H*******************************
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B4")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B17")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B29")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B42")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B55")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B67")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B80")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14H!B92")
        celda.value = Trim(Left(cboMes.Text, 10)) + "/" + Me.txtAnio.Text
        PB1.value = 23
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 1, 8)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B6")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C6")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D6")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E6")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B7")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C7")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D7")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E7")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B8")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C8")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D8")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E8")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B9")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C9")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D9")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E9")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B10")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C10")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D10")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E10")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B11")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C11")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D11")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E11")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B12")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C12")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D12")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E12")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B13")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C13")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D13")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E13")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 24
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 9, 15)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B19")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C19")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D19")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E19")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B20")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C20")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D20")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E20")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B21")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C21")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D21")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E21")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B22")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C22")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D22")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E22")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B23")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C23")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D23")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E23")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B24")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C24")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D24")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E24")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B25")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C25")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D25")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E25")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B26")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C26")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D26")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E26")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 25
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 16, 30)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B31")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C31")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D31")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E31")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B32")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C32")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D32")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E32")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B33")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C33")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D33")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E33")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B34")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C34")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D34")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E34")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B35")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C35")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D35")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E35")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B36")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C36")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D36")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E36")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B37")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C37")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D37")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E37")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B38")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C38")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D38")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E38")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 26
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 31, 60)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B44")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C44")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D44")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E44")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B45")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C45")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D45")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E45")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B46")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C46")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D46")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E46")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B47")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C47")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D47")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E47")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B48")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C48")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D48")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E48")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B49")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C49")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D49")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E49")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B50")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C50")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D50")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E50")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B51")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C51")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D51")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E51")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 27
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 61, 120)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B57")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C57")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D57")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E57")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B58")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C58")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D58")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E58")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B59")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C59")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D59")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E59")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B60")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C60")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D60")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E60")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B61")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C61")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D61")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E61")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B62")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C62")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D62")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E62")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B63")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C63")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D63")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E63")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B64")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C64")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D64")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E64")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 28
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 121, 180)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B69")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C69")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D69")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E69")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B70")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C70")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D70")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E70")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B71")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C71")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D71")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E71")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B72")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C72")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D72")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E72")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B73")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C73")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D73")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E73")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B74")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C74")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D74")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E74")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B75")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C75")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D75")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E75")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B76")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C76")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D76")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E76")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 29
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 181, 361)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B82")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C82")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D82")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E82")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B83")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C83")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D83")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E83")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B84")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C84")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D84")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E84")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B85")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C85")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D85")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E85")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B86")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C86")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D86")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E86")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B87")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C87")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D87")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E87")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B88")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C88")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D88")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E88")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B89")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C89")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D89")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E89")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 30
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15H(sMesAnio, 360, 999999)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14H!B94")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C94")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D94")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E94")
                        celda.value = rsCtaColumna(5)
                Case 250
                        Set celda = obj_excel.Range("14H!B95")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C95")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D95")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E95")
                        celda.value = rsCtaColumna(5)
                Case 350
                        Set celda = obj_excel.Range("14H!B96")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C96")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D96")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E96")
                        celda.value = rsCtaColumna(5)
                Case 450
                        Set celda = obj_excel.Range("14H!B97")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C97")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D97")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E97")
                        celda.value = rsCtaColumna(5)
                Case 550
                        Set celda = obj_excel.Range("14H!B98")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C98")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D98")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E98")
                        celda.value = rsCtaColumna(5)
                Case 650
                        Set celda = obj_excel.Range("14H!B99")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C99")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D99")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E99")
                        celda.value = rsCtaColumna(5)
                Case 750
                        Set celda = obj_excel.Range("14H!B100")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C100")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D100")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E100")
                        celda.value = rsCtaColumna(5)
                Case 850
                        Set celda = obj_excel.Range("14H!B101")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14H!C101")
                        celda.value = rsCtaColumna(3)
                        Set celda = obj_excel.Range("14H!D101")
                        celda.value = rsCtaColumna(4)
                        Set celda = obj_excel.Range("14H!E101")
                        celda.value = rsCtaColumna(5)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 31
        Set celda = obj_excel.Range("14I!C3") '****************************CARGAR DATOS 14I*******************************
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        PB1.value = 32
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15I(sMesAnio, "1")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14I!C5")
                        celda.value = rsCtaColumna(2)
                Case 250
                        Set celda = obj_excel.Range("14I!C6")
                        celda.value = rsCtaColumna(2)
                Case 350
                        Set celda = obj_excel.Range("14I!C7")
                        celda.value = rsCtaColumna(2)
                Case 450
                        Set celda = obj_excel.Range("14I!C8")
                        celda.value = rsCtaColumna(2)
                Case 550
                        Set celda = obj_excel.Range("14I!C9")
                        celda.value = rsCtaColumna(2)
                Case 650
                        Set celda = obj_excel.Range("14I!C10")
                        celda.value = rsCtaColumna(2)
                Case 750
                        Set celda = obj_excel.Range("14I!C11")
                        celda.value = rsCtaColumna(2)
                Case 850
                        Set celda = obj_excel.Range("14I!C12")
                        celda.value = rsCtaColumna(2)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 33
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15I(sMesAnio, "2")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14I!D5")
                        celda.value = rsCtaColumna(2)
                Case 250
                        Set celda = obj_excel.Range("14I!D6")
                        celda.value = rsCtaColumna(2)
                Case 350
                        Set celda = obj_excel.Range("14I!D7")
                        celda.value = rsCtaColumna(2)
                Case 450
                        Set celda = obj_excel.Range("14I!D8")
                        celda.value = rsCtaColumna(2)
                Case 550
                        Set celda = obj_excel.Range("14I!D9")
                        celda.value = rsCtaColumna(2)
                Case 650
                        Set celda = obj_excel.Range("14I!D10")
                        celda.value = rsCtaColumna(2)
                Case 750
                        Set celda = obj_excel.Range("14I!D11")
                        celda.value = rsCtaColumna(2)
                Case 850
                        Set celda = obj_excel.Range("14I!D12")
                        celda.value = rsCtaColumna(2)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 34
        Set celda = obj_excel.Range("14J!C3") '****************************CARGAR DATOS 15J*******************************
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14J!C14")
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        PB1.value = 35
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15J(sMesAnio, "0")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14J!C5")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D5")
                        celda.value = rsCtaColumna(3)
                Case 250
                        Set celda = obj_excel.Range("14J!C6")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D6")
                        celda.value = rsCtaColumna(3)
                Case 350
                        Set celda = obj_excel.Range("14J!C7")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D7")
                        celda.value = rsCtaColumna(3)
                Case 450
                        Set celda = obj_excel.Range("14J!C8")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D8")
                        celda.value = rsCtaColumna(3)
                Case 550
                        Set celda = obj_excel.Range("14J!C9")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D9")
                        celda.value = rsCtaColumna(3)
                Case 650
                        Set celda = obj_excel.Range("14J!C10")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D10")
                        celda.value = rsCtaColumna(3)
                Case 750
                        Set celda = obj_excel.Range("14J!C11")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D11")
                        celda.value = rsCtaColumna(3)
                Case 850
                        Set celda = obj_excel.Range("14J!C12")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14J!D12")
                        celda.value = rsCtaColumna(3)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 36
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15J(sMesAnio, "1")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14J!C16")
                        celda.value = rsCtaColumna(2)
                Case 250
                        Set celda = obj_excel.Range("14J!C17")
                        celda.value = rsCtaColumna(2)
                Case 350
                        Set celda = obj_excel.Range("14J!C18")
                        celda.value = rsCtaColumna(2)
                Case 450
                        Set celda = obj_excel.Range("14J!C19")
                        celda.value = rsCtaColumna(2)
                Case 550
                        Set celda = obj_excel.Range("14J!C20")
                        celda.value = rsCtaColumna(2)
                Case 650
                        Set celda = obj_excel.Range("14J!C21")
                        celda.value = rsCtaColumna(2)
                Case 750
                        Set celda = obj_excel.Range("14J!C22")
                        celda.value = rsCtaColumna(2)
                Case 850
                        Set celda = obj_excel.Range("14J!C23")
                        celda.value = rsCtaColumna(2)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 37
        
        cargarDatosColocacionesMorPromxFec15J obj_excel, "2" 'Incluido en un Método by NAGL
        PB1.value = 38
        cargarDatos15JCredCastigxCred obj_excel, "3" 'Creditos Castigados x Tipo de Cred - NAGL
        PB1.value = 39
        Set celda = obj_excel.Range("14K!B3") '****************************CARGAR DATOS 15K*******************************
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14K!F3")
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        PB1.value = 40
        cargarDatos15KSumatoria obj_excel
        PB1.value = 41
        cargarDatos15KPromedio obj_excel, "1"
        PB1.value = 42
        cargarDatos15KPromedio obj_excel, "2"
        PB1.value = 43
        cargarDatos15KSaldoxSituacionCont obj_excel 'Saldo Cartera segun situación Contable NAGL
        PB1.value = 44
        cargarDatos15KSaldoxCategoriaRiesgo obj_excel ' Saldo Cartera según Categoria de Riesgo NAGL
        PB1.value = 45
        cargarDatos15LCarteraBrutaRefAtrasxDepartamento obj_excel '************************CARGAR DATOS 15L***************************** NAGL
        PB1.value = 46
        sPathColMor = App.path & "\Spooler\COL_MOR_" + Trim(Right(cboMes.Text, 2)) + "_" + Me.txtAnio.Text + ".xlsx" 'verifica si existe el archivo
        If fs.FileExists(sPathColMor) Then
            If ArchivoEstaAbierto(sPathColMor) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathColMor)
            End If
            fs.DeleteFile sPathColMor, True
        End If
        Hoja.SaveAs sPathColMor 'guarda el archivo
        Libro.Close
        obj_excel.Quit
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_excel = Nothing
        Me.MousePointer = vbDefault
        Dim m_excel As New Excel.Application 'abre y muestra el archivo
        m_excel.Workbooks.Open (sPathColMor)
        m_excel.Visible = True
        'PB1.value = 47
        PB1.Visible = False
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        PB1.Visible = False
        Me.MousePointer = vbDefault
End Sub

Private Sub cargarDatos15ACuadroConsCarteraBrutaRefAtrasxCredito(ByVal obj_excel As Excel.Application, ByVal psTipoMon As String)
Dim sMesAnio As String
Dim celda As Excel.Range
Set oRepCtaColumna = New DRepCtaColumna
Dim rsCtaColumna As ADODB.Recordset
Dim nFilas As Integer
    sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
     Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15A(sMesAnio, "0")
     
     nFilas = 4
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("14A!P" & nFilas)
               celda.value = rsCtaColumna(2)
               Set celda = obj_excel.Range("14A!Q" & nFilas)
               celda.value = rsCtaColumna(3)
               Set celda = obj_excel.Range("14A!R" & nFilas)
               celda.value = rsCtaColumna(4)
               Set celda = obj_excel.Range("14A!S" & nFilas)
               celda.value = rsCtaColumna(5)
               Set celda = obj_excel.Range("14A!T" & nFilas)
               celda.value = rsCtaColumna(6)
               Set celda = obj_excel.Range("14A!U" & nFilas)
               celda.value = rsCtaColumna(7)
               Set celda = obj_excel.Range("14A!V" & nFilas)
               celda.value = rsCtaColumna(8)
               Set celda = obj_excel.Range("14A!W" & nFilas)
               celda.value = rsCtaColumna(9)
               Set celda = obj_excel.Range("14A!X" & nFilas)
               celda.value = rsCtaColumna(10)
               Set celda = obj_excel.Range("14A!Y" & nFilas)
               celda.value = rsCtaColumna(11)
               Set celda = obj_excel.Range("14A!Z" & nFilas)
               celda.value = rsCtaColumna(12)
               Set celda = obj_excel.Range("14A!AA" & nFilas)
               celda.value = rsCtaColumna(13)
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
     Set rsCtaColumna = Nothing
End Sub 'NAGL 20170417

Private Sub cargarDatosColocacionesMorPromxFec15J(ByVal obj_excel As Excel.Application, ByVal psTipoMon As String) 'NAGL
    Dim sMesAnio As String
    Dim celda As Excel.Range
       Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
    
    sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
     Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15J(sMesAnio, "2")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14J!D16")
                        celda.value = rsCtaColumna(2)
                Case 250
                        Set celda = obj_excel.Range("14J!D17")
                        celda.value = rsCtaColumna(2)
                Case 350
                        Set celda = obj_excel.Range("14J!D18")
                        celda.value = rsCtaColumna(2)
                Case 450
                        Set celda = obj_excel.Range("14J!D19")
                        celda.value = rsCtaColumna(2)
                Case 550
                        Set celda = obj_excel.Range("14J!D20")
                        celda.value = rsCtaColumna(2)
                Case 650
                        Set celda = obj_excel.Range("14J!D21")
                        celda.value = rsCtaColumna(2)
                Case 750
                        Set celda = obj_excel.Range("14J!D22")
                        celda.value = rsCtaColumna(2)
                Case 850
                        Set celda = obj_excel.Range("14J!D23")
                        celda.value = rsCtaColumna(2)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
End Sub

Private Sub cargarDatos15JCredCastigxCred(ByVal obj_excel As Excel.Application, ByVal psTipoMon As String) 'NAGL
    Dim sMesAnio As String
    Dim celda As Excel.Range
       Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
        
        Set celda = obj_excel.Range("14J!C26")
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15J(sMesAnio, "3") 'Creditos Castigados x Tipo de Cred - NAGL
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C28")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D28")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 250
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C29")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D29")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 350
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C30")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D30")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 450
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C31")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D31")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 550
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C32")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D32")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 650
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C33")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D33")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 750
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C34")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D34")
                            celda.value = rsCtaColumna(3)
                        End If
                Case 850
                        If rsCtaColumna!Moneda = "SOLES" Then
                            Set celda = obj_excel.Range("14J!C35")
                            celda.value = rsCtaColumna(3)
                        Else
                            Set celda = obj_excel.Range("14J!D35")
                            celda.value = rsCtaColumna(3)
                        End If
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If 'fin Conentario
        Set rsCtaColumna = Nothing
End Sub '**NAGL

Private Sub cargarDatos15KSumatoria(ByVal obj_excel As Excel.Application)
    Dim sMesAnio As String
    Dim celda As Excel.Range
       Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
 
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15K(sMesAnio, "0")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        Set celda = obj_excel.Range("14K!B5")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C5")
                        celda.value = rsCtaColumna(3)
                Case 250
                        Set celda = obj_excel.Range("14K!B6")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C6")
                        celda.value = rsCtaColumna(3)
                Case 350
                        Set celda = obj_excel.Range("14K!B7")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C7")
                        celda.value = rsCtaColumna(3)
                Case 450
                        Set celda = obj_excel.Range("14K!B8")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C8")
                        celda.value = rsCtaColumna(3)
                Case 550
                        Set celda = obj_excel.Range("14K!B9")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C9")
                        celda.value = rsCtaColumna(3)
                Case 650
                        Set celda = obj_excel.Range("14K!B10")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C10")
                        celda.value = rsCtaColumna(3)
                Case 750
                        Set celda = obj_excel.Range("14K!B11")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C11")
                        celda.value = rsCtaColumna(3)
                Case 850
                        Set celda = obj_excel.Range("14K!B12")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C12")
                        celda.value = rsCtaColumna(3)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
End Sub

Private Sub cargarDatos15KPromedio(ByVal obj_excel As Excel.Application, ByVal psMoneda As String)
    Dim sMesAnio As String
    Dim celda As Excel.Range
       Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
 
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15K(sMesAnio, psMoneda)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 150
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F5")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G5")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 250
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F6")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G6")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 350
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F7")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G7")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 450
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F8")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G8")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 550
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F9")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G9")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 650
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F10")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G10")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 750
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F11")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G11")
                            celda.value = rsCtaColumna(2)
                        End If
                Case 850
                        If psMoneda = "1" Then
                            Set celda = obj_excel.Range("14K!F12")
                            celda.value = rsCtaColumna(2)
                        Else
                            Set celda = obj_excel.Range("14K!G12")
                            celda.value = rsCtaColumna(2)
                        End If
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        
End Sub
Private Sub cargarDatos15KSaldoxSituacionCont(ByVal obj_excel As Excel.Application) '***NAGL
    Dim sMesAnio As String
    Dim celda As Excel.Range
       Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
    
        Set celda = obj_excel.Range("14K!B17")
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15K(sMesAnio, "3")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!cSitCtb
                Case 1
                        Set celda = obj_excel.Range("14K!B19")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C19")
                        celda.value = rsCtaColumna(3)
                Case 2
                        Set celda = obj_excel.Range("14K!B20")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C20")
                        celda.value = rsCtaColumna(3)
                Case 3
                        Set celda = obj_excel.Range("14K!B21")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C21")
                        celda.value = rsCtaColumna(3)
                Case 4
                        Set celda = obj_excel.Range("14K!B22")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!C22")
                        celda.value = rsCtaColumna(3)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
End Sub '****NAGL
Private Sub cargarDatos15KSaldoxCategoriaRiesgo(ByVal obj_excel As Excel.Application) '***NAGL
    Dim sMesAnio As String
    Dim celda As Excel.Range
       Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
    
        Set celda = obj_excel.Range("14K!F16")
        celda.value = Trim(Left(cboMes.Text, 10)) + "-" + Me.txtAnio.Text
        
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15K(sMesAnio, "4")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
            Select Case rsCtaColumna!nConsValor
                Case 0
                        Set celda = obj_excel.Range("14K!F18")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!G18")
                        celda.value = rsCtaColumna(3)
                Case 1
                        Set celda = obj_excel.Range("14K!F19")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!G19")
                        celda.value = rsCtaColumna(3)
                Case 2
                        Set celda = obj_excel.Range("14K!F20")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!G20")
                        celda.value = rsCtaColumna(3)
                Case 3
                        Set celda = obj_excel.Range("14K!F21")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!G21")
                        celda.value = rsCtaColumna(3)
                Case 4
                        Set celda = obj_excel.Range("14K!F22")
                        celda.value = rsCtaColumna(2)
                        Set celda = obj_excel.Range("14K!G22")
                        celda.value = rsCtaColumna(3)
            End Select
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
End Sub '****NAGL
Private Sub cargarDatos15LCarteraBrutaRefAtrasxDepartamento(ByVal obj_excel As Excel.Application) '****NAGL
    Dim sMesAnio As String
    Dim celda As Excel.Range
    Dim Contador1 As Integer
    Dim Contador2 As Integer
    Dim Contador3 As Integer
    Set oRepCtaColumna = New DRepCtaColumna
    Dim rsCtaColumna As ADODB.Recordset
    
        Set celda = obj_excel.Range("14L!B4")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14L!B34")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        Set celda = obj_excel.Range("14L!B64")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
     
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        
        Set rsCtaColumna = oRepCtaColumna.GetColocacionesMorxFec15L(sMesAnio)
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
                Contador1 = 4 + CInt(rsCtaColumna!cUbiGeoCod)
                Contador2 = 34 + CInt(rsCtaColumna!cUbiGeoCod)
                Contador3 = 64 + CInt(rsCtaColumna!cUbiGeoCod)
           Set celda = obj_excel.Range("14L!B" & Contador1)
            celda.value = rsCtaColumna(2)
            Set celda = obj_excel.Range("14L!B" & Contador2)
            celda.value = rsCtaColumna(3)
            Set celda = obj_excel.Range("14L!B" & Contador3)
            celda.value = rsCtaColumna(4)
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
End Sub 'NAGL 20170425

Private Sub generarMensualCaptaciones()
    Me.MousePointer = vbHourglass
        Dim sPathCapMor As String
        Dim sMesAnio As String
        
        Dim fs As New Scripting.FileSystemObject
        Dim obj_excel As Object, Libro As Object, Hoja As Object
        
        Dim convert As Double
        
        PB1.Min = 0
        PB1.Max = 10
        PB1.value = 0
        PB1.Visible = True
        On Error GoTo error_sub
          
        sPathCapMor = App.path & "\Spooler\CAP_MOR_" + Trim(Right(cboMes.Text, 2)) + "_" + Me.txtAnio.Text + ".xls"
        
        If fs.FileExists(sPathCapMor) Then
            
            If ArchivoEstaAbierto(sPathCapMor) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathCapMor) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
    
            fs.DeleteFile sPathCapMor, True
        End If
        
        sPathCapMor = App.path & "\FormatoCarta\CAP_MOR.xls"

        If Len(Dir(sPathCapMor)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathCapMor, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
        End If
        
        Set obj_excel = CreateObject("Excel.Application")
        obj_excel.DisplayAlerts = False
        Set Libro = obj_excel.Workbooks.Open(sPathCapMor)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
         Set oRepCtaColumna = New DRepCtaColumna
        Dim rsCtaColumna As ADODB.Recordset
        
        PB1.value = 1
        '****************************CARGAR DATOS 16A*******************************
        Set celda = obj_excel.Range("16A!C1")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        
        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16A(sMesAnio)
        Dim nFilas As Integer
        nFilas = 3
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("16A!B" & nFilas)
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("16A!C" & nFilas)
               celda.value = rsCtaColumna(1)
               Set celda = obj_excel.Range("16A!D" & nFilas)
               celda.value = rsCtaColumna(2)
               Set celda = obj_excel.Range("16A!E" & nFilas)
               celda.value = rsCtaColumna(3)
               Set celda = obj_excel.Range("16A!F" & nFilas)
               celda.value = rsCtaColumna(4)
               Set celda = obj_excel.Range("16A!G" & nFilas)
               celda.value = rsCtaColumna(5)
               Set celda = obj_excel.Range("16A!H" & nFilas)
               celda.value = rsCtaColumna(6)
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 2
        
        '****************************CARGAR DATOS 16B*******************************
        Set celda = obj_excel.Range("16B!C4")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        
        'PLAZO FIJO
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16BxPF(sMesAnio, "1")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               If rsCtaColumna!nNroMes = 1 Then
                    Set celda = obj_excel.Range("16B!D17")
                    celda.value = rsCtaColumna(1)
                    Set celda = obj_excel.Range("16B!D53")
                    celda.value = rsCtaColumna(1)
                    Set celda = obj_excel.Range("16C!B19")
                    celda.value = rsCtaColumna(1)
               End If
               Select Case rsCtaColumna!nNroMes
                    Case 1
                            Set celda = obj_excel.Range("16B!C24")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D24")
                            celda.value = rsCtaColumna(3)
                    Case 2
                            Set celda = obj_excel.Range("16B!C25")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D25")
                            celda.value = rsCtaColumna(3)
                    Case 3
                            Set celda = obj_excel.Range("16B!C26")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D26")
                            celda.value = rsCtaColumna(3)
                    Case 6
                            Set celda = obj_excel.Range("16B!C27")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D27")
                            celda.value = rsCtaColumna(3)
                    Case 12
                            Set celda = obj_excel.Range("16B!C28")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D28")
                            celda.value = rsCtaColumna(3)
                    Case 24
                            Set celda = obj_excel.Range("16B!C29")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D29")
                            celda.value = rsCtaColumna(3)
                    Case 36
                            Set celda = obj_excel.Range("16B!C30")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D30")
                            celda.value = rsCtaColumna(3)
                    Case 48
                            Set celda = obj_excel.Range("16B!C31")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D31")
                            celda.value = rsCtaColumna(3)
                    Case 60
                            Set celda = obj_excel.Range("16B!C32")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D32")
                            celda.value = rsCtaColumna(3)
                    Case 61
                            Set celda = obj_excel.Range("16B!C33")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D33")
                            celda.value = rsCtaColumna(3)
                End Select
            rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 3
        
        Set celda = obj_excel.Range("16B!C22")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16BxPF(sMesAnio, "2")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Select Case rsCtaColumna!nNroMes
                    Case 1
                            Set celda = obj_excel.Range("16B!C42")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D42")
                            celda.value = rsCtaColumna(3)
                    Case 2
                            Set celda = obj_excel.Range("16B!C43")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D43")
                            celda.value = rsCtaColumna(3)
                    Case 3
                            Set celda = obj_excel.Range("16B!C44")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D44")
                            celda.value = rsCtaColumna(3)
                    Case 6
                            Set celda = obj_excel.Range("16B!C45")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D45")
                            celda.value = rsCtaColumna(3)
                    Case 12
                            Set celda = obj_excel.Range("16B!C46")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D46")
                            celda.value = rsCtaColumna(3)
                    Case 24
                            Set celda = obj_excel.Range("16B!C47")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D47")
                            celda.value = rsCtaColumna(3)
                    Case 36
                            Set celda = obj_excel.Range("16B!C48")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D48")
                            celda.value = rsCtaColumna(3)
                    Case 48
                            Set celda = obj_excel.Range("16B!C49")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D49")
                            celda.value = rsCtaColumna(3)
                    Case 60
                            Set celda = obj_excel.Range("16B!C50")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D50")
                            celda.value = rsCtaColumna(3)
                    Case 61
                            Set celda = obj_excel.Range("16B!C51")
                            celda.value = rsCtaColumna(2)
                            Set celda = obj_excel.Range("16B!D51")
                            celda.value = rsCtaColumna(3)
                End Select
            rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 4
        
        'AHORROS
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16BxAH(sMesAnio, "1")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("16B!C60")
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("16B!D60")
               celda.value = rsCtaColumna(1)
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 5
        
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16BxAH(sMesAnio, "2")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("16B!C69")
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("16B!D69")
               celda.value = rsCtaColumna(1)
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 6
        
        'CTS
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16BxCTS(sMesAnio, "1")
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("16B!C80")
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("16B!D80")
               celda.value = rsCtaColumna(1)
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 7
        
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16BxCTS(sMesAnio, "2")
        If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("16B!C89")
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("16B!D89")
               celda.value = rsCtaColumna(1)
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 8
        
        '****************************CARGAR DATOS 16C*******************************
        Set celda = obj_excel.Range("16C!B3")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        
        Set rsCtaColumna = oRepCtaColumna.GetCaptacionesMorxFec16C(sMesAnio)
        nFilas = 4
        If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
                Set celda = obj_excel.Range("16C!A" & nFilas)
                celda.value = rsCtaColumna(1)
                Set celda = obj_excel.Range("16C!B" & nFilas)
                celda.value = rsCtaColumna(2)
               nFilas = nFilas + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
        PB1.value = 9
        'verifica si existe el archivo
        sPathCapMor = App.path & "\Spooler\CAP_MOR_" + Trim(Right(cboMes.Text, 2)) + "_" + Me.txtAnio.Text + ".xls"
        If fs.FileExists(sPathCapMor) Then
            
            If ArchivoEstaAbierto(sPathCapMor) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathCapMor)
            End If
            fs.DeleteFile sPathCapMor, True
        End If
        'guarda el archivo
        Hoja.SaveAs sPathCapMor

        Libro.Close
        obj_excel.Quit
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_excel = Nothing
        Me.MousePointer = vbDefault
        'abre y muestra el archivo
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathCapMor)
        m_excel.Visible = True
        PB1.value = 10
        PB1.Visible = False
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        PB1.Visible = False
        Me.MousePointer = vbDefault
End Sub
        
'MIOL 20130710,SEGUN RQ13331 ******************************************
Private Sub generarReporteAdeudados()
    Me.MousePointer = vbHourglass
        Dim sPathAdeudado As String
        Dim sMesAnio As String
        
        Dim fs As New Scripting.FileSystemObject
        Dim obj_excel As Object, Libro As Object, Hoja As Object
        
        Dim convert As Double
        
        PB1.Min = 0
        PB1.Max = 4
        PB1.value = 0
        PB1.Visible = True
        On Error GoTo error_sub
          
        sPathAdeudado = App.path & "\Spooler\ADEU_" + Trim(Right(cboMes.Text, 2)) + "_" + Me.txtAnio.Text + ".xls"
        
        If fs.FileExists(sPathAdeudado) Then
            
            If ArchivoEstaAbierto(sPathAdeudado) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAdeudado) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
    
            fs.DeleteFile sPathAdeudado, True
        End If
        
        sPathAdeudado = App.path & "\FormatoCarta\ADEU.xls"

        If Len(Dir(sPathAdeudado)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathAdeudado, vbCritical
           Me.MousePointer = vbDefault
           Exit Sub
        End If
        
        Set obj_excel = CreateObject("Excel.Application")
        obj_excel.DisplayAlerts = False
        Set Libro = obj_excel.Workbooks.Open(sPathAdeudado)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
         Set oRepCtaColumna = New DRepCtaColumna
        Dim rsCtaColumna As ADODB.Recordset
        
        PB1.value = 1
        '************************CARGAR DATOS ADEUDADOS*************************
        Set celda = obj_excel.Range("Adeudados!C1")
        celda.value = Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        
        Dim nTpoCam  As Double
        Dim sFecTpoCambio As String
        sFecTpoCambio = "25" + "/" + Trim(Right(cboMes.Text, 2)) + "/" + Me.txtAnio.Text
        nTpoCam = LeeTpoCambio(sFecTpoCambio, TCFijoDia)
        
        Set celda = obj_excel.Range("Adeudados!J1")
        celda.value = nTpoCam

        sMesAnio = Me.txtAnio.Text + Trim(Right(cboMes.Text, 2))
        Set rsCtaColumna = oRepCtaColumna.GetAdeudados(sMesAnio)
        Dim nFilas As Integer
        Dim nItem As Integer
        Dim nMontoSol As Currency
        Dim nMontoDol As Currency
        nFilas = 3
        nItem = 1
          If Not rsCtaColumna.EOF Or rsCtaColumna.BOF Then
           Do While Not rsCtaColumna.EOF
               Set celda = obj_excel.Range("Adeudados!A" & nFilas)
               celda.value = nItem
               Set celda = obj_excel.Range("Adeudados!B" & nFilas)
               celda.value = rsCtaColumna(0)
               Set celda = obj_excel.Range("Adeudados!C" & nFilas)
               celda.value = rsCtaColumna(1)
                If rsCtaColumna(1) = "Soles" Then
                    nMontoSol = nMontoSol + rsCtaColumna(2)
                    Set celda = obj_excel.Range("Adeudados!M3")
                    celda.value = nMontoSol
                ElseIf rsCtaColumna(1) = "Dolares" Then
                    nMontoDol = nMontoDol + rsCtaColumna(2)
                    Set celda = obj_excel.Range("Adeudados!M4")
                    celda.value = nMontoDol
                End If
               Set celda = obj_excel.Range("Adeudados!D" & nFilas)
               celda.value = Format(rsCtaColumna(2), "#,###0.00")
               Set celda = obj_excel.Range("Adeudados!E" & nFilas)
               celda.value = Format(rsCtaColumna(3), "#,###0.00")
               Set celda = obj_excel.Range("Adeudados!F" & nFilas)
               celda.value = Format(rsCtaColumna(4), "#,###0.00")
               Set celda = obj_excel.Range("Adeudados!G" & nFilas)
               celda.value = rsCtaColumna(5)
               Set celda = obj_excel.Range("Adeudados!H" & nFilas)
               celda.value = rsCtaColumna(6)
               Set celda = obj_excel.Range("Adeudados!I" & nFilas)
               celda.value = rsCtaColumna(7)
               nFilas = nFilas + 1
               nItem = nItem + 1
               rsCtaColumna.MoveNext
           Loop
        End If
        Set rsCtaColumna = Nothing
       
        PB1.value = 3
        'Verifica si existe el archivo
        sPathAdeudado = App.path & "\Spooler\ADEU_" + Trim(Right(cboMes.Text, 2)) + "_" + Me.txtAnio.Text + ".xls"
        If fs.FileExists(sPathAdeudado) Then
            
            If ArchivoEstaAbierto(sPathAdeudado) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathAdeudado)
            End If
            fs.DeleteFile sPathAdeudado, True
        End If
        'Guarda el archivo
        Hoja.SaveAs sPathAdeudado

        Libro.Close
        obj_excel.Quit
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_excel = Nothing
        Me.MousePointer = vbDefault
        'Abre y Muestra el archivo
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathAdeudado)
        m_excel.Visible = True
        PB1.value = 4
        PB1.Visible = False
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        PB1.Visible = False
        Me.MousePointer = vbDefault
End Sub
'END MIOL *************************************************************

Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function


