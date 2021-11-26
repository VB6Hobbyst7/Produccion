VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOperEgresosEfectivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones: Egresos en Efectivo Agencia"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frmOperEgresosEfectivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   225
      Left            =   1575
      TabIndex        =   11
      Top             =   4250
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   4620
      Top             =   6405
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComCtl2.MonthView dHasta 
      Height          =   2370
      Left            =   5880
      TabIndex        =   14
      Top             =   5145
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   71237633
      CurrentDate     =   41460
   End
   Begin MSComCtl2.MonthView dDesde 
      Height          =   2370
      Left            =   945
      TabIndex        =   13
      Top             =   5145
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   71237633
      CurrentDate     =   41460
   End
   Begin SICMACT.FlexEdit grdOperaciones 
      Height          =   3060
      Left            =   210
      TabIndex        =   9
      Top             =   945
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   5398
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Usuario-Fecha-Hora-Operacion-Moneda-Importe-Glosa"
      EncabezadosAnchos=   "500-1200-800-1000-1000-1000-1000-1000-3000"
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
      EncabezadosAlineacion=   "C-L-C-C-C-C-C-C-L"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8190
      TabIndex        =   12
      Top             =   4200
      Width           =   1170
   End
   Begin VB.CommandButton cmdHasta 
      Caption         =   ">>"
      Height          =   330
      Left            =   7470
      TabIndex        =   7
      Top             =   420
      Width           =   330
   End
   Begin VB.CommandButton cmdDesde 
      Caption         =   ">>"
      Height          =   330
      Left            =   4635
      TabIndex        =   5
      Top             =   420
      Width           =   330
   End
   Begin MSMask.MaskEdBox txtHasta 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   3
      EndProperty
      Height          =   330
      Left            =   5565
      TabIndex        =   6
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm dd-mmm-yy"
      Mask            =   "##:## ##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDesde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   3
      EndProperty
      Height          =   330
      Left            =   2730
      TabIndex        =   4
      Top             =   420
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm dd/mmm/yy"
      Mask            =   "##:## ##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8085
      TabIndex        =   8
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   10
      Top             =   4200
      Width           =   1170
   End
   Begin VB.ComboBox cboAgencia 
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
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   2115
   End
   Begin VB.Label lblBuscando 
      Caption         =   "Buscando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   4095
      TabIndex        =   15
      Top             =   5565
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      Height          =   225
      Left            =   5565
      TabIndex        =   2
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label2 
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
      Height          =   225
      Left            =   2730
      TabIndex        =   1
      Top             =   105
      Width           =   750
   End
   Begin VB.Label Label1 
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
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   1170
   End
End
Attribute VB_Name = "frmOperEgresosEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************************************
'*** NOMBRE      : frmOperEgresosEfectivo
'*** DESCRIPCION : Formulario creado con la finalidad de buscar operaciones realizadas por uno o mas usuarios
'***               con parametros de agencia, y fecha, tambien permite exportar el resultado a un archivo de excel.
'***
'*** CREACION    : RIRO 20130702 SEGUN TI-ERS083-2013
'*************************************************************************************************************************
Option Explicit
Public Sub Inicia()
    
    'Actual Formulario
    Me.Width = 9645
    Me.Height = 5070
    
    'Calendario Desde
    dDesde.Left = 2355
    dDesde.Top = 735
    
    'Calendario Hasta
    dHasta.Left = 5205
    dHasta.Top = 735
    
    'Label Buscando
    lblBuscando.Visible = False
    lblBuscando.Left = 3930
    lblBuscando.Top = 4200
    
    'Barra de Progreso
    pBar.Visible = False
    pBar.Left = 1575
    pBar.Top = 4250
    
    Dim a As COMDCaptaServicios.DCOMCaptaServicios
    Dim b As ADODB.Recordset
    Dim i As Integer
    
    Set a = New COMDCaptaServicios.DCOMCaptaServicios
    
    Set b = a.GetAgencia
     
    IniciaCombo cboAgencia, b
    
    dDesde.value = gdFecSis
    dHasta.value = gdFecSis
    
    If Trim(gsCodCargo) = "006005" Then
        For i = 0 To cboAgencia.ListCount - 1
            If Trim(Right(cboAgencia.List(i), 5)) = gsCodAge Then
                cboAgencia.ListIndex = i
                i = cboAgencia.ListCount - 1
            End If
        Next
        cboAgencia.Enabled = False
    End If
    
        
    Me.Show 1
    
End Sub

Public Sub IniciaCombo(ByRef combo As ComboBox, rs As ADODB.Recordset)
    Dim Campo As ADODB.Field
    Dim lsDato As String
    Dim sAgen As String
        
    If rs Is Nothing Then Exit Sub
    combo.Clear
    Do While Not rs.EOF
        lsDato = ""
        lsDato = lsDato & rs("cAgeDescripcion") & Space(70) & rs("cAgeCod")
        combo.AddItem lsDato
        
        sAgen = sAgen & rs("cAgeCod") & ","
        rs.MoveNext
    Loop
    rs.Close
    sAgen = Mid(sAgen, 1, Len(sAgen) - 1)
    combo.AddItem "Todos" & Space(70) & sAgen
    Set rs = Nothing
End Sub

Private Sub cmdBuscar_Click()
 
    Dim sAgencia, sDesde, sHasta As String
    Dim oCajImp As COMNCajaGeneral.NCOMCajero
    Dim rsReporte As ADODB.Recordset
    
    dDesde.Visible = False
    dHasta.Visible = False
    grdOperaciones.Clear
    grdOperaciones.Rows = 2
    grdOperaciones.FormaCabecera
    
    lblBuscando.Visible = True
    DoEvents
    
    Set oCajImp = New COMNCajaGeneral.NCOMCajero
    
    If Trim(Left(cboAgencia.Text, 10)) = "Todos" Then
        sAgencia = Trim(Right(cboAgencia.Text, Len(cboAgencia.Text) - 70))
    
    Else
        sAgencia = Trim(Right(cboAgencia.Text, 5))
    
    End If
    
    sDesde = Format(Trim(txtDesde.Text), "yyyymmddhhmm")
    sHasta = Format(Trim(txtHasta.Text), "yyyymmddhhmm")
        
    Set rsReporte = oCajImp.ReporteEgresos(sAgencia, sDesde, sHasta)
    
    If Not rsReporte.EOF And Not rsReporte.BOF Then
        grdOperaciones.rsFlex = rsReporte
        DoEvents
    Else
        
    End If
    
    lblBuscando.Visible = False
    
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdDesde_Click()
    dHasta.Visible = False
    If dDesde.Visible Then
        dDesde.Visible = False
    Else
        dDesde.Visible = True
    End If
End Sub

Private Sub cmdExportar_Click()

    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsArchivo1 As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim nFila, i As Double
    Dim nI As Long
On Error GoTo error_handler
    
    dDesde.Visible = False
    dHasta.Visible = False
    If grdOperaciones.Rows = 2 And Trim(grdOperaciones.TextMatrix(1, 1)) = "" Then
        MsgBox "No hay resultados a exportar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    dlgArchivo.FileName = Empty
    dlgArchivo.Filter = "Archivo *.xls" & "|*.xls"
    dlgArchivo.FileName = Year(gdFecSis) & Month(gdFecSis) & Day(gdFecSis) & ".xls"
    dlgArchivo.ShowSave
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    If fs.FileExists(dlgArchivo.FileName) Then
        
        MsgBox "El archivo '" & dlgArchivo.FileTitle & "' ya existe, debe asignarle un nombre diferente", vbExclamation, ""
        Exit Sub
    
    End If
    
    pBar.Visible = True
    pBar.Max = grdOperaciones.Rows
    pBar.Min = 0
    pBar.value = 0
    
    cmdCerrar.Cancel = False
    lsArchivo = "FormatoEgresosEfectivo"
    lsNomHoja = "Egresos"
    nFila = 3
    

           
    lsArchivo1 = dlgArchivo.FileName
    

    
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    
    For Each xlHoja1 In xlsLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
    
    For nI = 1 To grdOperaciones.Rows - 1
       nFila = nFila + 1
       xlHoja1.Cells(nFila, 2) = grdOperaciones.TextMatrix(nI, 1)
       xlHoja1.Cells(nFila, 3) = grdOperaciones.TextMatrix(nI, 2)
       xlHoja1.Cells(nFila, 4) = grdOperaciones.TextMatrix(nI, 3)
       xlHoja1.Cells(nFila, 5) = grdOperaciones.TextMatrix(nI, 4)
       xlHoja1.Cells(nFila, 6) = grdOperaciones.TextMatrix(nI, 5)
       xlHoja1.Cells(nFila, 7) = grdOperaciones.TextMatrix(nI, 6)
       xlHoja1.Cells(nFila, 8) = grdOperaciones.TextMatrix(nI, 7)
       xlHoja1.Cells(nFila, 9) = Trim(Replace(grdOperaciones.TextMatrix(nI, 8), vbNewLine, " "))
       If pBar.value + 1 <= pBar.Max Then
         pBar.value = pBar.value + 1
       End If
       DoEvents
     Next
     
    pBar.Visible = False
    xlHoja1.SaveAs lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
    grdOperaciones.Clear
    grdOperaciones.Rows = 2
    grdOperaciones.FormaCabecera
    cmdCerrar.Cancel = True
    
    Limpiar
    MsgBox "Se culminó el proceso con exito", vbInformation, "Aviso"
               
Exit Sub
    
error_handler:
    
    If err.Number = 32755 Then
        MsgBox "Se ha cancelado formulario", vbInformation, "Aviso"
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        pBar.Visible = False
        cmdCerrar.Cancel = True
        'Set ClsServicioRecaudo = Nothing
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Error al momento de generar el archivo", vbCritical, "Aviso"
    End If

End Sub

Private Sub cmdHasta_Click()
    dDesde.Visible = False
    If dHasta.Visible Then
        dHasta.Visible = False
    Else
        dHasta.Visible = True
    End If
End Sub

Private Sub dDesde_DateClick(ByVal DateClicked As Date)
    txtDesde.Text = Format(Now, "hh:mm ") & Format(DateClicked, "dd/mm/yyyy")
    dDesde.Visible = False
End Sub

Private Sub dHasta_DateClick(ByVal DateClicked As Date)
    txtHasta.Text = Format(Now, "hh:mm ") & Format(DateClicked, "dd/mm/yyyy")
    dHasta.Visible = False
End Sub

Private Sub Limpiar()

txtDesde.Text = "__:__ __/__/____"
txtHasta.Text = "__:__ __/__/____"
dDesde.value = gdFecSis
dHasta.value = gdFecSis
txtDesde.SetFocus

If Trim(gsCodCargo) <> "006005" Then
   cboAgencia.Enabled = True
   cboAgencia.ListIndex = 0
End If

End Sub

