VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmNIIFFormaB2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma B2"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "frmNIIFFormaB2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar BarraProgreso 
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2640
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Expresado en ..."
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
      Height          =   615
      Left            =   80
      TabIndex        =   12
      Top             =   1410
      Width           =   5055
      Begin VB.OptionButton OptMoneda 
         Caption         =   "Soles"
         Height          =   225
         Index           =   0
         Left            =   510
         TabIndex        =   3
         Top             =   270
         Width           =   1425
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "Miles de Soles"
         Height          =   225
         Index           =   1
         Left            =   2730
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Moneda"
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
      Height          =   675
      Left            =   80
      TabIndex        =   10
      Top             =   680
      Width           =   5055
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmNIIFFormaB2.frx":030A
         Left            =   1080
         List            =   "frmNIIFFormaB2.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "Generar Reporte Nota Estado"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Frame fraMes 
      Caption         =   "Periodo"
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
      Height          =   675
      Left            =   80
      TabIndex        =   7
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   280
         Left            =   720
         MaxLength       =   4
         TabIndex        =   0
         Top             =   240
         Width           =   780
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmNIIFFormaB2.frx":030E
         Left            =   2760
         List            =   "frmNIIFFormaB2.frx":0336
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   270
         Width           =   390
      End
   End
   Begin ComctlLib.StatusBar EstadoBarra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2625
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNIIFFormaB2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmNIIFFormaB2
'** Descripción : Generación de Reporte Forma B2 segun ERS052-2013
'** Creación : EJVG, 20130502 09:00:00 AM
'********************************************************************
Option Explicit
Private Type TCtaCont
    CuentaContable As String
    Saldo As Currency
End Type

Private Sub Form_Load()
    CentraForm Me
    cargarMoneda
End Sub
Private Sub OptMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
            Case 0
                OptMoneda(1).SetFocus
            Case 1
                cmdGenerar.SetFocus
        End Select
    End If
End Sub
Private Sub txtAnio_Change()
    If Len(txtAnio.Text) = 4 Then
        cboMes.SetFocus
    End If
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       OptMoneda(0).SetFocus
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub
Private Sub chkMuestraResultAnioAnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdGenerar_Click()
    Dim ldFecha As Date
    Dim lnMoneda As Integer
        
    On Error GoTo ErrGenerar
    Screen.MousePointer = 11
    If validaGenerar = False Then Exit Sub
    
    ldFecha = obtenerFechaFinMes(cboMes.ListIndex + 1, txtAnio.Text)
    lnMoneda = CInt(Trim(Right(cboMoneda.Text, 2)))
    
    Call generarReporteFormaB2(ldFecha, lnMoneda, OptMoneda(1).value)

    Screen.MousePointer = 0
    Exit Sub
ErrGenerar:
    MsgBox "Ha ocurrido un error al procesar el Reporte, vuelvo a intentarlo," & Chr(10) & " si persiste comuniquese con el Dpto de TI", vbCritical, "Aviso"
    Screen.MousePointer = 0
End Sub
Private Function validaGenerar() As Boolean
    validaGenerar = True
    If Val(txtAnio.Text) <= 1900 Then
        MsgBox "Ud. debe especificar el año", vbInformation, "Aviso"
        txtAnio.SetFocus
        validaGenerar = False
        Exit Function
    End If
    If cboMes.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el mes", vbInformation, "Aviso"
        cboMes.SetFocus
        validaGenerar = False
        Exit Function
    End If
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar el tipo de moneda", vbInformation, "Aviso"
        cboMoneda.SetFocus
        validaGenerar = False
        Exit Function
    End If
End Function
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
Private Sub cargarMoneda()
    cboMoneda.AddItem "UNIFICADO" & Space(200) & "0"
    '''cboMoneda.AddItem "SOLES" & Space(200) & "1" 'MARG ERS044-2016
    cboMoneda.AddItem StrConv(gcPEN_PLURAL, vbUpperCase) & Space(200) & "1" 'MARG ERS044-2016
    cboMoneda.AddItem "DOLARES" & Space(200) & "2"
End Sub
Private Sub generarReporteFormaB2(ByVal pdFecha As Date, ByVal pnMoneda As Integer, ByVal pbMilesSoles As Boolean)
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim fs As New Scripting.FileSystemObject
    Dim oRep As New DRepFormula
    Dim rsB2 As New ADODB.Recordset
    Dim i As Long
    Dim lsArchivo As String
    Dim lnFilaActual As Integer
    Dim lsNombreMes As String
    Dim lnAnio As Integer

    lsArchivo = "NIIF_FormaB2_" & UCase(gsCodUser) & "_" & Format(gdFecSis, "yyyymmdd") & Format(Now, "hhmmss") & "_" & Format(pdFecha, "yyyymmdd") & ".xls"
 
    Set rsB2 = oRep.CargaRepFormula(, gContRepBaseFormaB2)
    If Not RSVacio(rsB2) Then
        Set xlsLibro = xlsAplicacion.Workbooks.Add
        Set xlsHoja = xlsLibro.ActiveSheet
        
        BarraProgreso.value = 0
        BarraProgreso.Min = 0
        BarraProgreso.Max = rsB2.RecordCount
        BarraProgreso.value = 0
        EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        
        lsNombreMes = dameNombreMes(Month(pdFecha), True)
        lnAnio = Year(pdFecha)
        
        xlsHoja.Range("A:A").ColumnWidth = 40
        xlsHoja.Range(xlsHoja.Cells(4, 2), xlsHoja.Cells(4, 4)).ColumnWidth = 40
        xlsHoja.Cells(2, 1) = "Forma ""B-2"""
        xlsHoja.Cells(4, 1) = "ESTADO DE RESULTADOS Y OTROS RESULTADOS INTEGRAL"
        xlsHoja.Cells(5, 1) = "CMAC Maynas S.A."
        xlsHoja.Cells(5, 1) = "CMAC Maynas S.A."
        xlsHoja.Cells(6, 1) = "Al " & Day(pdFecha) & " de " & lsNombreMes & " del " & Year(pdFecha)
        '''xlsHoja.Cells(7, 1) = "(Expresado en " & IIf(pbMilesSoles, "Miles de Nuevos Soles", "Nuevos Soles") 'MARG ERS044-2016
        xlsHoja.Cells(7, 1) = "(Expresado en " & IIf(pbMilesSoles, "Miles de " & StrConv(gcPEN_PLURAL, vbProperCase), StrConv(gcPEN_PLURAL, vbProperCase)) & ")" 'MARG ERS044-2016
        xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(4, 4)).MergeCells = True
        xlsHoja.Range(xlsHoja.Cells(5, 1), xlsHoja.Cells(5, 4)).MergeCells = True
        xlsHoja.Range(xlsHoja.Cells(6, 1), xlsHoja.Cells(6, 4)).MergeCells = True
        xlsHoja.Range(xlsHoja.Cells(7, 1), xlsHoja.Cells(7, 4)).MergeCells = True
        xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(6, 4)).Font.Bold = True
        xlsHoja.Range(xlsHoja.Cells(4, 1), xlsHoja.Cells(7, 4)).HorizontalAlignment = xlCenter
        
'        If pbMilesSoles Then
'            xlsHoja.Range("D:D").NumberFormat = "#,##0"
'        End If
        
        lnFilaActual = 9
        For i = 1 To rsB2.RecordCount
            xlsHoja.Cells(lnFilaActual, 1) = rsB2!cDescrip
            If rsB2!cFormula <> "" Then
                xlsHoja.Cells(lnFilaActual, 4) = "=" & ObtenerResultadoFormula(pdFecha, rsB2!cFormula, pnMoneda) & IIf(pbMilesSoles, "/1000", "")
            Else
                xlsHoja.Cells(lnFilaActual, 4) = Format("0.00", gsFormatoNumeroView)
            End If
            
            rsB2.MoveNext
            lnFilaActual = lnFilaActual + 1
            BarraProgreso.value = i
            EstadoBarra.Panels(1) = "Proceso: " & Format((BarraProgreso.value / BarraProgreso.Max) * 100, "#0.00") & "%"
        Next
        xlsHoja.SaveAs App.path & "\Spooler\" & lsArchivo
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        EstadoBarra.Panels(1) = "Proceso Terminado"
    Else
        MsgBox "No existe la configuración respectiva para generar el presente Reporte", vbInformation, "Aviso"
    End If
    
    Set rsB2 = Nothing
    Set oRep = Nothing
    Set xlsHoja = Nothing
    Set xlsLibro = Nothing
    Set xlsAplicacion = Nothing
End Sub
Private Function ObtenerResultadoFormula(ByVal pdFecha As Date, ByVal psFormula As String, ByVal pnMoneda As Integer) As Currency
    Dim oBal As New DbalanceCont
    Dim oFormula As New NInterpreteFormula
    Dim lsFormula As String, lsTmp As String, lsTmp1 As String, lsCadFormula As String
    Dim MatDatos() As TCtaCont
    Dim i As Long, j As Long, nCtaCont As Long
    
    lsFormula = Trim(psFormula)
    ReDim MatDatos(0)
    nCtaCont = 0

    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                nCtaCont = nCtaCont + 1
                ReDim Preserve MatDatos(nCtaCont)
                MatDatos(nCtaCont).CuentaContable = lsTmp
                MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
            End If
            lsTmp = ""
        End If
    Next i
    If Len(lsTmp) > 0 Then
        nCtaCont = nCtaCont + 1
        ReDim Preserve MatDatos(nCtaCont)
        MatDatos(nCtaCont).CuentaContable = lsTmp
        MatDatos(nCtaCont).Saldo = oBal.ObtenerCtaContBalanceMensual2(Mid(MatDatos(nCtaCont).CuentaContable, 1, 2) & IIf(Len(MatDatos(nCtaCont).CuentaContable) > 2, CStr(pnMoneda), "") & Mid(MatDatos(nCtaCont).CuentaContable, 4, Len(MatDatos(nCtaCont).CuentaContable)), pdFecha, CStr(pnMoneda), "1", 0, True)
    End If
    'Genero la formula en cadena
    lsTmp = ""
    lsCadFormula = ""
    For i = 1 To Len(lsFormula)
        If (Mid(Trim(lsFormula), i, 1) >= "0" And Mid(Trim(lsFormula), i, 1) <= "9") Then
            lsTmp = lsTmp + Mid(Trim(lsFormula), i, 1)
        Else
            If Len(lsTmp) > 0 Then
                For j = 1 To nCtaCont
                    If MatDatos(j).CuentaContable = lsTmp Then
                        lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
                        Exit For
                    End If
                Next j
            End If
            lsTmp = ""
            lsCadFormula = lsCadFormula & Mid(Trim(lsFormula), i, 1)
        End If
    Next
    If Len(lsTmp) > 0 Then
        For j = 1 To nCtaCont
           If MatDatos(j).CuentaContable = lsTmp Then
               lsCadFormula = lsCadFormula & Format(MatDatos(j).Saldo, "#0.00")
               Exit For
           End If
        Next j
    End If
    
    ObtenerResultadoFormula = oFormula.ExprANum(lsCadFormula)
    Set oBal = Nothing
    Set oFormula = Nothing
End Function
