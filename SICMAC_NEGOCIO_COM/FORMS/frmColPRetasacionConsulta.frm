VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmColPRetasacionConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Retasación"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frmColPRetasacionConsulta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbProgreso 
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   7780
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Frame fmEstados 
      Caption         =   "Estados"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox checkTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox checkAdjudicadas 
         Caption         =   "Adjudicadas"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox checkDiferidas 
         Caption         =   "Diferidas"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox checkVigente 
         Caption         =   "Vigente"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango"
      Height          =   1215
      Left            =   4560
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
      Begin VB.TextBox txtAnio 
         Height          =   300
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "txtAnio"
         Top             =   600
         Width           =   1035
      End
      Begin VB.ComboBox cboTrimestre 
         Height          =   315
         ItemData        =   "frmColPRetasacionConsulta.frx":030A
         Left            =   1200
         List            =   "frmColPRetasacionConsulta.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtNumRetas 
         Height          =   315
         Left            =   960
         MaxLength       =   12
         TabIndex        =   9
         Top             =   440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58785793
         CurrentDate     =   36161
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   960
         TabIndex        =   24
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58785793
         CurrentDate     =   36161
      End
      Begin VB.Label lblNum 
         Caption         =   "Nº :"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbl2 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   630
         Width           =   495
      End
      Begin VB.Label lbl1 
         Caption         =   "Del:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criterios de Selección"
      Height          =   1215
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton optNroRetasacion 
         Caption         =   "N° Retasación"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   850
         Width           =   1815
      End
      Begin VB.OptionButton optTrimestral 
         Caption         =   "Trimestral"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optFecha 
         Caption         =   "Rango fechas"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   580
         Width           =   1575
      End
   End
   Begin VB.Frame fmAgencias 
      Caption         =   "Agencias"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      Begin VB.CheckBox chktodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.ListBox lsAgencias 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   600
         Width           =   4035
      End
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   7680
      Width           =   1095
   End
   Begin SICMACT.FlexEdit FEDatos 
      Height          =   3615
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6376
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Agencia-Fecha-Retasación de-Muestra-Total Lote-Estado-nCodigoID-cPrepaRetasacion-nEstado"
      EncabezadosAnchos=   "400-2400-1200-1600-1200-1200-1200-0-0-0"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmColPRetasacionConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPRetasacionConsulta
'** Descripción : Formulario para realizar la consulta de la retazasion de creditos prendarios
'** Creación    : RECO, 20140707 - ERS074-2014
'**********************************************************************************************
Option Explicit
Dim oNColP As COMNColoCPig.NCOMColPContrato
Dim oDRetas As ADODB.Recordset
Dim oDRM As ADODB.Recordset
Dim rs As ADODB.Recordset

Dim xlsAplicacion As Excel.Application
Dim fs As Scripting.FileSystemObject
Dim lsArchivo As String, lsFile As String, lsNomHoja As String
Dim xlsLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmdAbrir_Click()
    Screen.MousePointer = 11
    cmdAbrir.Enabled = False
    If FEDatos.TextMatrix(FEDatos.row, 1) = "" Then
        MsgBox "No se pudo generar el archivo de retasación, no se encontraron datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If CInt(FEDatos.TextMatrix(FEDatos.row, 9)) = 1 Then
        Call GeneraArchivoPreparacion
    Else
        Call GeneraArchivoRetasacion
    End If

    Screen.MousePointer = 0
    cmdAbrir.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaFormulario
End Sub

Private Sub cmdConsultar_Click()
    Dim oNColP   As New COMNColoCPig.NCOMColPContrato

    Dim nTpoBusq As Integer, i As Integer, nTrimiestre As Integer, cAnio  As String, psFechaDesde As String, psFechaHasta As String, pnEstado As Integer

    Dim oDrDatos As New ADODB.Recordset
    LimpiaFlex FEDatos
    
    If RecuperaListaAgencias() = "" Then
        MsgBox "Debe seleccionar por lo menos una agencia", vbInformation, "Aviso"
        Exit Sub

    End If

    If optTrimestral.value = True Then
        
        If Len(Trim(Me.txtAnio)) = 0 Then
            MsgBox "Ingrese el año", vbCritical, "Alerta"
            txtAnio.SetFocus
            Exit Sub

        End If

        If cboTrimestre.ListIndex = -1 Then
            MsgBox "Es necesario seleccionar el trimestre", vbCritical, "Alerta"
            cboTrimestre.SetFocus
            Exit Sub

        End If

        If Len(Trim(Me.txtAnio.Text)) < 4 Or txtAnio.Text < 1900 Or txtAnio.Text > 2099 Then
            MsgBox "El año ingresado no es el correcto", vbCritical, "Alerta"
            txtAnio.SetFocus
            Exit Sub

        End If
        
        If checkVigente.value = 1 And checkDiferidas.value = 1 And checkAdjudicadas.value = 1 Then
            pnEstado = 123
        ElseIf checkVigente.value = Checked Then

            If checkDiferidas.value = Checked Then
                pnEstado = 12
            ElseIf checkAdjudicadas.value = Checked Then
                pnEstado = 13
            Else
                pnEstado = 1

            End If

        ElseIf checkDiferidas.value = Checked Then

            If checkAdjudicadas.value = Checked Then
                pnEstado = 23
            Else
                pnEstado = 2

            End If

        ElseIf checkAdjudicadas.value = Checked Then
            pnEstado = 3

        End If
    
        nTpoBusq = 1
        nTrimiestre = cboTrimestre.ItemData(cboTrimestre.ListIndex)
        cAnio = txtAnio.Text
        
        Set oDrDatos = oNColP.CargaListaRetasaciones(nTpoBusq, "", "", "", RecuperaListaAgencias, nTrimiestre, cAnio, pnEstado)
    
        '#-Agencia-Fecha-Retasación de-Muestra-Total Lote-Estado-nCodigoID-cPrepaRetasacion-nEstado
        If Not (oDrDatos.EOF And oDrDatos.BOF) Then

            For i = 1 To oDrDatos.RecordCount
                FEDatos.AdicionaFila
                FEDatos.TextMatrix(i, 1) = oDrDatos!cAgeDescripcion
                FEDatos.TextMatrix(i, 2) = Format(oDrDatos!dFechaRetas, "dd/MM/yyyy")
                FEDatos.TextMatrix(i, 3) = oDrDatos!estadoPrepaRetasacion
                FEDatos.TextMatrix(i, 4) = oDrDatos!NNUMCRED
                FEDatos.TextMatrix(i, 5) = oDrDatos!nTotLtes
                FEDatos.TextMatrix(i, 6) = oDrDatos!cEstado
                FEDatos.TextMatrix(i, 7) = oDrDatos!nCodigoID
                FEDatos.TextMatrix(i, 8) = oDrDatos!cCodPrepacion
                FEDatos.TextMatrix(i, 9) = oDrDatos!nEstado
                oDrDatos.MoveNext
            Next
        Else
            MsgBox "No se encontró datos de retasaciones", vbInformation, "Aviso"
        End If

    End If

    If optFecha.value = True Then
        
        If dtpDesde.value = "__/__/____" Then
            MsgBox "Es necesario proporcionar la fecha de inicio", vbCritical, "Alerta"
            dtpDesde.SetFocus
            Exit Sub

        End If

        If dtpHasta.value = "__/__/____" Then
            MsgBox "Es necesario proporcionar la fecha final", vbCritical, "Alerta"
            dtpHasta.SetFocus
            Exit Sub

        End If
        
        If checkVigente.value = 1 And checkDiferidas.value = 1 And checkAdjudicadas.value = 1 Then
            pnEstado = 123
        ElseIf checkVigente.value = Checked Then

            If checkDiferidas.value = Checked Then
                pnEstado = 12
            ElseIf checkAdjudicadas.value = Checked Then
                pnEstado = 13
            Else
                pnEstado = 1

            End If

        ElseIf checkDiferidas.value = Checked Then

            If checkAdjudicadas.value = Checked Then
                pnEstado = 23
            Else
                pnEstado = 2

            End If

        ElseIf checkAdjudicadas.value = Checked Then
            pnEstado = 3

        End If
        
        nTpoBusq = 2
        psFechaDesde = Format(dtpDesde.value, "yyyyMMdd")
        psFechaHasta = Format(dtpHasta.value, "yyyyMMdd")
        
        Set oDrDatos = oNColP.CargaListaRetasaciones(nTpoBusq, "", psFechaDesde, psFechaHasta, RecuperaListaAgencias, 0, "1999", pnEstado)

        If Not (oDrDatos.EOF And oDrDatos.BOF) Then

            For i = 1 To oDrDatos.RecordCount
                FEDatos.AdicionaFila
                FEDatos.TextMatrix(i, 1) = oDrDatos!cAgeDescripcion
                FEDatos.TextMatrix(i, 2) = Format(oDrDatos!dFechaRetas, "dd/MM/yyyy")
                FEDatos.TextMatrix(i, 3) = oDrDatos!estadoPrepaRetasacion
                FEDatos.TextMatrix(i, 4) = oDrDatos!NNUMCRED
                FEDatos.TextMatrix(i, 5) = oDrDatos!nTotLtes
                FEDatos.TextMatrix(i, 6) = oDrDatos!cEstado
                FEDatos.TextMatrix(i, 7) = oDrDatos!nCodigoID
                FEDatos.TextMatrix(i, 8) = oDrDatos!cCodPrepacion
                FEDatos.TextMatrix(i, 9) = oDrDatos!nEstado
                oDrDatos.MoveNext
            Next

        End If

    End If

    If optNroRetasacion.value = True Then
        If txtNumRetas.Text = "" Then
            MsgBox "Los valores no pueden estar vacíos", vbCritical, "Alerta"
            txtNumRetas.SetFocus
            Exit Sub

        End If

        nTpoBusq = 3
        
        Set oDrDatos = oNColP.CargaListaRetasaciones(nTpoBusq, txtNumRetas.Text, "", "", "", 0, "1999", 1)

        If Not (oDrDatos.EOF And oDrDatos.BOF) Then

            For i = 1 To oDrDatos.RecordCount
                FEDatos.AdicionaFila
                FEDatos.TextMatrix(i, 1) = oDrDatos!cAgeDescripcion
                FEDatos.TextMatrix(i, 2) = Format(oDrDatos!dFechaRetas, "dd/MM/yyyy")
                FEDatos.TextMatrix(i, 3) = oDrDatos!estadoPrepaRetasacion
                FEDatos.TextMatrix(i, 4) = oDrDatos!NNUMCRED
                FEDatos.TextMatrix(i, 5) = oDrDatos!nTotLtes
                FEDatos.TextMatrix(i, 6) = oDrDatos!cEstado
                FEDatos.TextMatrix(i, 7) = oDrDatos!nCodigoID
                FEDatos.TextMatrix(i, 8) = oDrDatos!cCodPrepacion
                FEDatos.TextMatrix(i, 9) = oDrDatos!nEstado
                oDrDatos.MoveNext
            Next

        End If

    End If
    
    If FEDatos.TextMatrix(1, 1) <> "" Then
        cmdAbrir.Enabled = True
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtpDesde_Change()
    If dtpDesde.value > dtpHasta.value Then
        dtpDesde.value = gdFecSis
        MsgBox "La Fecha de Inicio no debe ser mayor a la Fecha de Final", vbInformation, "Aviso"
    End If
End Sub

Private Sub dtpHasta_Change()
    If dtpDesde.value > dtpHasta.value Then
        dtpHasta.value = gdFecSis
        MsgBox "La Fecha de Inicio debe ser mayor a la Fecha Final", vbInformation, "Aviso"
    End If
End Sub

Private Sub Form_Load()
    Call LimpiaFormulario
    Call CargarListaAgencia
End Sub


Public Sub CargarListaAgencia()

    Dim loCargaAg As COMDColocPig.DCOMColPFunciones

    Dim lrAgenc   As ADODB.Recordset
    
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    Set lrAgenc = loCargaAg.dObtieneAgencias(True)
        
    Set loCargaAg = Nothing
    
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.lsAgencias.Clear

        With lrAgenc

            Do While Not .EOF
                lsAgencias.AddItem !cAgeCod & " " & Trim(!cAgeDescripcion)

                If !cAgeCod = gsCodAge Then
                    lsAgencias.Selected(lsAgencias.ListCount - 1) = True

                End If

                .MoveNext
            Loop

        End With

    End If

End Sub

Public Sub LimpiaFormulario()
    'FEDatos
    LimpiaFlex FEDatos
    fmEstados.Enabled = True
    fmAgencias.Enabled = True
    Call LimpiarListaAge
    Call ActivarControles(True)
    txtNumRetas.Text = ""
    txtAnio.Text = ""
    dtpDesde.value = Format(gdFecSis, "dd/MM/yyyy")
    dtpHasta.value = Format(gdFecSis, "dd/MM/yyyy")
    optTrimestral.value = True
    cboTrimestre.ListIndex = 4
    checkTodos.value = Unchecked
    checkVigente.value = Checked
    checkDiferidas.value = Unchecked
    checkAdjudicadas.value = Unchecked
    cmdAbrir.Enabled = False

    'txtAnio.SetFocus
End Sub

Private Sub chkTodos_Click()
    If chktodos.value = 0 Then
        Call LimpiarListaAge
    Else
        Call SelecListaAgeTodos
    End If
End Sub

Public Sub SelecListaAgeTodos()
    Dim nIndex As Integer
    For nIndex = 0 To lsAgencias.ListCount - 1
        lsAgencias.Selected(nIndex) = True
    Next
End Sub
Public Sub LimpiarListaAge()
    Dim nIndex As Integer
    For nIndex = 0 To lsAgencias.ListCount - 1
        lsAgencias.Selected(nIndex) = False
    Next
End Sub

Public Function RecuperaListaAgencias() As String
    Dim nIndex As Integer
    Dim lsCadAge  As String
    RecuperaListaAgencias = 0
    For nIndex = 0 To Me.lsAgencias.ListCount - 1
        If Me.lsAgencias.Selected(nIndex) Then
            lsCadAge = lsCadAge & Left(Me.lsAgencias.List(nIndex), 2) & ","
            RecuperaListaAgencias = RecuperaListaAgencias + 1
        End If
    Next
    If lsCadAge = "" Then
        Exit Function
    End If
    lsCadAge = Mid(lsCadAge, 1, Len(lsCadAge) - 1)
    RecuperaListaAgencias = lsCadAge
End Function

Public Sub GeneraArchivoPreparacion()
'On Error GoTo ErrorExcel
    Set xlsAplicacion = New Excel.Application
    Set fs = New Scripting.FileSystemObject
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim lsArchivo As String, lsFile As String, lsNomHoja As String
    Dim lbExisteHoja As Boolean
    Dim lnValorConteo As Integer
    Dim i As Integer: Dim J As Integer
    
    'Set fs = New Scripting.FileSystemObject
    'Set xlsAplicacion = New Excel.Application
    
    Set oNColP = New COMNColoCPig.NCOMColPContrato
    
    Set rs = New ADODB.Recordset
    Set oDRetas = New ADODB.Recordset
    
    Dim lsCtaCod As String
    Dim HoraSis As Variant
    Dim HoraCrea As String
    
    Dim lnOrden As Integer
    Dim lnPosicion As Integer
    Dim lnFilaTmp As Integer
    Dim lnTotPiezas As Integer
    Dim lnPesoBruto As Currency
    Dim lnPesoNeto As Currency
    
    
    HoraSis = Time
    HoraCrea = CStr(Hour(HoraSis)) & Minute(HoraSis) & Second(HoraSis)
    
    lsNomHoja = "Retasacion"
    lsFile = "FormatoRetasacionPreparacion"
    
    lsArchivo = "\spooler\" & "Preparación_Retasación" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & HoraCrea & ".xls"
    
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        xlHoja1.Name = lsNomHoja
    End If
    
    Set rs = oNColP.ObtieneCtaCodRetasacionPrep(FEDatos.TextMatrix(FEDatos.row, 7))
    Set oDRetas = oNColP.DevuelveDatosListaCredPrepRetasacion(rs!cCtaCad, 1, CInt(FEDatos.TextMatrix(FEDatos.row, 7)))
    'Set oDR = oNColP.DevuelveDatosListaCredPrepRetasacion(rs!cCtaCad, CInt(FEDatos.TextMatrix(FEDatos.row, 7)), CInt(FEDatos.TextMatrix(FEDatos.row, 5)))
   
    If (oDRetas.BOF And oDRetas.EOF) Then
        MsgBox "No se encontro datos de la preparación de la retasación. Por favor comuníquese con TI.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    pbProgreso.Min = 0
    pbProgreso.Max = oDRetas.RecordCount
    pbProgreso.value = 0
    pbProgreso.Visible = True
    
    xlsAplicacion.DisplayAlerts = False
    
    xlHoja1.Range("A1:B1").Merge True
    xlHoja1.Range("A1:B1").HorizontalAlignment = xlLeft
    
    xlHoja1.Range("A2:B2").Font.Bold = False
    xlHoja1.Range("A2:B2").Merge True
    xlHoja1.Range("A2:B2").HorizontalAlignment = xlLeft

    xlHoja1.Range("A3:B3").Merge True
    xlHoja1.Range("A3:B3").HorizontalAlignment = xlLeft
    
    xlHoja1.Range("C2:S2").Font.Bold = True
    xlHoja1.Range("C2:S2").Merge True
    xlHoja1.Range("C2:S2").WrapText = True
    xlHoja1.Range("C2:S2").HorizontalAlignment = xlCenter
    
    xlHoja1.Cells(1, 1) = "Total de Lotes" & Space(16) & ":" & Space(5) & oDRetas!nTotLotes
    xlHoja1.Cells(2, 1) = "Total de Muestra" & Space(12) & ":" & Space(5) & oDRetas!nMuestra
    
    xlHoja1.Cells(2, 3) = "LISTADO DE ORO RETASADO DE LA " & UCase(FEDatos.TextMatrix(FEDatos.row, 1))
    xlHoja1.Cells(3, 1) = "Fecha de Preparción" & Space(5) & ":" & Space(5) & Format(gdFecSis, "dd/MM/yyyy")
    xlHoja1.Cells(4, 2) = oDRetas!cCodPrepacion & "-" & oDRetas!nCodigoID
    
    i = 5
    
    Do While Not oDRetas.EOF
         i = i + 1
        pbProgreso.value = pbProgreso.value + 1
        If lsCtaCod <> oDRetas!cPigCod Then
            lnTotPiezas = lnTotPiezas + oDRetas!nTotPiezas
            lnOrden = lnOrden + 1
            xlHoja1.Cells(i, 1) = lnOrden
            lnPosicion = i + 1
            lnFilaTmp = 0
        Else
            lnFilaTmp = lnFilaTmp + 1
        End If

        lnPesoBruto = lnPesoBruto + oDRetas!nPesoBruto
        lnPesoNeto = lnPesoNeto + oDRetas!nPesoNeto

        'xlHoja1.Cells(i, 1) = lnFilaTmp
        xlHoja1.Cells(i, 2) = oDRetas!cPigCod
        xlHoja1.Cells(i, 3) = oDRetas!cPersNombre
        xlHoja1.Cells(i, 4) = oDRetas!nItem
        xlHoja1.Cells(i, 5) = oDRetas!nTotPiezas
        xlHoja1.Cells(i, 6) = oDRetas!nPiezas
        xlHoja1.Cells(i, 7) = oDRetas!cDescrip
        xlHoja1.Cells(i, 8) = oDRetas!cUserTas
        xlHoja1.Cells(i, 9) = oDRetas!cKilataje
        xlHoja1.Cells(i, 10) = Format(oDRetas!nTotLote, gcFormView)
        xlHoja1.Cells(i, 11) = Format(oDRetas!nPesoBruto, gcFormView)
        xlHoja1.Cells(i, 12) = Format(oDRetas!nPesoNeto, gcFormView)
        xlHoja1.Cells(i, 13) = IIf(oDRetas!nHolograma = 0, "Sin Holograma", oDRetas!nHolograma)
        
        xlHoja1.Range("A" & Trim(str(i)) & ":" & "T" & Trim(str(i))).Borders.LineStyle = 1
        xlHoja1.Range("A" & Trim(str(i)) & ":" & "S" & Trim(str(i))).WrapText = True
        xlHoja1.Range("J" & Trim(str(i)) & ":" & "M" & Trim(str(i))).Interior.Color = RGB(204, 255, 255)
        xlHoja1.Range("N" & Trim(str(i)) & ":" & "S" & Trim(str(i))).Interior.Color = RGB(255, 255, 153)
        
        If lsCtaCod = oDRetas!cPigCod Then
            xlHoja1.Range("A" & Trim(str(lnPosicion - 1)) & ":" & "A" & Trim(str(i))).MergeCells = True
            'xlHoja1.Range("C" & Trim(str(lnPosicion - 1)) & ":" & "C" & Trim(str(I))).MergeCells = True
            'xlHoja1.Range("E" & Trim(str(lnPosicion - 1)) & ":" & "E" & Trim(str(I))).MergeCells = True
            'xlHoja1.Range("H" & Trim(str(lnPosicion - 1)) & ":" & "H" & Trim(str(I))).MergeCells = True
            'xlHoja1.Range("J" & Trim(str(lnPosicion - 1)) & ":" & "J" & Trim(str(I))).MergeCells = True
            'xlHoja1.Range("N" & Trim(str(lnPosicion - 1)) & ":" & "N" & Trim(str(I))).MergeCells = True
        End If
        
        lsCtaCod = oDRetas!cPigCod
        oDRetas.MoveNext
        If oDRetas.EOF Then
            Exit Do
        End If
    Loop
    
    lnValorConteo = i
    
    '[TORE RFC1811260001: ADD - Total de Piezas, Peso Bruto, Peso Neto]
    xlHoja1.Cells(i + 1, 5) = lnTotPiezas
    xlHoja1.Cells(i + 1, 11) = Format(lnPesoBruto, gcFormView)
    xlHoja1.Cells(i + 1, 12) = Format(lnPesoNeto, gcFormView)
    xlHoja1.Cells(i + 1, 5) = lnTotPiezas
    
    lsNomHoja = "Resumen"
    'Cargamos los datos de los miembros del comite de retasacion
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        xlHoja1.Name = lsNomHoja
    End If
    
    xlHoja1.Range("C3").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """)"
    xlHoja1.Range("C4").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "10" & """)"
    xlHoja1.Range("C5").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """)"
    xlHoja1.Range("C6").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """)"
    xlHoja1.Range("C7").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """)"
    xlHoja1.Range("C8").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """)"
    xlHoja1.Range("C9").FormulaLocal = "=CONTAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """)"
    
    xlHoja1.Range("D3").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D4").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "10" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D5").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D6").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D7").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D8").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D9").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """,Retasacion!O6:O" & CStr(lnValorConteo) & ")"
    
    xlHoja1.Range("E3").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E4").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "10" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E5").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E6").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E7").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E8").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E9").FormulaLocal = "=SUMAR.SI(Retasacion!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """,Retasacion!P6:P" & CStr(lnValorConteo) & ")"
    

    lsNomHoja = "Retasacion"
    'Cargamos los datos de los miembros del comite de retasacion
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    
    pbProgreso.Visible = False
    '[END TORE RFC1811260001: ADD - Total de Piezas, Peso Bruto, Peso Neto]
    
    Set oNColP = Nothing
    Set rs = Nothing
    Set oDRetas = Nothing
    
    xlsAplicacion.DisplayAlerts = False
    xlHoja1.SaveAs App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
'ErrorExcel:
'    MsgBox "El sistema está intentando guardar el archivo " & Chr(13) & Right(lsArchivo, 51) & Chr(13) & "pero este se encuentra abierto. Por favor cierre el archivo excel y vuelva a intentar abrir el documento", vbApplicationModal + vbInformation, "Aviso"
    'MsgBox Err.Description, vbApplicationModal + vbInformation, "Aviso"
End Sub


Public Sub GeneraArchivoRetasacion()
'On Error GoTo ErrorExcel
    Set xlsAplicacion = New Excel.Application
    Set fs = New Scripting.FileSystemObject
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim lsArchivo As String, lsFile As String, lsNomHoja As String
    Dim lbExisteHoja As Boolean
    Dim lnValorConteo As Integer
    Dim i As Integer: Dim J As Integer
    Dim iM As Integer
    
    'Set fs = New Scripting.FileSystemObject
    'Set xlsAplicacion = New Excel.Application
    
    Set oNColP = New COMNColoCPig.NCOMColPContrato
    
    Set rs = New ADODB.Recordset
    'Set rsM = New ADODB.Recordset
    Set oDRetas = New ADODB.Recordset
    
    Dim lsCtaCod As String
    Dim HoraSis As Variant
    Dim HoraCrea As String
    
    Dim lnOrden As Integer
    Dim lnPosicion As Integer
    Dim lnFilaTmp As Integer
    Dim lnTotPiezas As Integer
    Dim lnPesoBruto As Currency
    Dim lnPesoNeto As Currency
    
    HoraSis = Time
    HoraCrea = CStr(Hour(HoraSis)) & Minute(HoraSis) & Second(HoraSis)
    
    lsNomHoja = "RetasacionFinal"
    lsFile = "FormatoRetasacionFinal"
    
    lsArchivo = "\spooler\" & "Retasación" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & HoraCrea & ".xls"
    
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        xlHoja1.Name = lsNomHoja
    End If
    
    Set rs = oNColP.ObtieneCtaCodRetasacionPrep(FEDatos.TextMatrix(FEDatos.row, 7))
    'Set oDR = oNColP.DevuelveDatosListaCredPrepRetasacion(rs!cCtaCad, CInt(FEDatos.TextMatrix(FEDatos.row, 7)), CInt(FEDatos.TextMatrix(FEDatos.row, 5)))
    Set oDRetas = oNColP.DevuelveDatosListaCredPrepRetasacion(rs!cCtaCad, 2, CInt(FEDatos.TextMatrix(FEDatos.row, 7)))
    If (oDRetas.BOF And oDRetas.EOF) Then
        MsgBox "No se encontro datos de la preparación de la retasación. Por favor comuniquese con TI.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    pbProgreso.Min = 0
    pbProgreso.Max = oDRetas.RecordCount
    pbProgreso.value = 0
    pbProgreso.Visible = True
    
    xlsAplicacion.DisplayAlerts = False
    
    xlHoja1.Range("A1:B1").Merge True
    xlHoja1.Range("A1:B1").HorizontalAlignment = xlLeft
    
    xlHoja1.Range("A2:B2").Font.Bold = False
    xlHoja1.Range("A2:B2").Merge True
    xlHoja1.Range("A2:B2").HorizontalAlignment = xlLeft

    xlHoja1.Range("A3:B3").Merge True
    xlHoja1.Range("A3:B3").HorizontalAlignment = xlLeft
    
    xlHoja1.Range("C2:S2").Font.Bold = True
    xlHoja1.Range("C2:S2").Merge True
    xlHoja1.Range("C2:S2").WrapText = True
    xlHoja1.Range("C2:S2").HorizontalAlignment = xlCenter
    
    xlHoja1.Cells(1, 1) = "Total de Lotes" & Space(16) & ":" & Space(5) & oDRetas!nTotLotes
    xlHoja1.Cells(2, 1) = "Total de Muestra" & Space(12) & ":" & Space(5) & oDRetas!nMuestra
    
    xlHoja1.Cells(2, 3) = "LISTADO DE ORO RETASADO DE LA " & UCase(FEDatos.TextMatrix(FEDatos.row, 1))
    xlHoja1.Cells(3, 1) = "Fecha de Preparción" & Space(5) & ":" & Space(5) & Format(gdFecSis, "dd/MM/yyyy")
    xlHoja1.Cells(4, 2) = oDRetas!cCodPrepacion & "-" & oDRetas!nCodigoID
    
    i = 5
    lnValorConteo = 0
    Do While Not oDRetas.EOF
        i = i + 1
        pbProgreso.value = pbProgreso.value + 1
         If lsCtaCod <> oDRetas!cPigCod Then
            lnTotPiezas = lnTotPiezas + oDRetas!PCANT
            lnOrden = lnOrden + 1
            xlHoja1.Cells(i, 1) = lnOrden
            lnPosicion = i + 1
            lnFilaTmp = 0
        Else
            lnFilaTmp = lnFilaTmp + 1
        End If

        lnPesoBruto = lnPesoBruto + oDRetas!nPesoBruto
        lnPesoNeto = lnPesoNeto + oDRetas!nPesoNeto

        xlHoja1.Cells(i, 2) = oDRetas!cPigCod
        xlHoja1.Cells(i, 3) = oDRetas!cPersNombre
        xlHoja1.Cells(i, 4) = oDRetas!nItem
        xlHoja1.Cells(i, 5) = oDRetas!PCANT
        xlHoja1.Cells(i, 6) = oDRetas!nPiezas
        xlHoja1.Cells(i, 7) = oDRetas!cDescrip
        xlHoja1.Cells(i, 8) = oDRetas!cUserTasador
        xlHoja1.Cells(i, 9) = oDRetas!cKilataje
        xlHoja1.Cells(i, 10) = Format(oDRetas!PnTot, gcFormView)
        xlHoja1.Cells(i, 11) = Format(oDRetas!nPesoBruto, gcFormView)
        xlHoja1.Cells(i, 12) = Format(oDRetas!nPesoNeto, gcFormView)
        xlHoja1.Cells(i, 13) = IIf(oDRetas!nHolograma = 0, "Sin Holograma", CStr(oDRetas!nHolograma))

        xlHoja1.Cells(i, 14) = oDRetas!RKilataje
        xlHoja1.Cells(i, 15) = Format(oDRetas!RPBruto, gcFormView)
        xlHoja1.Cells(i, 16) = Format(oDRetas!RPNeto, gcFormView)
        xlHoja1.Cells(i, 17) = oDRetas!RnTot
        xlHoja1.Cells(i, 18) = oDRetas!RNroHolograma
        xlHoja1.Cells(i, 19) = oDRetas!dFechRetasacion
        xlHoja1.Cells(i, 20) = oDRetas!RObservacion

        xlHoja1.Range("A" & Trim(str(i)) & ":" & "T" & Trim(str(i))).Borders.LineStyle = 1
        xlHoja1.Range("J" & Trim(str(i)) & ":" & "M" & Trim(str(i))).Interior.Color = RGB(204, 255, 255)
        xlHoja1.Range("N" & Trim(str(i)) & ":" & "S" & Trim(str(i))).Interior.Color = RGB(255, 255, 153)
        
          If lsCtaCod = oDRetas!cPigCod Then
            xlHoja1.Range("A" & Trim(str(lnPosicion - 1)) & ":" & "A" & Trim(str(i))).MergeCells = True
            xlHoja1.Range("B" & Trim(str(lnPosicion - 1)) & ":" & "B" & Trim(str(i))).MergeCells = True
            xlHoja1.Range("C" & Trim(str(lnPosicion - 1)) & ":" & "C" & Trim(str(i))).MergeCells = True
            xlHoja1.Range("E" & Trim(str(lnPosicion - 1)) & ":" & "E" & Trim(str(i))).MergeCells = True
            xlHoja1.Range("H" & Trim(str(lnPosicion - 1)) & ":" & "H" & Trim(str(i))).MergeCells = True
            xlHoja1.Range("J" & Trim(str(lnPosicion - 1)) & ":" & "J" & Trim(str(i))).MergeCells = True
            xlHoja1.Range("Q" & Trim(str(lnPosicion - 1)) & ":" & "Q" & Trim(str(i))).MergeCells = True
        End If
        
        lsCtaCod = oDRetas!cPigCod
          If oDRetas.EOF Then
            Exit Do
        End If
        
        oDRetas.MoveNext
    Loop
    
    lnValorConteo = i
    '[TORE RFC1811260001: ADD - Total de Piezas, Peso Bruto, Peso Neto]
    xlHoja1.Cells(i + 1, 5) = lnTotPiezas
    xlHoja1.Cells(i + 1, 11) = Format(lnPesoBruto, gcFormView)
    xlHoja1.Cells(i + 1, 12) = Format(lnPesoNeto, gcFormView)
    
'    xlHoja1.Range("W8").FormulaLocal = "=CONTAR.SI(N6:N" & CStr(i) & ",""" & "F" & """)"
'    xlHoja1.Range("W9").FormulaLocal = "=CONTAR.SI(N6:N" & CStr(i) & ",""" & "12" & """)"
'    xlHoja1.Range("W10").FormulaLocal = "=CONTAR.SI(N6:N" & CStr(i) & ",""" & "14" & """)"
'    xlHoja1.Range("W11").FormulaLocal = "=CONTAR.SI(N6:N" & CStr(i) & ",""" & "16" & """)"
'    xlHoja1.Range("W12").FormulaLocal = "=CONTAR.SI(N6:N" & CStr(i) & ",""" & "18" & """)"
'    xlHoja1.Range("W13").FormulaLocal = "=CONTAR.SI(N6:N" & CStr(i) & ",""" & "21" & """)"

    '[END TORE RFC1811260001: ADD - Total de Piezas, Peso Bruto, Peso Neto]
    
    lsNomHoja = "Comite"
    'Cargamos los datos de los miembros del comite de retasacion
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        xlHoja1.Name = lsNomHoja
    End If
    
    Set oDRetas = oNColP.ObtieneMiembrosRetasacion(FEDatos.TextMatrix(FEDatos.row, 7))
    If Not (oDRetas.BOF And oDRetas.EOF) Then
        pbProgreso.Min = 0
        pbProgreso.Max = oDRetas.RecordCount
        pbProgreso.value = 0
        
        i = 2
        Do While Not oDRetas.EOF
            i = i + 1
            pbProgreso.value = pbProgreso.value + 1
            xlHoja1.Cells(i, 1) = oDRetas!nOrden
            xlHoja1.Cells(i, 2) = oDRetas!cMiembro
            xlHoja1.Cells(i, 3) = oDRetas!cRolMiembro
            xlHoja1.Range("A" & Trim(str(i)) & ":" & "C" & Trim(str(i))).Borders.LineStyle = 1
        
            If oDRetas.EOF Then
                Exit Do
            End If
            oDRetas.MoveNext
        Loop
    Else
         MsgBox "No se encontro datos de los miembros de la retasación.", vbInformation, "Aviso"
    End If
    
    lsNomHoja = "Resumen"
    'Cargamos los datos de los miembros del comite de retasacion
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        xlHoja1.Name = lsNomHoja
    End If
    
'    xlHoja1.Range("C5").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """)"
'    xlHoja1.Range("C6").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """)"
'    xlHoja1.Range("C7").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """)"
'    xlHoja1.Range("C8").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """)"
'    xlHoja1.Range("C9").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """)"
'    xlHoja1.Range("C10").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """)"

    xlHoja1.Range("C3").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """)"
    xlHoja1.Range("C4").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "10" & """)"
    xlHoja1.Range("C5").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """)"
    xlHoja1.Range("C6").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """)"
    xlHoja1.Range("C7").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """)"
    xlHoja1.Range("C8").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """)"
    xlHoja1.Range("C9").FormulaLocal = "=CONTAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """)"
    
    xlHoja1.Range("D3").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D4").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "10" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D5").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D6").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D7").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D8").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("D9").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """,RetasacionFinal!O6:O" & CStr(lnValorConteo) & ")"
    
    xlHoja1.Range("E3").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "F" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E4").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "10" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E5").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "12" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E6").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "14" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E7").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "16" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E8").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "18" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    xlHoja1.Range("E9").FormulaLocal = "=SUMAR.SI(RetasacionFinal!N6:N" & CStr(lnValorConteo) & ",""" & "21" & """,RetasacionFinal!P6:P" & CStr(lnValorConteo) & ")"
    
    
    
    lsNomHoja = "RetasacionFinal"
    'Cargamos los datos de los miembros del comite de retasacion
    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
       End If
    Next
    
    
    pbProgreso.Visible = False
    
    Set oNColP = Nothing
    Set rs = Nothing
    Set oDRetas = Nothing
    
    xlsAplicacion.DisplayAlerts = False
    xlHoja1.SaveAs App.Path & lsArchivo
    xlsAplicacion.Visible = True
    
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
'ErrorExcel:
'    MsgBox "El sistema está intentando guardar el archivo " & Chr(13) & Right(lsArchivo, 51) & Chr(13) & "pero este se encuentra abierto. Por favor cierre el archivo excel y vuelva a intentar abrir el documento", vbApplicationModal + vbInformation, "Aviso"
'    MsgBox Err.Description, vbApplicationModal + vbInformation, "Aviso"
End Sub






'TORE ERS054-2017
Private Sub ExcelEnd(ByRef xlAplicacion As Excel.Application, ByRef xlLibro As Excel.Workbook, ByRef xlHoja As Excel.Worksheet)
    xlLibro.Close
    Sleep (800)
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja = Nothing
End Sub



'TORE ERS054-2017
Private Sub ActivarControles(ByVal Estado As Boolean)
    Dim rangot As Boolean
    Dim RangoF As Boolean
    Dim RangoNroRetacacion As Boolean
    rangot = optTrimestral.value
    RangoF = optFecha.value
    RangoNroRetacacion = optNroRetasacion.value
    
    If rangot = Estado Then
        lbl1.Visible = Estado
        lbl2.Visible = Estado
        lbl1.Caption = "Trimestre:"
        lbl2.Caption = "Año :"
        cboTrimestre.Visible = Estado
        txtAnio.Visible = Estado
        dtpDesde.Visible = Not Estado
        dtpHasta.Visible = Not Estado
        
        lblNum.Visible = Not Estado
        txtNumRetas.Visible = Not Estado
        
    End If
    If RangoF = Estado Then
        lbl1.Visible = Estado
        lbl2.Visible = Estado
        lbl1.Caption = "Desde :"
        lbl2.Caption = "Hasta :"
        dtpDesde.Visible = Estado
        dtpDesde.value = Format(gdFecSis, "dd/MM/yyyy")
        dtpHasta.Visible = Estado
        dtpHasta.value = Format(gdFecSis, "dd/MM/yyyy")
        cboTrimestre.Visible = Not Estado
        txtAnio.Visible = Not Estado
        lblNum.Visible = Not Estado
        txtNumRetas.Visible = Not Estado
    End If
    If RangoNroRetacacion = Estado Then
        dtpDesde.Visible = Not Estado
        dtpDesde.value = Format(gdFecSis, "dd/MM/yyyy")
        dtpHasta.Visible = Not Estado
        dtpHasta.value = Format(gdFecSis, "dd/MM/yyyy")
        lbl1.Visible = Not Estado
        lbl2.Visible = Not Estado
        cboTrimestre.Visible = Not Estado
        txtAnio.Visible = Not Estado
        txtAnio.Text = ""
        
        lblNum.Visible = Estado
        txtNumRetas.Visible = Estado
    End If
    
End Sub

Private Sub optNroRetasacion_Click()
    checkVigente.value = 0
    checkDiferidas.value = 0
    checkAdjudicadas.value = 0
    Call LimpiarListaAge
    fmEstados.Enabled = False
    fmAgencias.Enabled = False
    Call ActivarControles(True)
End Sub
Private Sub optTrimestral_Click()
    'Call LimpiarListaAge
    fmEstados.Enabled = True
    fmAgencias.Enabled = True

    Call ActivarControles(True)
End Sub

Private Sub optFecha_Click()
    'Call LimpiarListaAge
    fmEstados.Enabled = True
    fmAgencias.Enabled = True
    
    Call ActivarControles(True)
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = SoloNumerosTxt(KeyAscii)
End Sub

'[TORE RFC1811260001: Comentado tras modificacion en el codigo de retasacion]
'Private Sub txtNumRetas_KeyPress(KeyAscii As Integer)
'    KeyAscii = SoloNumerosTxt(KeyAscii)
'End Sub
'[END TORE RFC1811260001: Comentado tras modificacion en el codigo de retasacion]

Private Function SoloNumerosTxt(ByVal KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumerosTxt = 0
    Else
        SoloNumerosTxt = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumerosTxt = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumerosTxt = KeyAscii 'Enter
End Function




'Comentado por TORE según modificaciones del ERS054-2017
'Private Sub checkTodos_Click()
'    If checkTodos.value = Checked Then
'        checkVigente.value = Checked
'        checkDiferidas.value = Checked
'        checkAdjudicadas.value = Checked
'    Else
'        checkVigente.value = Unchecked
'        checkDiferidas.value = Unchecked
'        checkAdjudicadas.value = Unchecked
'    End If
'End Sub

'Private Sub checkAdjudicadas_Click()
'    If checkVigente.value = Checked And checkDiferidas.value = Checked Then
'        checkTodos.value = Checked
'    Else
'
'    End If
'End Sub

'Private Sub checkDiferidas_Click()
'    If checkVigente.value = Checked And checkAdjudicadas.value = Checked Then
'        checkTodos.value = Checked
'    Else
'
'    End If
'End Sub

'Private Sub checkVigente_Click()
'    If checkDiferidas.value = Checked And checkAdjudicadas.value = Checked Then
'        checkTodos.value = Checked
'    Else
'
'    End If
'End Sub


'END TORE

