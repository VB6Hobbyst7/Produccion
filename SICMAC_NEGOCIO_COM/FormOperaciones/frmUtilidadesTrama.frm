VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmUtilidadesTrama 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trama Utilidades Ex Trabajadores"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10575
   Icon            =   "frmUtilidadesTrama.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Cancelar"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   4815
      Width           =   915
   End
   Begin ComctlLib.ProgressBar pbProgres 
      Height          =   195
      Left            =   1170
      TabIndex        =   9
      Top             =   4880
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   1260
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   315
      Left            =   9585
      TabIndex        =   8
      Top             =   4815
      Width           =   915
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   315
      Left            =   8460
      TabIndex        =   7
      Top             =   4815
      Width           =   1050
   End
   Begin VB.Frame fraPeriodo 
      Caption         =   "Periodo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4695
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   10410
      Begin VB.CommandButton cmdFormato 
         Caption         =   "Formato"
         Height          =   315
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   870
      End
      Begin VB.ComboBox cbPeriodoTrama 
         Height          =   315
         Left            =   9675
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   600
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   3885
         Left            =   90
         TabIndex        =   6
         Top             =   720
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   6853
         Cols0           =   26
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmUtilidadesTrama.frx":030A
         EncabezadosAnchos=   "500-3300-1100-1000-1200-1200-1200-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C-R-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-C-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-2-2"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   315
         Left            =   6765
         TabIndex        =   4
         Top             =   315
         Width           =   420
      End
      Begin VB.TextBox txtArchivo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   3750
      End
      Begin VB.ComboBox cbPeriodo 
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   7245
         TabIndex        =   5
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo:"
         Height          =   240
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmUtilidadesTrama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'* NOMBRE          : "frmUtilidadesTrama"
'* DESCRIPCION   : Formulario creado para la carga de la trama de utilidades de ex trabajadores
'* CREACION        : RIRO, 20150513 ERS162-2014
'************************************************************************************

Option Explicit

Private sNomApell As String
Private sNroDoc As String
Private sPeriodo As String
Private sMoneda As String
Private sImporte As String

Private oUtil As PagoUtilidades

Private sCargo As String
Private sArea As String
Private sfechaIngreso As String

Private sC09_ParticipAdistribuir As String
Private sC10_DiasLaborTodosTrabAnio As String
Private sC11_RemunPercibTodosTrabAnio As String
Private sC12_MontDistribXdiasLabor As String
Private sC13_MontDistribXremunPercib As String
Private sC14_TotDiasEfectivLabor As String
Private sC15_TotRemuneraciones As String
Private sC16_ParticipXdiasLabor As String
Private sC17_ParticipXremuneraciones As String
Private sC18_TotParticipUtilidades As String
Private sC19_RetencionImpuestoRenta As String
Private sC20_TotalDescuento As String
Private sC21_TotalPagar As String

Private sRetencionJudicial As String
Private sOtrosDescuentos As String

Private bCerrar As Boolean
Private nIdTrama As Long
Private nEstadoCarga As Integer ' 1: Trama Cargada, 0: Trama NO cargada
Private nTipoOperacion As Integer ' 1: Alta de Trama, 2: Baja de Trama

Private Sub cmdGuardar_Click()
    
    On Error GoTo error_handler
        
    Dim oPer As COMDPersona.DCOMPersonas
    Dim I As Integer
    Dim bResultado As Boolean
    Dim sMensaje As String

    sMensaje = ""
    sNomApell = ""
    sNroDoc = ""
    sPeriodo = ""
    sMoneda = ""
    sImporte = ""
    sMensaje = Trim(validacion) 'Proceso de validacion

    If Len(sMensaje) > 0 Then
        MsgBox "Se presentaron las siguientes observaciones: " & vbNewLine & sMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    If nTipoOperacion = 1 Then 'Alta de Trama
        sMensaje = "¿Está seguro de grabar la información?"
    Else ' Baja de Trama
        sMensaje = "¿Está seguro de DAR DE BAJA a la trama seleccionada?"
    End If
    If MsgBox(sMensaje, vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oPer = New COMDPersona.DCOMPersonas
    If nTipoOperacion = 2 Then 'Baja de Trama
                
        bResultado = oPer.BajaTramaUtilidades(CInt(Val(Trim(Right(cbPeriodoTrama.Text, 10)))))
        If bResultado Then
            LimpiarUtilidades
            MsgBox "Se concluyó el proceso de anulacion de trama de utildiades Satisfactoriamente", vbInformation, "Aviso"
        Else
            MsgBox "Se presentó un error durante el proceso anulación de trama", vbExclamation, "Aviso"
        End If
    
    ElseIf nTipoOperacion = 1 Then 'Alta de Trama
    
        sNomApell = ""
        sNroDoc = ""
        sPeriodo = ""
        sMoneda = ""
        sImporte = ""
        sCargo = ""
        sArea = ""
        sfechaIngreso = ""
        sC09_ParticipAdistribuir = ""
        sC10_DiasLaborTodosTrabAnio = ""
        sC11_RemunPercibTodosTrabAnio = ""
        sC12_MontDistribXdiasLabor = ""
        sC13_MontDistribXremunPercib = ""
        sC14_TotDiasEfectivLabor = ""
        sC15_TotRemuneraciones = ""
        sC16_ParticipXdiasLabor = ""
        sC17_ParticipXremuneraciones = ""
        sC18_TotParticipUtilidades = ""
        sC19_RetencionImpuestoRenta = ""
        sC20_TotalDescuento = ""
        sC21_TotalPagar = ""
        sRetencionJudicial = ""
        sOtrosDescuentos = ""
    
            'Anidando los valores a registrar
            For I = 1 To grdCliente.Rows - 1

                sNomApell = sNomApell & Replace(grdCliente.TextMatrix(I, 1), ",", " ") & ","
                sNroDoc = sNroDoc & Replace(grdCliente.TextMatrix(I, 2), ",", " ") & ","
                sPeriodo = sPeriodo & Replace(grdCliente.TextMatrix(I, 3), ",", " ") & ","
                sMoneda = sMoneda & Replace(grdCliente.TextMatrix(I, 4), ",", " ") & ","
                sImporte = sImporte & Replace(grdCliente.TextMatrix(I, 5), ",", " ") & ","

                sCargo = sCargo & Replace(grdCliente.TextMatrix(I, 8), ",", " ") & ","
                sArea = sArea & Replace(grdCliente.TextMatrix(I, 9), ",", " ") & ","
                sfechaIngreso = sfechaIngreso & Replace(grdCliente.TextMatrix(I, 10), ",", " ") & ","
                
                sC09_ParticipAdistribuir = sC09_ParticipAdistribuir & Replace(grdCliente.TextMatrix(I, 11), ",", " ") & ","
                sC10_DiasLaborTodosTrabAnio = sC10_DiasLaborTodosTrabAnio & Replace(grdCliente.TextMatrix(I, 12), ",", " ") & ","
                sC11_RemunPercibTodosTrabAnio = sC11_RemunPercibTodosTrabAnio & Replace(grdCliente.TextMatrix(I, 13), ",", " ") & ","
                sC12_MontDistribXdiasLabor = sC12_MontDistribXdiasLabor & Replace(grdCliente.TextMatrix(I, 14), ",", " ") & ","
                sC13_MontDistribXremunPercib = sC13_MontDistribXremunPercib & Replace(grdCliente.TextMatrix(I, 15), ",", " ") & ","
                sC14_TotDiasEfectivLabor = sC14_TotDiasEfectivLabor & Replace(grdCliente.TextMatrix(I, 16), ",", " ") & ","
                sC15_TotRemuneraciones = sC15_TotRemuneraciones & Replace(grdCliente.TextMatrix(I, 17), ",", " ") & ","
                sC16_ParticipXdiasLabor = sC16_ParticipXdiasLabor & Replace(grdCliente.TextMatrix(I, 18), ",", " ") & ","
                sC17_ParticipXremuneraciones = sC17_ParticipXremuneraciones & Replace(grdCliente.TextMatrix(I, 19), ",", " ") & ","
                sC18_TotParticipUtilidades = sC18_TotParticipUtilidades & Replace(grdCliente.TextMatrix(I, 20), ",", " ") & ","
                sC19_RetencionImpuestoRenta = sC19_RetencionImpuestoRenta & Replace(grdCliente.TextMatrix(I, 21), ",", " ") & ","
                sC20_TotalDescuento = sC20_TotalDescuento & Replace(grdCliente.TextMatrix(I, 22), ",", " ") & ","
                sC21_TotalPagar = sC21_TotalPagar & Replace(grdCliente.TextMatrix(I, 23), ",", " ") & ","
                
                sRetencionJudicial = sRetencionJudicial & Replace(grdCliente.TextMatrix(I, 24), ",", " ") & ","
                sOtrosDescuentos = sOtrosDescuentos & Replace(grdCliente.TextMatrix(I, 25), ",", " ") & ","
                                
            Next I
            'quitando la coma del final
            If I > 1 Then
            
                sNomApell = Trim(Mid(sNomApell, 1, Len(sNomApell) - 1))
                sNroDoc = Trim(Mid(sNroDoc, 1, Len(sNroDoc) - 1))
                sPeriodo = Trim(Mid(sPeriodo, 1, Len(sPeriodo) - 1))
                sMoneda = Trim(Mid(sMoneda, 1, Len(sMoneda) - 1))
                sImporte = Trim(Mid(sImporte, 1, Len(sImporte) - 1))
                
                sCargo = Trim(Mid(sCargo, 1, Len(sCargo) - 1))
                sArea = Trim(Mid(sArea, 1, Len(sArea) - 1))
                sfechaIngreso = Trim(Mid(sfechaIngreso, 1, Len(sfechaIngreso) - 1))
                
                sC09_ParticipAdistribuir = Trim(Mid(sC09_ParticipAdistribuir, 1, Len(sC09_ParticipAdistribuir) - 1))
                sC10_DiasLaborTodosTrabAnio = Trim(Mid(sC10_DiasLaborTodosTrabAnio, 1, Len(sC10_DiasLaborTodosTrabAnio) - 1))
                sC11_RemunPercibTodosTrabAnio = Trim(Mid(sC11_RemunPercibTodosTrabAnio, 1, Len(sC11_RemunPercibTodosTrabAnio) - 1))
                sC12_MontDistribXdiasLabor = Trim(Mid(sC12_MontDistribXdiasLabor, 1, Len(sC12_MontDistribXdiasLabor) - 1))
                sC13_MontDistribXremunPercib = Trim(Mid(sC13_MontDistribXremunPercib, 1, Len(sC13_MontDistribXremunPercib) - 1))
                sC14_TotDiasEfectivLabor = Trim(Mid(sC14_TotDiasEfectivLabor, 1, Len(sC14_TotDiasEfectivLabor) - 1))
                sC15_TotRemuneraciones = Trim(Mid(sC15_TotRemuneraciones, 1, Len(sC15_TotRemuneraciones) - 1))
                sC16_ParticipXdiasLabor = Trim(Mid(sC16_ParticipXdiasLabor, 1, Len(sC16_ParticipXdiasLabor) - 1))
                sC17_ParticipXremuneraciones = Trim(Mid(sC17_ParticipXremuneraciones, 1, Len(sC17_ParticipXremuneraciones) - 1))
                sC18_TotParticipUtilidades = Trim(Mid(sC18_TotParticipUtilidades, 1, Len(sC18_TotParticipUtilidades) - 1))
                sC19_RetencionImpuestoRenta = Trim(Mid(sC19_RetencionImpuestoRenta, 1, Len(sC19_RetencionImpuestoRenta) - 1))
                sC20_TotalDescuento = Trim(Mid(sC20_TotalDescuento, 1, Len(sC20_TotalDescuento) - 1))
                sC21_TotalPagar = Trim(Mid(sC21_TotalPagar, 1, Len(sC21_TotalPagar) - 1))
                
                sRetencionJudicial = Trim(Mid(sRetencionJudicial, 1, Len(sRetencionJudicial) - 1))
                sOtrosDescuentos = Trim(Mid(sOtrosDescuentos, 1, Len(sOtrosDescuentos) - 1))
                
            End If
            'Procediendo con el registro de la trama
            If Len(sNomApell) > 0 And _
               Len(sNroDoc) > 0 And _
               Len(sPeriodo) > 0 And _
               Len(sMoneda) > 0 And _
               Len(sImporte) > 0 And _
               Len(Trim(sC09_ParticipAdistribuir)) > 0 And _
               Len(Trim(sC10_DiasLaborTodosTrabAnio)) > 0 And _
               Len(Trim(sC11_RemunPercibTodosTrabAnio)) > 0 And _
               Len(Trim(sC12_MontDistribXdiasLabor)) > 0 And _
               Len(Trim(sC13_MontDistribXremunPercib)) > 0 And _
               Len(Trim(sC14_TotDiasEfectivLabor)) > 0 And _
               Len(Trim(sC15_TotRemuneraciones)) > 0 And _
               Len(Trim(sC16_ParticipXdiasLabor)) > 0 And _
               Len(Trim(sC17_ParticipXremuneraciones)) > 0 And _
               Len(Trim(sC18_TotParticipUtilidades)) > 0 And _
               Len(Trim(sC19_RetencionImpuestoRenta)) > 0 And _
               Len(Trim(sC20_TotalDescuento)) > 0 And _
               Len(Trim(sC21_TotalPagar)) > 0 And _
               Len(Trim(sRetencionJudicial)) > 0 And _
               Len(Trim(sOtrosDescuentos)) > 0 Then

               bResultado = oPer.GrabarTramaUtilidades(sNomApell, sNroDoc, _
                                                       sPeriodo, sMoneda, _
                                                       sImporte, gsCodUser, _
                                                       sCargo, sArea, gsCodAge, sfechaIngreso, _
                                                       sC09_ParticipAdistribuir, sC10_DiasLaborTodosTrabAnio, _
                                                       sC11_RemunPercibTodosTrabAnio, sC12_MontDistribXdiasLabor, _
                                                       sC13_MontDistribXremunPercib, sC14_TotDiasEfectivLabor, _
                                                       sC15_TotRemuneraciones, sC16_ParticipXdiasLabor, sC17_ParticipXremuneraciones, _
                                                       sC18_TotParticipUtilidades, sC19_RetencionImpuestoRenta, _
                                                       sC20_TotalDescuento, sC21_TotalPagar, sRetencionJudicial, _
                                                       sOtrosDescuentos)
                Set oPer = Nothing
                Set oPer = Nothing
                
                If bResultado Then
                    LimpiarUtilidades
                    MsgBox "Se concluyó el proceso de carga satisfactoriamente", vbInformation, "Aviso"
                Else
                    MsgBox "Se presentó un error durante el proceso de carga", vbExclamation, "Aviso"
                End If
            End If
    End If
        
    Exit Sub
error_handler:
End Sub

Public Sub inicia(ByVal nIndex As Integer)
    
    nTipoOperacion = nIndex
    'cbPeriodoTrama.Height = txtArchivo.Height
    cbPeriodoTrama.Width = txtArchivo.Width
    cbPeriodoTrama.Top = txtArchivo.Top
    cbPeriodoTrama.Left = txtArchivo.Left
    nIdTrama = -1
    nEstadoCarga = -1
    bCerrar = False
    
    Select Case nIndex

        ' Cargar Trama
        Case 1
            txtArchivo.Visible = True
            cbPeriodoTrama.Visible = False
            grdCliente.ColWidth(6) = 0
            grdCliente.ColWidth(1) = 4500
            Me.Caption = "Trama Utilidades Ex Trabajadores"
        ' Baja trama
        Case 2
            txtArchivo.Visible = False
            cmdBuscar.Visible = False
            cmdCargar.Left = cmdCargar.Left - cmdBuscar.Width
            cbPeriodoTrama.Visible = True
            grdCliente.ColWidth(6) = 1200
            cmdGuardar.Caption = "Dar de Baja"
            Me.Caption = "Baja de Trama de Utilidades Ex Trabajadores"

    End Select
    Me.Show 1
End Sub

Private Sub cbPeriodo_Click()

On Error GoTo error_handler

    'solo efectua las consultas si la operacion es de baja de la trama
    If nTipoOperacion <> 2 Then
        Exit Sub
    End If
    
    Dim oPer As New COMDPersona.DCOMPersonas
    Dim rsUtilidad As ADODB.Recordset
    Dim bValidador As Boolean
    
    Set rsUtilidad = oPer.ObtenerTramaUtilidades(Val(Trim(Right(cbPeriodo.Text, 10))))
    bValidador = False
    If Not rsUtilidad Is Nothing Then
        If Not rsUtilidad.BOF And Not rsUtilidad.EOF Then
            bValidador = True
        Else
            bValidador = False
        End If
    Else
        bValidador = False
    End If
    If bValidador Then
        CargaCombo cbPeriodoTrama, rsUtilidad
    Else
        cbPeriodoTrama.Clear
    End If
    If cbPeriodoTrama.ListCount > 0 Then
        cbPeriodoTrama.ListIndex = 0
    End If
    
Exit Sub
error_handler:
MsgBox "Se presentó un error durante el proceso de selección", vbExclamation, "Aviso"
End Sub

Private Sub cmdBuscar_Click()
   On Error GoTo error_handler
    
    txtArchivo.Text = Empty
    
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Archivos de Excel (*.xls)|*.xls"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
        Exit Sub
    Else
        txtArchivo.Text = ""
    End If
    
    Exit Sub
error_handler:
    
    If err.Number = 32755 Then
    ElseIf err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        MsgBox "Error al momento de seleccionar el archivo", vbCritical, "Aviso"
    End If
End Sub

Private Sub cargarTramaExcel()
    
    Dim Col As Integer, fila As Integer, nTmp As Integer
    Dim psArchivoAGrabar As String, psArchivoAGrabarMenores As String, sClientes As String, sMontos As String, sMensaje As String
    Dim oExcel As Object, oBook As Object, oSheet As Object
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim bFormato As Boolean
    Dim oPer As COMDPersona.DCOMPersonas
    Dim rsPers As ADODB.Recordset
    
    sNroDoc = ""
    If Trim(txtArchivo.Text) = "" Then
        MsgBox "No selecciono ningun archivo", vbExclamation, "Aviso"
        Exit Sub
    End If
    grdCliente.lbEditarFlex = False
    pbProgres.Max = 10
    pbProgres.Min = 1
    pbProgres.value = 1
    pbProgres.Visible = True
    DoEvents
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(txtArchivo.Text)
    psArchivoAGrabar = App.Path & "\SPOOLER\NoCumpleValidacion_" & Format(gdFecSis, "yyyymmdd") & ".xls"
    grdCliente.SetFocus
    If Dir(psArchivoAGrabar) <> "" Then
        Kill psArchivoAGrabar
    End If
    
    pbProgres.value = 2
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    pbProgres.value = 3
    DoEvents
    bFormato = True
    
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 1))) <> "DNI" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 2))) <> "APELLIDO Y NOMBRE" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 3))) <> "IMPORTE" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 4))) <> "MONEDA" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 5))) <> "PERIODO UTILIDAD" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 6))) <> "CARGO" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 7))) <> "AREA" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 8))) <> "FECHA DE INGRESO" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 9))) <> "PARTICIPACION A DISTRIBUIR" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 10))) <> "DIAS LABORADOS POR TODOS LOS TRABAJADORES EN EL AÑO" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 11))) <> "REMUNERACION PERCIBIDA POR TODOS LOS TRABAJADORES EN EL AÑO" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 12))) <> "MONTO A DISTRIBUIR POR DIAS LABORADOS" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 13))) <> "MONTO A DISTRIBUIR POR REMUNERACIONES PERCIBIDAS" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 14))) <> "TOTAL DIAS EFECTIVAMENTE LABORADOS" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 15))) <> "TOTAL REMUNERACIONES" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 16))) <> "PARTICIPACION POR DIAS LABORADOS" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 17))) <> "PARTICIPACION POR REMUNERACIONES" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 18))) <> "TOTAL PARTICIPACION UTILIDADES" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 19))) <> "RETENCION IMPUESTO A LA RENTA" Then bFormato = False
    
    'CTI2 ADD 20190311 INI
    
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 20))) <> "RETENCION JUDICIAL" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 21))) <> "OTROS DESCUENTOS" Then bFormato = False
    
    'CTI2 ADD 20190311 FIN
    
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 22))) <> "TOTAL DESCUENTO" Then bFormato = False
    If UCase(Trim(xLibro.Sheets(1).Cells(1, 23))) <> "TOTAL PAGAR" Then bFormato = False
    
    If bFormato = False Then
        MsgBox "El archivo seleccionado no tiene el formato adecuado para la carga en lote, verifíquelo e inténtelo de nuevo", vbInformation, "Aviso"
        If Not objExcel Is Nothing Then
            objExcel.Workbooks.Close
            Set objExcel = Nothing
        End If
        If Not oExcel Is Nothing Then
            oExcel.Workbooks.Close
            Set oExcel = Nothing
        End If
        pbProgres.Visible = False
        Exit Sub
    End If
    'validando el contenido de la trama de excel ***********
    fila = 2
    Col = 1
    sMensaje = ""
    sNroDoc = ""
    
    pbProgres.value = 4
    DoEvents
    With xLibro
        With .Sheets(1)
            Do While Len(Trim(.Cells(fila, Col))) > 0
            
                'validando el DOI
                If Len(Trim(.Cells(fila, 1))) <> 8 Then
                    sMensaje = sMensaje & "El Doi del cliente " & Trim(.Cells(fila, 1)) & " debe tener 8 dígitos" & vbNewLine
                    nTmp = nTmp + 1
                End If
                
                'validando el importe a pagar
                If Not IsNumeric(Trim(.Cells(fila, 23))) Then
                    sMensaje = sMensaje & "El importe a pagar del cliente " & Trim(.Cells(fila, 1)) & " debe ser un valor numerico" & vbNewLine
                    nTmp = nTmp + 1
                End If
                
                'Validando la moneda
                If UCase(Trim(.Cells(fila, 4))) <> "SOLES" And UCase(Trim(.Cells(fila, 4))) <> "DOLARES" Then
                    sMensaje = sMensaje & "La moneda seleccionada por el cliente " & Trim(.Cells(fila, 1)) & " no está bien definida" & vbNewLine
                    nTmp = nTmp + 1
                End If
                
                'Validando la fecha de ingreso
                If Not IsDate(Trim(.Cells(fila, 8))) Then
                    sMensaje = sMensaje & "La fecha de ingreso no tiene " & Trim(.Cells(fila, 1)) & " no está bien definida" & vbNewLine
                    nTmp = nTmp + 1
                End If
                
                'Validando campos numericos
                If Not IsNumeric(Trim(.Cells(fila, 9))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 10))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 11))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 12))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 13))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 14))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 15))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 16))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 17))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 18))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 19))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 20))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 21))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 22))) Or _
                   Not IsNumeric(Trim(.Cells(fila, 23))) Then
                   
                    sMensaje = sMensaje & "En la Fila " & fila & " hay valores que deberían ser numéricos, pero no los son" & vbNewLine

                End If


                'If Val(Trim(.Cells(fila, 9))) <= 0 Then sMensaje = sMensaje & "El campo Nro 9, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 10))) <= 0 Then sMensaje = sMensaje & "El campo Nro 10, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 11))) <= 0 Then sMensaje = sMensaje & "El campo Nro 11, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 12))) <= 0 Then sMensaje = sMensaje & "El campo Nro 12, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 13))) <= 0 Then sMensaje = sMensaje & "El campo Nro 13, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 14))) <= 0 Then sMensaje = sMensaje & "El campo Nro 14, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 15))) <= 0 Then sMensaje = sMensaje & "El campo Nro 15, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 16))) <= 0 Then sMensaje = sMensaje & "El campo Nro 16, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 17))) <= 0 Then sMensaje = sMensaje & "El campo Nro 17, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 18))) <= 0 Then sMensaje = sMensaje & "El campo Nro 18, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 19))) <= 0 Then sMensaje = sMensaje & "El campo Nro 19, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 20))) <= 0 Then sMensaje = sMensaje & "El campo Nro 20, debe ser mayor que cero" & vbNewLine
                'If Val(Trim(.Cells(fila, 21))) <= 0 Then sMensaje = sMensaje & "El campo Nro 21, debe ser mayor que cero" & vbNewLine

                sNroDoc = sNroDoc & Trim(.Cells(fila, Col)) & ","
                fila = fila + 1
                If nTmp > 20 Then
                    Col = 200
                End If
                
            Loop
            sNroDoc = Trim(sNroDoc)
            If Len(sNroDoc) > 2 Then
                sNroDoc = Mid(sNroDoc, 1, Len(sNroDoc) - 1)
            End If
        End With
    End With
    
    If Len(Trim(sMensaje)) > 0 Or Len(Trim(sNroDoc)) = 0 Then
        If Len(Trim(sMensaje)) > 0 Then
            MsgBox "Durante la carga se presentaron las siguientes observaciones: " & vbNewLine & sMensaje, vbInformation, "Aviso"
        ElseIf Len(Trim(sNroDoc)) = 0 Then
            MsgBox "La trama seleccionada no contiene datos para la carga", vbInformation, "Aviso"
        End If
        If Not objExcel Is Nothing Then
            objExcel.Workbooks.Close
            Set objExcel = Nothing
        End If
        If Not oExcel Is Nothing Then
            oExcel.Workbooks.Close
            Set oExcel = Nothing
        End If
        pbProgres.Visible = False
        Exit Sub
    End If
    
    fila = 0
    Col = 0
    nTmp = 0
    sMensaje = ""
    ' Fin de la validacion **********
    
    pbProgres.value = 7
    DoEvents
    Set oPer = New COMDPersona.DCOMPersonas
    Set rsPers = oPer.ValidaTramaUtilidades(sNroDoc, Val(cbPeriodo.Text))
    sMensaje = Trim(validaRSUtilidad(rsPers))
    Set oPer = Nothing
    Set rsPers = Nothing
    pbProgres.value = 8
    DoEvents
    If Len(sMensaje) > 0 Then
        MsgBox "Se presentaron las siguientes observaciones: " & vbNewLine & sMensaje, vbInformation, "Aviso"
        pbProgres.value = 1
        If Not objExcel Is Nothing Then
            objExcel.Workbooks.Close
            Set objExcel = Nothing
        End If
        If Not oExcel Is Nothing Then
            oExcel.Workbooks.Close
            Set oExcel = Nothing
        End If
        pbProgres.Visible = False
        Exit Sub
    Else
        fila = 2
        Col = 1
        With xLibro
            With .Sheets(1)
                Do While Len(Trim(.Cells(fila, Col))) > 0
                    grdCliente.AdicionaFila
                    
                    'Nombre y Apellido
                    grdCliente.TextMatrix(fila - 1, 1) = UCase(Trim(.Cells(fila, 2)))
                    
                    'DNI
                    grdCliente.TextMatrix(fila - 1, 2) = UCase(Trim(.Cells(fila, 1)))
                    
                    'Periodo de Utilidad
                    grdCliente.TextMatrix(fila - 1, 3) = UCase(Trim(.Cells(fila, 5)))
                    
                    'Moneda
                    grdCliente.TextMatrix(fila - 1, 4) = UCase(Trim(.Cells(fila, 4)))
                    
                    'Importe
                    grdCliente.TextMatrix(fila - 1, 5) = Format(Trim(.Cells(fila, 23)), "##0.00")
                    
                    'Cargo
                    grdCliente.TextMatrix(fila - 1, 8) = Trim(.Cells(fila, 6))
                    
                    'Area
                    grdCliente.TextMatrix(fila - 1, 9) = Trim(.Cells(fila, 7))
                    
                    'Fecha de Ingreso
                    grdCliente.TextMatrix(fila - 1, 10) = Trim(.Cells(fila, 8))
                    
                    'Participacion a Distribuir
                    grdCliente.TextMatrix(fila - 1, 11) = Trim(.Cells(fila, 9))
                    
                    'Dias Laborados por Todos los Trabajadores en el Añio
                    grdCliente.TextMatrix(fila - 1, 12) = Trim(.Cells(fila, 10))
                                        
                    'Remuneracion Percibida por Todos los Trabajadores en el Año
                    grdCliente.TextMatrix(fila - 1, 13) = Trim(.Cells(fila, 11))
                                        
                    'Monto a Distribuir por dias laborados
                    grdCliente.TextMatrix(fila - 1, 14) = Trim(.Cells(fila, 12))
                    
                    'Monto a Distribuir por remuneraciones percibidas
                    grdCliente.TextMatrix(fila - 1, 15) = Trim(.Cells(fila, 13))
                    
                    'Total dias efectivamente laborados
                    grdCliente.TextMatrix(fila - 1, 16) = Trim(.Cells(fila, 14))
                    
                    'Total Remuneraciones
                    grdCliente.TextMatrix(fila - 1, 17) = Trim(.Cells(fila, 15))
                    
                    'Participacion por dias laborados
                    grdCliente.TextMatrix(fila - 1, 18) = Trim(.Cells(fila, 16))
                                        
                    'Participacion por remuneraciones
                    grdCliente.TextMatrix(fila - 1, 19) = Trim(.Cells(fila, 17))
                    
                    'Total Participacion Utilidades
                    grdCliente.TextMatrix(fila - 1, 20) = Trim(.Cells(fila, 18))
                    
                    'Retencion Impuesto a la Renta
                    grdCliente.TextMatrix(fila - 1, 21) = Trim(.Cells(fila, 19))
                    
                    'Total Descuento
                    grdCliente.TextMatrix(fila - 1, 22) = Trim(.Cells(fila, 22))
                    
                    'Total Pagar
                    grdCliente.TextMatrix(fila - 1, 23) = Trim(.Cells(fila, 23))
                                        
                    'Retención Judicial
                    grdCliente.TextMatrix(fila - 1, 24) = Trim(.Cells(fila, 20))
                    'Otros Descuentos
                    grdCliente.TextMatrix(fila - 1, 25) = Trim(.Cells(fila, 21))
                                                        
                    fila = fila + 1
                Loop
            End With
        End With
        'verifica si se cargaron datos en el grid y procede a bloquear el combo y otros controles
        If fila > 2 Then
            GridFormato
            cbPeriodo.Enabled = False
            nEstadoCarga = 1
        End If
    End If
    'Ultima validacion *******************
    sMensaje = sMensaje & validacion
    If Len(sMensaje) > 0 Then
        MsgBox "Se presentaron las siguientes observaciones: " & vbNewLine & sMensaje, vbInformation, "Aviso"
        pbProgres.value = 1
        If Not objExcel Is Nothing Then
            objExcel.Workbooks.Close
            Set objExcel = Nothing
        End If
        If Not oExcel Is Nothing Then
            oExcel.Workbooks.Close
            Set oExcel = Nothing
        End If
        pbProgres.Visible = False
        Exit Sub
    End If
    'End validacion **********************
    pbProgres.value = 9
    DoEvents
    pbProgres.value = 10
    DoEvents
    pbProgres.Visible = False
    objExcel.Quit
    Set objExcel = Nothing
    Set xLibro = Nothing
    Set oBook = Nothing
    oExcel.Quit
    Set oExcel = Nothing
End Sub

Private Sub cargarTramaBD()
    Dim oPer As COMDPersona.DCOMPersonas
    Dim rsPers As ADODB.Recordset
    Dim bResultado As Boolean
    If cbPeriodoTrama.ListCount > 0 Then
        Set oPer = New COMDPersona.DCOMPersonas
        Set rsPers = oPer.ObtenerTramaUtilidadesDetalle(Val(Trim(Right(cbPeriodoTrama.Text, 10))))
        If Not rsPers Is Nothing Then
            If Not rsPers.BOF And Not rsPers.EOF Then
                bResultado = True
            Else
                bResultado = False
            End If
        Else
            bResultado = False
        End If
        If bResultado Then
            nIdTrama = CLng(rsPers("nIdTrama"))
            grdCliente.rsFlex = rsPers
            GridFormato
            cbPeriodo.Enabled = False
            cbPeriodoTrama.Enabled = False
            nEstadoCarga = 1
        End If
    Else
        MsgBox "No se ha seleccionado alguna trama para la carga", vbInformation, "Aviso"
    End If
End Sub
Private Sub GridFormato()
    Dim I As Integer
    For I = 1 To grdCliente.Rows - 1
        grdCliente.TextMatrix(I, 5) = Format(CDbl(grdCliente.TextMatrix(I, 5)), "##0.00")
    Next I
End Sub
Private Sub cmdCargar_Click()
    
    On Error GoTo Error
    
    If Not verificaGridVacio Then
        If MsgBox("Al cargar la trama, se limpiaran los registros del Grid, ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        Else
            LimpiaFlex grdCliente
            nEstadoCarga = -1
        End If
    End If
    'Cargar trama de Excel
    If nTipoOperacion = 1 Then
        cargarTramaExcel
    Else
    'Cargar Trama de BD
        cargarTramaBD
    End If
    
    Exit Sub
Error:
  MsgBox "Se presentó un error durante el proceso de carga", vbExclamation, "Aviso"

End Sub

Private Function validaRSUtilidad(pRs As ADODB.Recordset) As String

On Error GoTo error_valida

Dim sMensaje As String
Dim nTmp As Integer

sMensaje = ""
nTmp = 0
If Not pRs Is Nothing Then
    If Not pRs.EOF And Not pRs.BOF Then
        pRs.MoveFirst
        Do While Not pRs.EOF
            'valida si el DOI del cliente está registrado en la BD
            If pRs("nEstaEnBase") = 0 Then
                sMensaje = sMensaje & "El DOI " & pRs("cDoi") & " no se encuentra registrado en la Base de datos" & vbNewLine
                nTmp = nTmp + 1
            End If
            'Valida si el cliente está activo o cesado
            If pRs("nActivo") = 1 Then
                'Se quitó validacion del estado del usuario a pedido de Operaciones.
                'sMensaje = sMensaje & "El cliente con DOI " & pRs("cDoi") & " se encuentra activo en la base de datos, verificar con Recursos Humanos" & vbNewLine
                nTmp = nTmp + 1
            End If
            'Valida si el cliente le pagaron sus utilidades
            If pRs("nRecibioUtilidad") = 1 Then
                sMensaje = sMensaje & "El cliente con DOI " & pRs("cDoi") & " ya recibió el pago por sus utilidades" & vbNewLine
                nTmp = nTmp + 1
            End If
            'Valida si el cliente ya ha sido cargado
            If pRs("nCargado") = 1 Then
                sMensaje = sMensaje & "El cliente con DOI " & pRs("cDoi") & " ya ha sido cargado en una trama anterior" & vbNewLine
                nTmp = nTmp + 1
            End If
            'Valida si hay duplicidad de DOI
            If pRs("nDobleDOI") = 1 Then
                sMensaje = sMensaje & "El cliente con DOI " & pRs("cDoi") & " se encuentra registrado mas de una vez en el sistema" & vbNewLine
                nTmp = nTmp + 1
            End If
            'Valida si el cliente ya ha sido cargado
            If pRs("nTrabajador") = 0 Then
                sMensaje = sMensaje & "El cliente con DOI " & pRs("cDoi") & " no está registrado como trabajador o ex trabajador en nuestra Base de Datos" & vbNewLine
                nTmp = nTmp + 1
            End If
            If nTmp > 20 Then
                Exit Do
            Else
                pRs.MoveNext
            End If
        Loop
    End If
End If
validaRSUtilidad = sMensaje
Exit Function
error_valida:
    validaRSUtilidad = "Se presento un error en el proceso de validacion de la trama"
End Function

Private Function verificaGridVacio() As Boolean

    If grdCliente.Rows = 2 And _
       Trim(grdCliente.TextMatrix(1, 1)) = "" And _
       Trim(grdCliente.TextMatrix(1, 2)) = "" And _
       Trim(grdCliente.TextMatrix(1, 3)) = "" And _
       Trim(grdCliente.TextMatrix(1, 4)) = "" And _
       Trim(grdCliente.TextMatrix(1, 5)) = "" Then
               
       verificaGridVacio = True
       
    Else
       verificaGridVacio = False
       
    End If

End Function

Private Function validacion() As String

On Error GoTo error_handler
    
    Dim sMensaje As String
    Dim I As Integer, e As Integer
    sMensaje = ""
    'validando en caso de dar de baja una trama ***************
    If nTipoOperacion = 2 Then
        If nIdTrama < 0 Then
            sMensaje = "Debes seleccionar una trama para continuar con el proceso de Baja de Trama" & vbNewLine
        End If
        If Not nEstadoCarga = 1 Then
            sMensaje = "La trama de utilidades debe de estar cargada" & vbNewLine
        End If
    End If
    'validando en caso de cargar una trama *********************
    If nTipoOperacion = 1 Then
        If grdCliente.Rows = 2 And verificaGridVacio Then
            validacion = "No existen registros para grabar la operación"
            Exit Function
        End If
        'Validando contenido de grilla
        I = 0
        e = 0
        For I = 1 To grdCliente.Rows - 1
            If Len(Trim(grdCliente.TextMatrix(I, 1))) = 0 Then
                sMensaje = sMensaje & "El campo ""Apellidos y Nombres"" del registro " & grdCliente.TextMatrix(I, 1) & " se encuentra vacio " & vbNewLine
                e = e + 1
            End If
            If Len(Trim(grdCliente.TextMatrix(I, 2))) <> 8 Then
                sMensaje = sMensaje & "El campo ""DNI"" del registro " & grdCliente.TextMatrix(I, 1) & " debe contener 8 dígitos " & vbNewLine
                e = e + 1
            End If
            If Not IsNumeric(Trim(grdCliente.TextMatrix(I, 5))) Then
                sMensaje = sMensaje & "El campo ""Importe"" del registro " & grdCliente.TextMatrix(I, 1) & " debe ser un valor numérico " & vbNewLine
                e = e + 1
            Else
                If CDbl(grdCliente.TextMatrix(I, 5)) <= 0 Then
                    sMensaje = sMensaje & "El valor del campo ""Importe"" del registro " & grdCliente.TextMatrix(I, 1) & " debe ser mayor o igual a 0 " & vbNewLine
                    e = e + 1
                End If
            End If
            If Not IsNumeric(Trim(grdCliente.TextMatrix(I, 3))) Then
                sMensaje = sMensaje & "El campo ""Periodo Utilidad"" del registro " & grdCliente.TextMatrix(I, 1) & " debe ser un valor numérico " & vbNewLine
                e = e + 1
            Else
                If Val(Trim(grdCliente.TextMatrix(I, 3))) <> Val(Trim(cbPeriodo.Text)) Then
                    sMensaje = sMensaje & "El campo ""Periodo Utilidad"" de la trama no coincide con el periodo seleccionado " & vbNewLine
                    e = e + 1
                End If
            End If
            If e > 20 Then
             Exit For
            End If
        Next I
    End If
    validacion = sMensaje
Exit Function
error_handler:
    validacion = "Se presento un error durante el proceso de validacion."

End Function

Private Sub cmdFormato_Click()

    Dim oBook As Object
    Dim oSheet As Object
    Dim sDireccion As String
    Dim oExcel As Object
    Dim bResult As Boolean
    Dim I As Long
    
    On Error GoTo error_handler
    
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    sDireccion = App.Path & "\SPOOLER\FormatoPagoUtilidades.xls"
    If Dir(sDireccion) <> "" Then
        Kill sDireccion
    End If
    oSheet.Range("A1:F1").Font.Bold = True
    oSheet.Columns("A:A").NumberFormat = "@"
    oSheet.Columns("B:B").NumberFormat = "@"
    oSheet.Columns("D:D").NumberFormat = "@"
    oSheet.Columns("E:E").NumberFormat = "@"
    oSheet.Columns("F:F").NumberFormat = "@"
    oSheet.Columns("G:G").NumberFormat = "@"
    
    oSheet.Columns("A:A").ColumnWidth = 10.71
    oSheet.Columns("B:B").ColumnWidth = 35
    oSheet.Columns("C:C").ColumnWidth = 11.43
    oSheet.Columns("D:D").ColumnWidth = 10.71
    oSheet.Columns("E:E").ColumnWidth = 17
    oSheet.Rows("2:2").RowHeight = 0
    
    
    oSheet.Range("A1").value = "DNI"
    oSheet.Range("B1").value = "APELLIDO Y NOMBRE"
    oSheet.Range("C1").value = "IMPORTE"
    oSheet.Range("D1").value = "MONEDA"
    oSheet.Range("E1").value = "PERIODO UTILIDAD"
    
    oSheet.Range("F1").value = "CARGO"
    oSheet.Range("G1").value = "AREA"
    oSheet.Range("H1").value = "FECHA DE INGRESO"
    oSheet.Range("I1").value = "PARTICIPACION A DISTRIBUIR"
    oSheet.Range("J1").value = "DIAS LABORADOS POR TODOS LOS TRABAJADORES EN EL AÑO"
    oSheet.Range("K1").value = "REMUNERACION PERCIBIDA POR TODOS LOS TRABAJADORES EN EL AÑO"
    oSheet.Range("L1").value = "MONTO A DISTRIBUIR POR DIAS LABORADOS"
    oSheet.Range("M1").value = "MONTO A DISTRIBUIR POR REMUNERACIONES PERCIBIDAS"
    oSheet.Range("N1").value = "TOTAL DIAS EFECTIVAMENTE LABORADOS"
    oSheet.Range("O1").value = "TOTAL REMUNERACIONES"
    oSheet.Range("P1").value = "PARTICIPACION POR DIAS LABORADOS"
    oSheet.Range("Q1").value = "PARTICIPACION POR REMUNERACIONES"
    oSheet.Range("R1").value = "TOTAL PARTICIPACION UTILIDADES"
    oSheet.Range("S1").value = "RETENCION IMPUESTO A LA RENTA"
    oSheet.Range("T1").value = "RETENCION JUDICIAL"
    oSheet.Range("U1").value = "OTROS DESCUENTOS"
    oSheet.Range("V1").value = "TOTAL DESCUENTO"
    oSheet.Range("W1").value = "TOTAL PAGAR"
 
    oBook.SaveAs sDireccion
    oExcel.Quit
    Set oExcel = Nothing
    Set oBook = Nothing
    
    Dim m_Excel As New Excel.Application
    m_Excel.Workbooks.Open (sDireccion)
    m_Excel.Visible = True
        
    Exit Sub
    
error_handler:
        oExcel.Quit
        Set oExcel = Nothing
        Set oBook = Nothing
        MsgBox "Error al generar el formato de pago", vbCritical, "Aviso"
        
        
End Sub


Private Sub LimpiarUtilidades()
    
    LimpiaFlex grdCliente
    cbPeriodo_Click
    txtArchivo.Text = ""
    cbPeriodo.Enabled = True
    cbPeriodoTrama.Enabled = True
    
    sNomApell = ""
    sNroDoc = ""
    sPeriodo = ""
    sMoneda = ""
    sImporte = ""
    
    sCargo = ""
    sArea = ""
    sfechaIngreso = ""
    
    sC09_ParticipAdistribuir = ""
    sC10_DiasLaborTodosTrabAnio = ""
    sC11_RemunPercibTodosTrabAnio = ""
    sC12_MontDistribXdiasLabor = ""
    sC13_MontDistribXremunPercib = ""
    sC14_TotDiasEfectivLabor = ""
    sC15_TotRemuneraciones = ""
    sC16_ParticipXdiasLabor = ""
    sC17_ParticipXremuneraciones = ""
    sC18_TotParticipUtilidades = ""
    sC19_RetencionImpuestoRenta = ""
    sC20_TotalDescuento = ""
    sC21_TotalPagar = ""
    
    nIdTrama = -1
    nEstadoCarga = -1
    If cbPeriodo.Enabled Then cbPeriodo.SetFocus
    
End Sub

Private Sub cmdLimpiar_Click()
    LimpiarUtilidades
End Sub

Private Sub CmdSalir_Click()
    If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        bCerrar = True
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_Load()
    cargarControles
End Sub

Private Sub cargarControles()

Dim dActual As Date
Dim nAnios As Integer, I As Integer
dActual = Now
nAnios = 10
dActual = DateAdd("yyyy", -nAnios, dActual)
For I = 0 To nAnios
    Call cbPeriodo.AddItem(DatePart("yyyy", dActual), 0)
    dActual = DateAdd("yyyy", 1, dActual)
Next I
pbProgres.Visible = False
cbPeriodo.ListIndex = 1

End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not bCerrar Then
        If MsgBox("¿Deseas salir de la formulario?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Cancel = 1
        Else
            bCerrar = True
        End If
    End If
End Sub
