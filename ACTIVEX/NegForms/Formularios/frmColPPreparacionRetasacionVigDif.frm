VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColPPreparacionRetasacionVigDif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "    Preparación Retasación Vigentes/Diferidas/Adjudicadas"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmColPPreparacionRetasacionVigDif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmEstado 
      Caption         =   "Estado"
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   7335
      Begin VB.OptionButton optAdjudicadas 
         Caption         =   "Adjudicadas"
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDiferidas 
         Caption         =   "Diferidas"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optVigente 
         Caption         =   "Vigente"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   7335
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "Abrir"
         Height          =   375
         Left            =   6360
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin NegForms.FlexEdit FEListaCred 
         Height          =   2535
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4471
         Cols0           =   8
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Fecha-Total-Muestra-cAgeCod-cCtaCad-nIdCodPrepa"
         EncabezadosAnchos=   "400-2050-1200-1200-1200-0-3000-0"
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
         ColumnasAEditar =   "X-X-X-X-4-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
      Begin VB.Frame fmRango 
         Caption         =   "Rango"
         Height          =   1455
         Left            =   4680
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
         Begin VB.ComboBox cboTrimestre 
            Height          =   315
            ItemData        =   "frmColPPreparacionRetasacionVigDif.frx":030A
            Left            =   1080
            List            =   "frmColPPreparacionRetasacionVigDif.frx":031A
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   600
            Width           =   1035
         End
         Begin VB.OptionButton optFecha 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   1320
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optTrimestral 
            Caption         =   "Trimestral"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtAnio 
            Height          =   300
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   24
            Text            =   "txtAnio"
            Top             =   960
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker dtpDesde 
            Height          =   300
            Left            =   720
            TabIndex        =   28
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
            Format          =   51773441
            CurrentDate     =   36161
         End
         Begin MSComCtl2.DTPicker dtpHasta 
            Height          =   300
            Left            =   720
            TabIndex        =   29
            Top             =   960
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
            Format          =   51773441
            CurrentDate     =   36161
         End
         Begin VB.Label lblopt2 
            Caption         =   "lblopt2"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1005
            Width           =   615
         End
         Begin VB.Label lblopt1 
            Caption         =   "lblopt1"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   615
            Width           =   855
         End
      End
      Begin VB.Frame fmTpo 
         Caption         =   "Tipo"
         Height          =   975
         Left            =   4680
         TabIndex        =   13
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton optPersonalizado 
            Caption         =   "Personalizado"
            Height          =   375
            Left            =   1200
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton optMuestra 
            Caption         =   "Muestra"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   610
            Width           =   1095
         End
         Begin VB.OptionButton optTotal 
            Caption         =   "Total"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   250
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdPreparar 
         Caption         =   "Preparar"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Frame fmAgencias 
         Caption         =   "Agencias"
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   4335
         Begin VB.ListBox lsAgencias 
            Height          =   2310
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   600
            Width           =   4035
         End
         Begin VB.CheckBox chktodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin MSComctlLib.ProgressBar pbProgreso 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   7520
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmColPPreparacionRetasacionVigDif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPPreparacionRetasacionVigDif
'** Descripción : Formulario para realizar la preparación de la retazasion de creditos prendarios
'** Creación    : RECO, 20140707 - ERS074-2014
'**********************************************************************************************

Option Explicit
Dim nAccion As Integer
Dim fs As Scripting.FileSystemObject
Dim xlsAplicacion As Excel.Application

Private Sub cmdAbrir_Click()
    If FEListaCred.TextMatrix(FEListaCred.row, 1) = "" Then
        MsgBox "No se pudo generar archivo de la preparación, no se encontraron datos de la preparación de la retasación.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If nAccion <> 1 Then
        MsgBox "Antes de generar el excel debe de guardar la muestra", vbInformation, "Aviso"
        Exit Sub
    Else
        Dim oNColP As New COMNColoCPig.NCOMColPContrato
        Dim oDR As New ADODB.Recordset
        Dim i As Integer
        If FEListaCred.Rows > 1 And FEListaCred.TextMatrix(1, 1) <> "" Then
        Else
            MsgBox "No se encontraron datos", vbCritical, "Aviso"
            Exit Sub
        End If
        Screen.MousePointer = 11
        cmdAbrir.Enabled = False
        Set oDR = oNColP.DevuelveDatosListaCredPrepRetasacion(FEListaCred.TextMatrix(FEListaCred.row, 6), 1, 1)
        Call GeneraArchivo(oDR, FEListaCred.TextMatrix(FEListaCred.row, 4))
        Screen.MousePointer = 0
        cmdAbrir.Enabled = True
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdEditar_Click()
    Dim nMuestraEdit As Integer
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    Dim nCanMuestra As Integer
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    cmdEditar.Enabled = False
    Dim nEstadoJoy As Integer
    
     'TORE102042017 - ERS054-2017
    If optVigente.value = True Then
        nEstadoJoy = 2101 'Vigente
    End If
    If optDiferidas.value = True Then
        nEstadoJoy = 2102 'Diferido
    End If
    If optAdjudicadas.value = True Then
        nEstadoJoy = 2108 'Adjudicado
    End If
    'End TORE
   
    If FEListaCred.Rows > 1 And FEListaCred.TextMatrix(1, 1) <> "" Then
    Else
        MsgBox "No se encontraron datos", vbCritical, "Aviso"
        Exit Sub
    End If
    If optMuestra.value = True Then
        nMuestraEdit = InputBox("Ingrese la nueva cantidad de la muestra", "SICMACM")
        FEListaCred.TextMatrix(FEListaCred.row, 4) = nMuestraEdit
        nCanMuestra = FEListaCred.TextMatrix(FEListaCred.row, 4)
'    Else 'Comentado por TORE - ERS054-2017
'        frmColPPreparacionRetasacionPersVigDif.Inicio FEListaCred.TextMatrix(FEListaCred.row, 1), FEListaCred.TextMatrix(FEListaCred.row, 5), Format(dtpDesde.value, "yyyy/MM/dd"), Format(dtpHasta.value, "yyyy/MM/dd")
'        FEListaCred.TextMatrix(FEListaCred.row, 4) = frmColPPreparacionRetasacionPersVigDif.pnNumSelec
'        FEListaCred.TextMatrix(FEListaCred.row, 3) = frmColPPreparacionRetasacionPersVigDif.pnNumSelec
'        FEListaCred.TextMatrix(FEListaCred.row, 6) = frmColPPreparacionRetasacionPersVigDif.psCtaCad
'        cmdEditar.Enabled = True
    End If
    
    
    If Err > 0 Then
        MsgBox "Debe ingresar solo valores numéricos", vbCritical, "Aviso"
            Screen.MousePointer = 0
            cmdEditar.Enabled = True
        Exit Sub
    End If
    If optMuestra.value = True Then
        If nMuestraEdit < 1 Or nMuestraEdit > FEListaCred.TextMatrix(FEListaCred.row, 3) Then
            MsgBox "Valor fuera del rango. Debe ingresar un valor entre 1 y " & FEListaCred.TextMatrix(FEListaCred.row, 3), vbCritical, "Aviso"
            FEListaCred.TextMatrix(FEListaCred.row, 4) = nCanMuestra
            Screen.MousePointer = 0
            cmdEditar.Enabled = True
            Exit Sub
        End If
        FEListaCred.TextMatrix(FEListaCred.row, 6) = oNColP.DevuelveCtaCodCredPRetasacion(FEListaCred.TextMatrix(FEListaCred.row, 5), FEListaCred.TextMatrix(FEListaCred.row, 4), nEstadoJoy, Format(dtpDesde.value.Text, "yyyy/MM/dd"), Format(dtpHasta.value, "yyyy/MM/dd"), Right(cboTrimestre.ListIndex, 15), Trim(txtAnio.Text))
    End If
    Screen.MousePointer = 0
    cmdEditar.Enabled = True
End Sub

Private Sub CmdEliminar_Click()
    If FEListaCred.Rows > 1 And FEListaCred.TextMatrix(1, 1) <> "" Then
       FEListaCred.EliminaFila (FEListaCred.row)
    Else
        MsgBox "No se encontraron datos", vbCritical, "Aviso"
    End If
End Sub

Private Sub CmdGrabar_Click()
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    Dim i As Integer, J As Integer, ntpoPrerara As Integer, nCodigoID As Integer, nEstadoJoy As String, cRetasacion As String
    
    Dim lnCriterioPrep As Integer, ldFechaIni As String, ldFechaFin As String, lnTrimestre As Integer, lsAnio As String
    
    If ValidaDatosForm = False Then
        MsgBox "Los datos no pueden ser vacíos", vbInformation, "Alert"
        Exit Sub
    End If
    
    cmdGrabar.Enabled = False
    Screen.MousePointer = 11
    If optMuestra.value = True Then
        ntpoPrerara = 2
    ElseIf optTotal.value = True Then
        ntpoPrerara = 1
'Comentado por TORE10042018 - ERS054-2017
'    ElseIf optPersonalizado.value = True Then
'        ntpoPrerara = 3
    End If
    
     'TORE102042017 - ERS054-2017
    If optVigente.value = True Then
        nEstadoJoy = "01" 'Vigente
    End If
    If optDiferidas.value = True Then
        nEstadoJoy = "02" 'Diferido
    End If
    If optAdjudicadas.value = True Then
        nEstadoJoy = "03" 'Adjudicado
    End If
    'End TORE
    
    '[TORE RFC1811260001: ADD - Modificado con la finalidad de mejorar el filtro de busqueda de las preparaciones]
    If optTrimestral = True Then
        lnCriterioPrep = 1
        ldFechaIni = "01/01/1999"
        ldFechaFin = "01/01/1999"
    End If
    If optFecha = True Then
        lnCriterioPrep = 2
        ldFechaIni = Format(dtpDesde.value, "dd/MM/yyyy")
        ldFechaFin = Format(dtpHasta.value, "dd/MM/yyyy")
    End If
    
    lnTrimestre = Trim(Right(cboTrimestre.Text, 10))
    lsAnio = IIf(Trim(txtAnio.Text) = "", "----", Trim(txtAnio.Text))
    
    For i = 1 To FEListaCred.Rows - 1
        Dim h  As Integer
        Dim sCuenta As String
        Dim nNroRetasacion  As String
        nCodigoID = oNColP.RegistraPreparacionRetasacion(FEListaCred.TextMatrix(i, 5), _
                                                         ntpoPrerara, FEListaCred.TextMatrix(i, 4), _
                                                         FEListaCred.TextMatrix(i, 2), gsCodUser, _
                                                         nEstadoJoy, FEListaCred.TextMatrix(i, 3), _
                                                         lnCriterioPrep, ldFechaIni, ldFechaFin, _
                                                         lnTrimestre, lsAnio) '[TORE RFC1811260001: ADD FEListaCred.TextMatrix(i, 3)]
        h = 1
        For J = 1 To FEListaCred.TextMatrix(i, 4)
            If J = FEListaCred.TextMatrix(i, 4) Then
                sCuenta = Mid(FEListaCred.TextMatrix(i, 6), Len(FEListaCred.TextMatrix(i, 6)) - 17, 18)
            Else
                sCuenta = Mid(FEListaCred.TextMatrix(i, 6), h, 18)
            End If
            nNroRetasacion = FEListaCred.TextMatrix(i, 5) & nEstadoJoy & IIf(J >= 10, J, "0" & J) & Format(gdFecSis, "yyyy")
            oNColP.RegistraPreparacionRetasacionDet nCodigoID, sCuenta, nNroRetasacion
            'FEListaCred.TextMatrix(i, 6) = nCodigoID
            FEListaCred.TextMatrix(i, 7) = nCodigoID
            h = h + 19
        Next
    Next
    '[END TORE RFC1811260001]
    MsgBox "Los datos se guardaron con éxito", vbInformation, "Aviso"
    nAccion = 1
    cmdAbrir.Enabled = True
    'Call LimpiarFormulario
    Screen.MousePointer = 0
    cmdGrabar.Enabled = False
End Sub

Private Sub cmdPreparar_Click()
    Dim nRango As Integer
    Dim nTpoCarga As Integer
    Dim nEstadoJoy As Integer
    Dim nTrimiestre  As Integer
    Dim cAnio As String
    
    'TORE102042017 - ERS054-2017
    If optVigente.value = True Then
        nEstadoJoy = 2101 'Vigente
    End If
    If optDiferidas.value = True Then
        nEstadoJoy = 2102 'Diferido
    End If
    If optAdjudicadas.value = True Then
        nEstadoJoy = 2108 'Adjudicado
    End If
    'End TORE
    
    If optMuestra.value = True Then
        cmdEditar.Enabled = True
        nTpoCarga = 1
    ElseIf optTotal.value = True Then
        cmdEditar.Enabled = False
        nTpoCarga = 2
        'Comentado por TORE10042018 - ERS054-2017
        'ElseIf optPersonalizado.value = True Then
        'cmdEditar.Enabled = True
        'nTpoCarga = 3
    End If
    
    'TORE102042017 - ERS054-2017
    If optTrimestral.value = True Then
       If Len(Trim(Me.txtAnio)) = 0 Then
            MsgBox "Ingrese el año", vbCritical, "Alerta"
            'Call LimpiarFormulario
            'cboTrimestre.SetFocus
            Exit Sub
       End If
       If cboTrimestre.ListIndex = -1 Then
            MsgBox "Es necesario seleccionar el trimestre", vbCritical, "Alerta"
            'Call LimpiarFormulario
            'cboTrimestre.SetFocus
            Exit Sub
        End If
        If Len(Trim(Me.txtAnio.Text)) < 4 Or txtAnio.Text < 1900 Or txtAnio.Text > 2099 Then
            MsgBox "El año ingresado no es el correcto", vbCritical, "Alerta"
            'Call LimpiarFormulario
            'txtAnio.SetFocus
            Exit Sub
        End If

        nTrimiestre = cboTrimestre.ItemData(cboTrimestre.ListIndex)
        cAnio = txtAnio.Text
        nRango = 1
        
        Call CargarListaPrepara(nRango, RecuperaListaAgencias, nTpoCarga, nEstadoJoy, nTrimiestre, cAnio)
        
    ElseIf optFecha.value = True Then
        If dtpDesde.value = "__/__/____" Then
            MsgBox "Es necesario proporcionar la fecha de inicio", vbCritical, "Alerta"
            'Call LimpiarFormulario
            'optFecha.SetFocus
            dtpDesde.SetFocus
            Exit Sub
        End If
        If dtpHasta.value = "__/__/____" Then
             MsgBox "Es necesario proporcionar la fecha final", vbCritical, "Alerta"
            'Call LimpiarFormulario
            'optFecha.SetFocus
            dtpHasta.SetFocus
            Exit Sub
        End If
        nTrimiestre = 0
        cAnio = "2000"
        nRango = 2
        Call CargarListaPrepara(nRango, RecuperaListaAgencias, nTpoCarga, nEstadoJoy, nTrimiestre, cAnio)
    End If
    'End TORE
    'cmdPreparar.Enabled = True
    Screen.MousePointer = 0
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
    Call CargarListaAgencia
    Call LimpiarFormulario
End Sub
Public Sub CargarListaAgencia()
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    
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
Public Sub LimpiarFormulario()
    'TORE10042017
    cboTrimestre.ListIndex = 0
    optVigente.value = True
    optDiferidas.value = False
    optAdjudicadas.value = False
    
    optTrimestral.value = True
    optFecha.value = False
    txtAnio.Text = ""
    
    fmEstado.Enabled = True
    fmRango.Enabled = True
    'End TORE
    optTotal.value = True
    optMuestra.value = False
    'optPersonalizado.value = False 'Comentado por TORE10042018 - ERS054-2017
    FEListaCred.Clear
    FormateaFlex FEListaCred
    fmAgencias.Enabled = True
    fmTpo.Enabled = True
    cmdAbrir.Enabled = False
    
    Call LimpiarListaAge
    
    dtpDesde.value = Format(gdFecSis, "dd/MM/yyyy")
    dtpHasta.value = Format(gdFecSis, "dd/MM/yyyy")
    cmdPreparar.Enabled = True
    chktodos.value = 0
    Screen.MousePointer = 0
End Sub
Public Sub CargarListaPrepara(ByVal nRango As Integer, ByVal psAgeCad As String, ByVal nTpoCarga As Integer, ByVal nEstadoJoyas As Integer _
                                , ByVal pnTrimestre As Integer, ByVal psAnio As String)
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    Dim oDR As New ADODB.Recordset
    Dim sFecDesde As String
    Dim sFecHasta As String
    
    Dim i As Integer
    
    If psAgeCad = "0" Then
        MsgBox "Debe seleccionar al menos una agencia.", vbCritical, "Aviso"
        'Call LimpiarFormulario
        Exit Sub
    End If
    
    fmEstado.Enabled = False
    fmAgencias.Enabled = False
    fmTpo.Enabled = False
    fmRango.Enabled = False
    
    Screen.MousePointer = 11
    cmdPreparar.Enabled = False

    'TORE-ERS054-2017
    If nRango = 1 Then
        sFecDesde = Format("2000/01/01", "yyyy/MM/dd")
        sFecHasta = Format("2000/01/01", "yyyy/MM/dd")
        Set oDR = oNColP.ListaCredPrepRetasacion(psAgeCad, sFecDesde, sFecHasta, nEstadoJoyas, pnTrimestre, psAnio)
    Else
        sFecDesde = Format(dtpDesde.value, "yyyy/MM/dd")
        sFecHasta = Format(dtpHasta.value, "yyyy/MM/dd")
        Set oDR = oNColP.ListaCredPrepRetasacion(psAgeCad, sFecDesde, sFecHasta, nEstadoJoyas, pnTrimestre, psAnio)
    End If
    'End TORE
    
    FEListaCred.Clear
    FormateaFlex FEListaCred
    
    If Not (oDR.EOF And oDR.BOF) Then
        For i = 1 To oDR.RecordCount
            FEListaCred.AdicionaFila
            FEListaCred.TextMatrix(i, 1) = oDR!cAgeDescripcion
            FEListaCred.TextMatrix(i, 2) = gdFecSis
            FEListaCred.TextMatrix(i, 3) = oDR!nCant
            Select Case nTpoCarga
                Case 2
                    FEListaCred.TextMatrix(i, 4) = oDR!nCant
                Case 1
                    FEListaCred.TextMatrix(i, 4) = fnCalculaMuestra(oDR!nCant)
                'Comentado por TORE10012018 - ERS054-2017
                'Case 3
                'FEListaCred.TextMatrix(i, 3) = 0
                'FEListaCred.TextMatrix(i, 4) = 0
            End Select
            FEListaCred.TextMatrix(i, 5) = oDR!cAgeCod
            If nRango = 1 Then
                sFecDesde = Format("2000/01/01", "yyyy/MM/dd")
                sFecHasta = Format("2000/01/01", "yyyy/MM/dd")
                FEListaCred.TextMatrix(i, 6) = oNColP.DevuelveCtaCodCredPRetasacion(FEListaCred.TextMatrix(i, 5), FEListaCred.TextMatrix(i, 4), nEstadoJoyas, sFecDesde, sFecHasta, pnTrimestre, psAnio)
            Else
                sFecDesde = Format(dtpDesde.value, "yyyy/MM/dd")
                sFecHasta = Format(dtpHasta.value, "yyyy/MM/dd")
                FEListaCred.TextMatrix(i, 6) = oNColP.DevuelveCtaCodCredPRetasacion(FEListaCred.TextMatrix(i, 5), FEListaCred.TextMatrix(i, 4), nEstadoJoyas, sFecDesde, sFecHasta, pnTrimestre, psAnio)
            End If
            oDR.MoveNext
        Next
    Else
        'Call LimpiarFormulario
        cmdPreparar.Enabled = True
        MsgBox "No se encontraron datos para la preparación de la retasación", vbInformation, "Aviso"
    End If
    cmdGrabar.Enabled = True
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
Public Function fnCalculaMuestra(ByVal pnTotalLote As Integer) As Integer
    Dim n As Integer
    Dim Za As Double, P As Double, q As Double, d As Double, Res As Double, m As Double
    
    n = pnTotalLote
    Za = 2.58
    P = 0.7
    q = 1 - P
    d = 0.05
    
    Res = n * (Za ^ 2) * P * q
    m = Res / ((d ^ 2) * (n - 1) + (Za ^ 2) * P * q)
    fnCalculaMuestra = m
End Function

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

Public Sub GeneraArchivo(ByVal pDrDatos As ADODB.Recordset, ByVal nMuestreo As Integer)
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim lsArchivo As String, lsFile As String, lsNomHoja As String
    Dim lbExisteHoja As Boolean
    Dim lnValorConteo As Integer
    Dim i As Integer: Dim J As Integer

    'Dim lsCodPrepa As String
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
    
    pbProgreso.Min = 0
    pbProgreso.Max = pDrDatos.RecordCount
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
    
    

    xlHoja1.Cells(1, 1) = "Total de Lotes" & Space(16) & ":" & Space(5) & pDrDatos!nTotLotes
    xlHoja1.Cells(2, 1) = "Total de Muestra" & Space(12) & ":" & Space(5) & pDrDatos!nMuestra
 
    'lsCodPrepa = pDrDatos!cCodPrepacion
    xlHoja1.Cells(2, 3) = "LISTADO DE ORO RETASADO DE LA " & UCase(FEListaCred.TextMatrix(FEListaCred.row, 1))
    xlHoja1.Cells(3, 1) = "Fecha de Preparción" & Space(5) & ":" & Space(5) & Format(gdFecSis, "dd/MM/yyyy")
    xlHoja1.Cells(4, 2) = pDrDatos!cCodPrepacion & "-" & pDrDatos!nCodigoID
    i = 5
    'lnFilaTmp = 5
    Dim contarsi As Integer
    
    Do While Not pDrDatos.EOF
        i = i + 1
        pbProgreso.value = pbProgreso.value + 1
        If lsCtaCod <> pDrDatos!cPigCod Then
            lnTotPiezas = lnTotPiezas + pDrDatos!nTotPiezas
            lnOrden = lnOrden + 1
            xlHoja1.Cells(i, 1) = lnOrden
            lnPosicion = i + 1
            lnFilaTmp = 0
        Else
            lnFilaTmp = lnFilaTmp + 1
        End If
        

        lnPesoBruto = lnPesoBruto + pDrDatos!nPesoBruto
        lnPesoNeto = lnPesoNeto + pDrDatos!nPesoNeto
        
        'xlHoja1.Cells(i, 1) = pDrDatos!nNroRetasacion '[TORE RFC1811260001: Comentado segun RFC]
        xlHoja1.Cells(i, 2) = pDrDatos!cPigCod
        xlHoja1.Cells(i, 3) = pDrDatos!cPersNombre
        xlHoja1.Cells(i, 4) = pDrDatos!nItem
        xlHoja1.Cells(i, 5) = pDrDatos!nTotPiezas
        xlHoja1.Cells(i, 6) = pDrDatos!nPiezas
        xlHoja1.Cells(i, 7) = pDrDatos!cDescrip
        xlHoja1.Cells(i, 8) = pDrDatos!cUserTas
        xlHoja1.Cells(i, 9) = pDrDatos!cKilataje
        xlHoja1.Cells(i, 10) = Format(pDrDatos!nTotLote, gcFormView)
        xlHoja1.Cells(i, 11) = Format(pDrDatos!nPesoBruto, gcFormView)
        xlHoja1.Cells(i, 12) = Format(pDrDatos!nPesoNeto, gcFormView)
        xlHoja1.Cells(i, 13) = IIf(pDrDatos!nHolograma = 0, "Sin Holograma", pDrDatos!nHolograma)
        
        xlHoja1.Range("A" & Trim(Str(i)) & ":" & "T" & Trim(Str(i))).Borders.LineStyle = 1
        xlHoja1.Range("A" & Trim(Str(i)) & ":" & "S" & Trim(Str(i))).WrapText = True
        xlHoja1.Range("J" & Trim(Str(i)) & ":" & "M" & Trim(Str(i))).Interior.Color = RGB(204, 255, 255)
        xlHoja1.Range("N" & Trim(Str(i)) & ":" & "S" & Trim(Str(i))).Interior.Color = RGB(255, 255, 153)
        
        If lsCtaCod = pDrDatos!cPigCod Then
            xlHoja1.Range("A" & Trim(Str(lnPosicion - 1)) & ":" & "A" & Trim(Str(i))).MergeCells = True
'            xlHoja1.Range("C" & Trim(str(lnPosicion - 1)) & ":" & "C" & Trim(str(i))).MergeCells = True
'            xlHoja1.Range("E" & Trim(str(lnPosicion - 1)) & ":" & "E" & Trim(str(i))).MergeCells = True
'            xlHoja1.Range("H" & Trim(str(lnPosicion - 1)) & ":" & "H" & Trim(str(i))).MergeCells = True
'            xlHoja1.Range("J" & Trim(str(lnPosicion - 1)) & ":" & "J" & Trim(str(i))).MergeCells = True
'            xlHoja1.Range("Q" & Trim(str(lnPosicion - 1)) & ":" & "Q" & Trim(str(i))).MergeCells = True
        End If
        
        lsCtaCod = pDrDatos!cPigCod
        pDrDatos.MoveNext
        If pDrDatos.EOF Then
            Exit Do
        End If
    Loop
    
    lnValorConteo = i
    
    '[TORE RFC1811260001: ADD - Total de Piezas, Peso Bruto, Peso Neto]
    xlHoja1.Cells(i + 1, 5) = lnTotPiezas
    xlHoja1.Cells(i + 1, 11) = Format(lnPesoBruto, gcFormView)
    xlHoja1.Cells(i + 1, 12) = Format(lnPesoNeto, gcFormView)
    
    
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
    
    '[END TORE RFC1811260001: ADD - Total de Piezas, Peso Bruto, Peso Neto]
    
    pbProgreso.Visible = False
    
    xlHoja1.SaveAs App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
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
'END TORE
Public Function ValidaDatosForm() As Boolean
    ValidaDatosForm = True
    If RecuperaListaAgencias = "0" Then
        ValidaDatosForm = False
        Exit Function
    End If
    If FEListaCred.TextMatrix(1, 1) = "" Then
        ValidaDatosForm = False
        Exit Function
    End If
End Function

'TORE ERS054-2017
Private Sub ActivarControles(ByVal Estado As Boolean)
    Dim rangot As Boolean
    Dim RangoF As Boolean
    rangot = optTrimestral.value
    RangoF = optFecha.value
    
    If rangot = Estado Then
        lblopt1.Caption = "Trimestre:"
        lblopt2.Caption = "Año :"
        cboTrimestre.Visible = Estado
        txtAnio.Visible = Estado
        dtpDesde.Visible = Not Estado
        dtpHasta.Visible = Not Estado
    End If
    If RangoF = Estado Then
        lblopt1.Caption = "Desde :"
        lblopt2.Caption = "Hasta :"
        dtpDesde.Visible = Estado
        dtpHasta.Visible = Estado
        cboTrimestre.Visible = Not Estado
        txtAnio.Visible = Not Estado
    End If
    
End Sub

Private Sub optVigente_Click()
    optMuestra.Enabled = True
End Sub
Private Sub optDiferidas_Click()
    optMuestra.Enabled = True
End Sub
Private Sub optAdjudicadas_Click()
    optTotal.value = True
    optMuestra.Enabled = False
End Sub
Private Sub optTrimestral_Click()
    ActivarControles (True)
End Sub

Private Sub optFecha_Click()
    ActivarControles (True)
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
   KeyAscii = SoloNumerosTxt(KeyAscii)
End Sub

Private Function SoloNumerosTxt(ByVal KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumerosTxt = 0
    Else
        SoloNumerosTxt = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumerosTxt = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumerosTxt = KeyAscii 'Enter
End Function

'END TORE

