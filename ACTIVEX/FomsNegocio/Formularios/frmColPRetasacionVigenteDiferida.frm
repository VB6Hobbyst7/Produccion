VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColPRetasacionVigenteDiferida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Retasación Vigentes/Diferidas/Adjudicadas"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18465
   Icon            =   "frmColPRetasacionVigenteDiferida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   18465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   15120
      TabIndex        =   29
      Top             =   7080
      Width           =   3255
      Begin VB.CommandButton cmdMiembrosRetas 
         Caption         =   "Comité"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   39
         ToolTipText     =   "Mostrar los miembros del comité de retasación"
         Top             =   380
         Width           =   990
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   380
         Width           =   975
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   380
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Carga de Archivo"
      Height          =   735
      Left            =   13200
      TabIndex        =   28
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdArchivo 
         Caption         =   "..."
         Height          =   360
         Left            =   3360
         TabIndex        =   1
         Top             =   235
         Width           =   480
      End
      Begin VB.TextBox txtArchivoRetas 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdValidar 
         Caption         =   "Validar"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fmRangoFechas 
      Enabled         =   0   'False
      Height          =   735
      Left            =   8040
      TabIndex        =   16
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox checkRangoFechas 
         Caption         =   "Buscar por rango de fechas"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   64356353
         CurrentDate     =   36161
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   10
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   64356353
         CurrentDate     =   36161
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   3960
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fmCriterioSeleccion 
      Caption         =   "Buscar por"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtBuscar 
         Height          =   315
         Left            =   4920
         TabIndex        =   7
         Top             =   285
         Width           =   2775
      End
      Begin VB.OptionButton optTasador 
         Caption         =   "Tasador"
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   340
         Width           =   975
      End
      Begin VB.OptionButton optNRetasacion 
         Caption         =   "N° Retasación"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   340
         Width           =   1335
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   340
         Width           =   855
      End
      Begin VB.OptionButton optCuenta 
         Caption         =   "Nº Contrato"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   340
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar Proceso"
      Height          =   375
      Left            =   15120
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkGrabarOnLine 
      Caption         =   "Grabar en Linea"
      Height          =   255
      Left            =   15120
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   18255
      Begin VB.TextBox txtRetas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   295
         Left            =   9000
         TabIndex        =   48
         Top             =   1080
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   1080
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.ComboBox cboKilataje 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
         Height          =   5895
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   18015
         _ExtentX        =   31776
         _ExtentY        =   10398
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedRows       =   2
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   2
         GridLinesFixed  =   1
         GridLinesUnpopulated=   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Totales"
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   7080
      Width           =   14895
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fantasia"
         Height          =   255
         Left            =   2280
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotFant 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   2280
         TabIndex        =   45
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblTot12k 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   4440
         TabIndex        =   44
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12 Kilates"
         Height          =   255
         Left            =   4440
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTot10k 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   3360
         TabIndex        =   42
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10 Kilates"
         Height          =   255
         Left            =   3360
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotLoteRetas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   13320
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Total Lote Retasación"
         Height          =   255
         Left            =   13080
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblTotPiezas 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   1200
         TabIndex        =   34
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "N° Piezas"
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTotLote 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   11280
         TabIndex        =   32
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Total Lote Preparación"
         Height          =   255
         Left            =   10920
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblTot21k 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   8760
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "21 kilates"
         Height          =   255
         Left            =   8760
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTot18k 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   7680
         TabIndex        =   25
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "18 kilates"
         Height          =   255
         Left            =   7680
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTot16k 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   6600
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "16 kilates"
         Height          =   255
         Left            =   6600
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTot14k 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "14 kilates"
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotalRetasado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Retasados"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgCarga 
      Left            =   14880
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmColPRetasacionVigenteDiferida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre       : frmColPRetasacionVigenteDiferida
'** Descripción  : Formulario para realizar la retazasion de creditos prendarios
'** Creación     : RECO, 20140707 - ERS074-2014
'** Modificación : TORE, 20180508 - ERS054-2017
'**                TORE, 20190601 - RFC1811260001
'**********************************************************************************************

Option Explicit
Dim xlsAplicacion As Excel.Application
Private rutaArchivo As String
Private rsMiembros As ADODB.Recordset 'TORE ERS054-2017
Private rsRetasacionBusqueda As ADODB.Recordset 'TORE ERS054-2017
Private rsRetas As ADODB.Recordset '[TORE ADD - Mejoras]
Private rsComite As ADODB.Recordset '[TORE ADD - Mejoras]
Dim lnFilaGuardar As Integer
Dim lnTpoBusq As Integer
Dim lnCodigoID As Long
Dim accionar As Integer
Dim Buscar As String
Dim lsCodigoID As String


'[TORE ADD - RFC1811260001]
Private Sub cmdMiembrosRetas_Click()
    Call MostrarMiembrosRetasacion
End Sub

Private Sub cboKilataje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fg.TextMatrix(cboKilataje.Tag, 12) = IIf(CInt(Trim(Right(cboKilataje.Text, 10))) = 0, "F", CStr(Trim(Right(cboKilataje.Text, 10))))
        fg.SetFocus
    End If
End Sub

Private Sub cboKilataje_LostFocus()
    cboKilataje.Visible = False
End Sub

Private Sub cmdArchivo_Click()
    dlgCarga.FileName = ""
    txtArchivoRetas.Text = ""
    dlgCarga.DialogTitle = "Abrir Excel de Preparación de Retasación"
    dlgCarga.Filter = "Archivos Excel(*.xls)|*.xls"
    dlgCarga.FilterIndex = 2
    dlgCarga.ShowOpen
    
    If dlgCarga.FileName <> "" Then
        txtArchivoRetas.Text = dlgCarga.FileName 'obtiene la ruta completa del archivo seleccionado
        rutaArchivo = dlgCarga.FileName
        Call CargarDatosPrepaRetasacionExcel
        cmdValidar.Enabled = True
        cmdMiembrosRetas.Enabled = True
        
        fmCriterioSeleccion.Enabled = True
        fmRangoFechas.Enabled = True
        cmdBuscar.Enabled = True
        
    ElseIf dlgCarga.FileName = "" Or dlgCarga.CancelError Then
        cmdValidar.Enabled = False
        cmdMiembrosRetas.Enabled = False
    End If
End Sub
Private Sub LimpiarFormulario(ByVal proceso As Integer)
    If proceso = 1 Then
         txtArchivoRetas.Text = ""
    End If
    fg.CellBackColor = RGB(255, 255, 255)
    cmdValidar.Enabled = False
    cmdMiembrosRetas.Enabled = False
    cmdBuscar.Enabled = False
    cmdGuardar.Enabled = False
    dtpDesde.Value = CDate(gdFecSis)
    dtpHasta.Value = CDate(gdFecSis)
    
    lblTotalRetasado.Caption = "0"
    lblTotPiezas.Caption = "0"
    lblTot14k.Caption = "0"
    lblTot16k.Caption = "0"
    lblTot18k.Caption = "0"
    lblTot21k.Caption = "0"
    lblTotLote.Caption = "0.00"
    lblTotLoteRetas.Caption = "0.00"
    
    
End Sub

Private Sub cmdBuscar_Click()
    If checkRangoFechas.Value = Checked Then
        FiltrarRetasaciones (2)
    End If
End Sub

Private Sub checkRangoFechas_Click()
    If checkRangoFechas.Value = 2 Then
        rsRetas.Filter = ""
    End If
End Sub

Private Sub cmdGuardar_Click()
    If MsgBox("¿Está ud. seguro de guardar los datos de la retasación?", vbQuestion + vbYesNo, "Confirmar") = vbNo Then Exit Sub
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    Dim nIndice As Integer
    Dim nCodigoID As Long
    Dim lsCredito As String
    Dim lsCreoError As String
    Dim nValorRetas As Integer
    
    Call CrearRSRetasaciones
    
    If fg.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos para retasar", vbInformation, "Alerta"
        Exit Sub
    End If
    lsCreoError = ""
    nCodigoID = CLng(lsCodigoID)
    For nIndice = 2 To fg.Rows - 1
        If fg.TextMatrix(nIndice, 1) <> "" _
            And fg.TextMatrix(nIndice, 21) <> "" _
            And fg.TextMatrix(nIndice, 22) <> "" _
            And fg.TextMatrix(nIndice, 12) <> "" _
            And fg.TextMatrix(nIndice, 4) <> "" Then
          
            Me.Caption = "Retasación Vigente/Diferida/Adjudicadas" & _
                         "- Retasación del lote " & lsCredito & "[" & fg.TextMatrix(nIndice, 22) & "] " & _
                         "[" & nIndice & " de " & fg.Rows - 1 & "]"
            
            'Guardamos los datos de la retasacion en la base de datos -- TORE
            
            nValorRetas = oNColP.RegistraRetasacion(nCodigoID, CStr(fg.TextMatrix(nIndice, 1)), _
                                 CInt(fg.TextMatrix(nIndice, 22)), CStr(fg.TextMatrix(nIndice, 12)), _
                                 CInt(fg.TextMatrix(nIndice, 4)), _
                                 IIf(fg.TextMatrix(nIndice, 14) = "", 0, fg.TextMatrix(nIndice, 14)), _
                                 IIf(fg.TextMatrix(nIndice, 15) = "", 0, fg.TextMatrix(nIndice, 15)), 0, _
                                 fg.TextMatrix(nIndice, 16), fg.TextMatrix(nIndice, 19), _
                                 IIf(fg.TextMatrix(nIndice, 18) = "", gdFecSis, fg.TextMatrix(nIndice, 18)))
            'Set oNColP = Nothing
            
            
            
'            If GuardarRetasacion(nCodigoID, CStr(fg.TextMatrix(nIndice, 1)), _
'                                 CInt(fg.TextMatrix(nIndice, 22)), CStr(fg.TextMatrix(nIndice, 12)), _
'                                 CInt(fg.TextMatrix(nIndice, 4)), _
'                                 IIf(fg.TextMatrix(nIndice, 14) = "", 0, CCur(fg.TextMatrix(nIndice, 14))), _
'                                 IIf(fg.TextMatrix(nIndice, 15) = "", 0, CCur(fg.TextMatrix(nIndice, 15))), 0, _
'                                 fg.TextMatrix(nIndice, 16), fg.TextMatrix(nIndice, 19), _
'                                 IIf(fg.TextMatrix(nIndice, 18) = "", gdFecSis, CDate(fg.TextMatrix(nIndice, 18)))) = 1 Then
    
'            ByVal pnCodigoID As Long,
'            ByVal psCtaCod As String,
'            ByVal pnItem As Integer,
'            ByVal psKilataje As String,
'            ByVal pnPiezas As Integer,
'            ByVal pnPesoBruto As Double,
'            ByVal pnPesoNeto As Double,
'            ByVal pnValTasac As Double,
'            ByVal psNroHolograma As String,
'            ByVal psObservaciones As String,
'            ByVal pdFechaRetas As Date
                                       
'            Else
'                MsgBox "No se registro la retasación del lote " & fg.TextMatrix(nIndice, 1) & Space(1) & "Item " & fg.TextMatrix(nIndice, 3) & _
'                vbNewLine & "", vbInformation, "Aviso"
'            End If
        Else
            lsCreoError = "ERROR"
        End If
        
        lsCredito = fg.TextMatrix(nIndice, 1)
        
    Next
    
    'Actualiza a la fecha de Retasacion si la retasacion se realizo correctamente
    If lsCreoError = "" Then
        'Dim oNColP As New COMNColoCPig.NCOMColPContrato
        oNColP.RegistraFechaRetasacion nCodigoID, gsCodUser
        Set oNColP = Nothing
        
        Call GuardarMiembrosRetasacion(nCodigoID)
    
    End If
    Set oNColP = Nothing
    Me.Caption = "Retasación Vigente/Diferida/Adjudicadas"
    MsgBox "Los datos se registraron con éxito.", vbInformation, "Alerta"
End Sub

Public Function GuardarRetasacion(ByVal pnCodigoID As Long, ByVal psCtaCod As String, ByVal pnItem As Integer, ByVal psKilataje As String, _
                              ByVal pnPiezas As Integer, ByVal pnPesoBruto As Double, ByVal pnPesoNeto As Double, ByVal pnValTasac As Double, _
                              ByVal psNroHolograma As String, ByVal psObservaciones As String, ByVal pdFechaRetas As Date) As Integer
    
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    GuardarRetasacion = oNColP.RegistraRetasacion(pnCodigoID, psCtaCod, pnItem, psKilataje, pnPiezas, pnPesoBruto, pnPesoNeto, pnValTasac, psNroHolograma, psObservaciones, pdFechaRetas)
    Set oNColP = Nothing
End Function
Private Function ValidarDatosRetasacionCarga() As Boolean
    Dim MatRetas() As String
    Dim lsTipRetas As String
    Dim nIndice As Integer
    ReDim MatRetas(1, 5)
    
    ValidarDatosRetasacionCarga = True
    If Trim(txtArchivoRetas.Text) = "" Then
        MsgBox "No se cargo ningún archivo de preparación de retasación", vbInformation, "Aviso"
        ValidarDatosRetasacionCarga = False
    End If
    
    For nIndice = 2 To fg.Rows - 1
        lsTipRetas = Mid(Trim(fg.TextMatrix(nIndice, 17)), 3, 2)
        If lsTipRetas = "01" Or lsTipRetas = "02" Then 'Validacion de los Vigentes/Diferidos
            If Trim(fg.TextMatrix(nIndice, 12)) = "F" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "10" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "12" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "14" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "16" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "18" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "21" Then
            Else
                MsgBox "El Kilataje de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no es reconocido por el sistema", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 12
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
            End If
            If Trim(fg.TextMatrix(nIndice, 12)) = "" Then
                MsgBox "El Kilataje de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 12
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
                'Exit Sub
            End If
            If Trim(Left(fg.TextMatrix(nIndice, 13), 15)) = "" Then
                MsgBox "El P. Total Lote de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 13
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
                'Exit Sub
            End If
            If fg.TextMatrix(nIndice, 14) = "" Then
                MsgBox "El P. Bruto de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 14
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
            End If
            If fg.TextMatrix(nIndice, 18) = "" Then
                MsgBox "La Fecha de retasación del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 18
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
            End If
'            If fg.TextMatrix(nIndice, 14) <> "" Then
'                MsgBox "El P. Bruto de la retasación del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío"
'                fg.SetFocus
'                fg.Col = 15
'                ValidarDatosRetasacionCarga = False
'                Exit For
'            End If
        ElseIf lsTipRetas = "03" Then 'Validacion de los Adjudicados
            If Trim(fg.TextMatrix(nIndice, 12)) = "F" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "10" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "12" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "14" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "16" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "18" Or _
                Trim(fg.TextMatrix(nIndice, 12)) = "21" Then
            Else
                MsgBox "El Kilataje de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no es reconocido por el sistema", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 12
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
            End If
            If Trim(fg.TextMatrix(nIndice, 12)) = "" Then
                MsgBox "El Kilataje de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 12
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
                'Exit Sub
            End If
            If Trim(Left(fg.TextMatrix(nIndice, 13), 15)) = "" Then
                MsgBox "El P. Total Lote de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 13
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
                'Exit Sub
            End If
            If fg.TextMatrix(nIndice, 14) = "" Then
                MsgBox "El P. Bruto de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 14
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                Exit For
            End If
            If fg.TextMatrix(nIndice, 15) = "" Then
                MsgBox "El P. Neto de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 15
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                Exit For
            End If
            If fg.TextMatrix(nIndice, 18) = "" Then
                MsgBox "La Fecha de retasación de la pieza " & Trim(fg.TextMatrix(nIndice, 5)) & " del lote " & fg.TextMatrix(nIndice, 1) & " no puede ser vacío", vbInformation, "Aviso"
                fg.SetFocus
                fg.Col = 18
                fg.row = nIndice
                fg.CellBackColor = RGB(255, 255, 153)
                ValidarDatosRetasacionCarga = False
                SendKeys "{Enter}"
                Exit For
            End If
            
        End If
   Next
    



'  If psColumna = 15 Then     'Peso Neto
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(psFila, 14)) * 0.1
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(psFila, 14)) - lnPesoNetoDesc
'        If FEDatos.TextMatrix(psFila, 15) <> "" Then
'            If CCur(FEDatos.TextMatrix(psFila, 15)) < 0 Then
'                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
'                FEDatos.TextMatrix(psFila, 15) = 0
'            Else
'                If CCur(FEDatos.TextMatrix(psFila, 15)) > lnPesoNetoDesc Then
'                    MsgBox "N° Crédito : " & FEDatos.TextMatrix(psFila, 1) & Chr(13) & _
'                            "Cliente : " & FEDatos.TextMatrix(psFila, 2) & Chr(13) & _
'                            "Joya: " & FEDatos.TextMatrix(psFila, 5) & Chr(13) & _
'                            "Peso Neto " & CCur(FEDatos.TextMatrix(psFila, 15)) & " debe ser menor a peso neto base " & lnPesoNetoDesc, vbInformation, _
'                    "Corrección de Retasación"
'                    FEDatos.TextMatrix(psFila, 15) = lnPesoNetoDesc
'                    'FEDatos.TextMatrix(psFila, 13) =
'                    'FEDatos.SetFocus
'                    'FEDatos.Col = 15
'                    'SendKeys "{Enter}"
'                Else
'                    'CalculaTasacion
'                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, Left(FEDatos.TextMatrix(psFila, 12), 2), 1)
'                        lnMatOro = Left(FEDatos.TextMatrix(psFila, 12), 2)
'                        Set loDR = loColContrato.PigObtenerValorTasacionxTpoClienteKt(ObtenerTipoCiente(FEDatos.TextMatrix(psFila, 1)))
'                        If Not (loDR.BOF And loDR.EOF) Then
'                            If lnMatOro = 14 Then
'                                lnValorPOro = loDR!n14kt
'                            ElseIf lnMatOro = 16 Then
'                                lnValorPOro = loDR!n16kt
'                            ElseIf lnMatOro = 18 Then
'                                lnValorPOro = loDR!n18kt
'                            ElseIf lnMatOro = 21 Then
'                                lnValorPOro = loDR!n21kt
'                            End If
'                        End If
'                        If lnPOro <= 0 Then
'                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
'                            Exit Function
'                        End If
'                        Set loColPCalculos = Nothing
'                        'Calcula el Valor de Tasacion
'                        CalcularValorTasacion = Format$(val(FEDatos.TextMatrix(psFila, 15) * lnValorPOro), "#####.00")
'                End If
'            End If
'        End If
'    End If
    'ValidarDatosRetasacionCarga = True
End Function
'TORE ERS054-2017
Private Sub GuardarMiembrosRetasacion(ByVal pnCodigoID As Integer)
    Dim i As Integer
    Dim rsMRetas As ADODB.Recordset
    Set rsMRetas = rsMiembros.Clone
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    For i = 1 To rsMRetas.RecordCount
        oNColP.RegistraMiembrosRetasacion pnCodigoID, i, rsMRetas!cNombreMiembro, rsMRetas!cRolMiembro
        rsMRetas.MoveNext
    Next i
    Set oNColP = Nothing
    
End Sub

Private Sub CrearRSRetasaciones() 'aqui
    Set rsRetas = New ADODB.Recordset
    rsRetas.CursorType = adOpenKeyset
    rsRetas.LockType = adLockBatchOptimistic
    rsRetas.CursorLocation = adUseClient
    'add campos
    rsRetas.Fields.Append "cCredito", adBSTR, 30, adFldUpdatable
    rsRetas.Fields.Append "cCliente", adBSTR, 255, adFldUpdatable
    rsRetas.Fields.Append "nTotPiezas", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "nDetPiezas", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "cDescJoya", adBSTR, 255, adFldUpdatable
    rsRetas.Fields.Append "nKilataje", adBSTR, 2, adFldUpdatable
    rsRetas.Fields.Append "cTasador", adBSTR, 4, adFldUpdatable
    rsRetas.Fields.Append "nPTotLote", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "nPBruto", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "nPNeto", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "cHolograma", adBSTR, 20, adFldUpdatable
    rsRetas.Fields.Append "nRKilataje", adBSTR, 2, adFldUpdatable
    rsRetas.Fields.Append "nRPTotLote", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "nRPBruto", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "nRPNeto", adBSTR, 10, adFldUpdatable
    rsRetas.Fields.Append "cRNroHolograma", adBSTR, 20, adFldUpdatable
    rsRetas.Fields.Append "cRNroRetasacion", adBSTR, 20, adFldUpdatable
    rsRetas.Fields.Append "cRFechaRetas", adBSTR, 50, adFldUpdatable
    rsRetas.Fields.Append "cRObservacion", adBSTR, 255, adFldUpdatable
    rsRetas.Fields.Append "nCodigoID", adBSTR, 50, adFldUpdatable
    rsRetas.Fields.Append "nItem", adBSTR, 5, adFldUpdatable
    rsRetas.Open
End Sub

Private Sub CrearRSComiteRetas()
    Set rsMiembros = New ADODB.Recordset
    rsMiembros.CursorType = adOpenKeyset
    rsMiembros.LockType = adLockBatchOptimistic
    rsMiembros.CursorLocation = adUseClient

    rsMiembros.Fields.Append "cNombreMiembro", adBSTR, 150, adFldUpdatable
    rsMiembros.Fields.Append "cRolMiembro", adBSTR, 150, adFldUpdatable
    rsMiembros.Open
End Sub

Private Sub GuardarRetasacionRecorSet(ByVal psCtaCod As String, ByVal pscliente As String, ByVal pTotPiezas As String, _
                                        ByVal psDetaPiezas As String, ByVal psDescripcion As String, ByVal psKilataje As String, _
                                        ByVal psTasador As String, ByVal psPTotLote As String, ByVal pnPBruto As String, _
                                        ByVal pnPNeto As String, ByVal psNroHolograma As String, ByVal psRKilataje As String, _
                                        ByVal psRPTotLote As String, ByVal psRPBruto As String, ByVal psRPNeto As String, _
                                        ByVal psRNroHolograma As String, ByVal psRNroRetasacion As String, ByVal psFechaRetas As String, _
                                        ByVal psRObservaciones As String, ByVal nCodigoID As String, ByVal nItem As String)

    rsRetas.AddNew
    rsRetas.Fields("cCredito").Value = psCtaCod
    rsRetas.Fields("cCliente").Value = pscliente
    rsRetas.Fields("nTotPiezas").Value = pTotPiezas
    rsRetas.Fields("nDetPiezas").Value = psDetaPiezas
    rsRetas.Fields("cDescJoya").Value = psDescripcion
    rsRetas.Fields("nKilataje").Value = psKilataje
    rsRetas.Fields("cTasador").Value = psTasador
    rsRetas.Fields("nPTotLote").Value = psPTotLote
    rsRetas.Fields("nPBruto").Value = pnPBruto
    rsRetas.Fields("nPNeto").Value = pnPNeto
    rsRetas.Fields("cHolograma").Value = psNroHolograma
    rsRetas.Fields("nRKilataje").Value = psRKilataje
    rsRetas.Fields("nRPTotLote").Value = psRPTotLote
    rsRetas.Fields("nRPBruto").Value = psRPBruto
    rsRetas.Fields("nRPNeto").Value = psRPNeto
    rsRetas.Fields("cRNroHolograma").Value = psRNroHolograma
    rsRetas.Fields("cRNroRetasacion").Value = psRNroRetasacion
    rsRetas.Fields("cRFechaRetas").Value = psFechaRetas
    rsRetas.Fields("cRObservacion").Value = psRObservaciones
    rsRetas.Fields("nCodigoID").Value = nCodigoID
    rsRetas.Fields("nItem").Value = nItem
End Sub

Private Sub GuardarComiteRecorSet(ByVal psNombre As String, ByVal psCargo As String)
    rsMiembros.AddNew
    rsMiembros.Fields("cNombreMiembro").Value = psNombre
    rsMiembros.Fields("cRolMiembro").Value = psCargo
End Sub

Private Sub HalitilarControles(ByVal Estado As Boolean)
    fmCriterioSeleccion.Enabled = Estado
    fmRangoFechas.Enabled = Estado
    cmdBuscar.Enabled = Estado
End Sub

'Public Function CalcularValorTasacion(ByVal psFila As Integer, ByVal psColumna As Integer) As Currency
'    Dim loColPCalculos As New COMDColocPig.DCOMColPCalculos
'    Dim loColContrato As New COMDColocPig.DCOMColPContrato
'    Dim loDR As New ADODB.Recordset
'    Dim lnPOro As Double
'    Dim lnValorPOro As Double
'    Dim lnMatOro As Integer
'    Dim lnPesoNetoDesc As Double
'
'    If psColumna = 15 Then     'Peso Neto
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(psFila, 14)) * 0.1
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(psFila, 14)) - lnPesoNetoDesc
'        If FEDatos.TextMatrix(psFila, 15) <> "" Then
'            If CCur(FEDatos.TextMatrix(psFila, 15)) < 0 Then
'                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
'                FEDatos.TextMatrix(psFila, 15) = 0
'            Else
'                If CCur(FEDatos.TextMatrix(psFila, 15)) > lnPesoNetoDesc Then
'                    MsgBox "N° Crédito : " & FEDatos.TextMatrix(psFila, 1) & Chr(13) & _
'                            "Cliente : " & FEDatos.TextMatrix(psFila, 2) & Chr(13) & _
'                            "Joya: " & FEDatos.TextMatrix(psFila, 5) & Chr(13) & _
'                            "Peso Neto " & CCur(FEDatos.TextMatrix(psFila, 15)) & " debe ser menor a peso neto base " & lnPesoNetoDesc, vbInformation, _
'                    "Corrección de Retasación"
'                    FEDatos.TextMatrix(psFila, 15) = lnPesoNetoDesc
'                    'FEDatos.TextMatrix(psFila, 13) =
'                    'FEDatos.SetFocus
'                    'FEDatos.Col = 15
'                    'SendKeys "{Enter}"
'                Else
'                    'CalculaTasacion
'                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, Left(FEDatos.TextMatrix(psFila, 12), 2), 1)
'                        lnMatOro = Left(FEDatos.TextMatrix(psFila, 12), 2)
'                        Set loDR = loColContrato.PigObtenerValorTasacionxTpoClienteKt(ObtenerTipoCiente(FEDatos.TextMatrix(psFila, 1)))
'                        If Not (loDR.BOF And loDR.EOF) Then
'                            If lnMatOro = 14 Then
'                                lnValorPOro = loDR!n14kt
'                            ElseIf lnMatOro = 16 Then
'                                lnValorPOro = loDR!n16kt
'                            ElseIf lnMatOro = 18 Then
'                                lnValorPOro = loDR!n18kt
'                            ElseIf lnMatOro = 21 Then
'                                lnValorPOro = loDR!n21kt
'                            End If
'                        End If
'                        If lnPOro <= 0 Then
'                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
'                            Exit Function
'                        End If
'                        Set loColPCalculos = Nothing
'                        'Calcula el Valor de Tasacion
'                        CalcularValorTasacion = Format$(val(FEDatos.TextMatrix(psFila, 15) * lnValorPOro), "#####.00")
'                End If
'            End If
'        End If
'    End If
'End Function
'TORE ERS054-2017
'Public Function CalcularValorRetasacion(ByVal psFila As Integer, ByVal psColumna As Integer) As Currency
'    Dim loColPCalculos As New COMDColocPig.DCOMColPCalculos
'    Dim loColContrato As New COMDColocPig.DCOMColPContrato
'    Dim loDR As New ADODB.Recordset
'    Dim lnPOro As Double
'    Dim lnValorPOro As Double
'    Dim lnMatOro As Integer
'    Dim lnPesoNetoDesc As Double
'
'    If psColumna = 15 Then     'Peso Neto
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(psFila, 14)) * 0.1
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(psFila, 14)) - lnPesoNetoDesc
'        If FEDatos.TextMatrix(psFila, 15) <> "" Then
'            If CCur(FEDatos.TextMatrix(psFila, 15)) < 0 Then
'                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
'                FEDatos.TextMatrix(psFila, 15) = 0
'            Else
'                If CCur(FEDatos.TextMatrix(psFila, 15)) > lnPesoNetoDesc Then
'                    MsgBox "N° Crédito : " & FEDatos.TextMatrix(psFila, 1) & Chr(13) & _
'                            "Cliente : " & FEDatos.TextMatrix(psFila, 2) & Chr(13) & _
'                            "Joya: " & FEDatos.TextMatrix(psFila, 5) & Chr(13) & _
'                            "Peso Neto " & CCur(FEDatos.TextMatrix(psFila, 15)) & " debe ser menor a peso neto base " & lnPesoNetoDesc, vbInformation, _
'                    "Corrección de Retasación"
'                    FEDatos.TextMatrix(psFila, 15) = lnPesoNetoDesc
'                    'CalcularValorRetasacion = lnPesoNetoDesc
''                Else
''                    'CalculaTasacion
''                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, Left(FEDatos.TextMatrix(psFila, 12), 2), 1)
''                        lnMatOro = Left(FEDatos.TextMatrix(psFila, 12), 2)
''                        Set loDR = loColContrato.PigObtenerValorTasacionxTpoClienteKt(ObtenerTipoCiente(FEDatos.TextMatrix(psFila, 1)))
''                        If Not (loDR.BOF And loDR.EOF) Then
''                            If lnMatOro = 14 Then
''                                lnValorPOro = loDR!n14kt
''                            ElseIf lnMatOro = 16 Then
''                                lnValorPOro = loDR!n16kt
''                            ElseIf lnMatOro = 18 Then
''                                lnValorPOro = loDR!n18kt
''                            ElseIf lnMatOro = 21 Then
''                                lnValorPOro = loDR!n21kt
''                            End If
''                        End If
''                        If lnPOro <= 0 Then
''                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
''                            Exit Function
''                        End If
''                        Set loColPCalculos = Nothing
''                        'Calcula el Valor de Tasacion
''                        CalcularValorTasacion = Format$(val(FEDatos.TextMatrix(psFila, 15) * lnValorPOro), "#####.00")
'                End If
'            End If
'        End If
'    End If
'End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdValidar_Click()
    If ValidarDatosRetasacionCarga() <> False Then
        Call ValidarDatosRetasacion
        accionar = 1
        cmdGuardar.Enabled = True
    End If
End Sub

Private Function ValidarArchio(ByVal psCodPreparacion As String) As Integer
Dim oNColP As New COMNColoCPig.NCOMColPContrato
'Dim nProcesoRetasacion As Integer
'nProcesoRetasacion = oNColP.ObtieneCodPrepaRetasacion(psCodPreparacion)
ValidarArchio = oNColP.ObtieneCodPrepaRetasacion(psCodPreparacion)
Set oNColP = Nothing
'If nProcesoRetasacion = 3 Then
'
'ElseIf nProcesoRetasacion = 2 Then
'    MsgBox "El archivo cargado ya fue retasado con anterioridad", vbInformation, "Aviso"
'    txtNombreArchivoRetasacion.Text = ""
'    txtNombreArchivoRetasacion.SetFocus
'ElseIf nProcesoRetasacion = 1 Then
'    MsgBox "Archivo no fue preparado para la retasación", vbInformation, "Aviso"
'End If
End Function

Private Sub CargarDatosPrepaRetasacionExcel()
    Call LimpiarFormulario(2)
    Dim oNColP As New COMNColoCPig.NCOMColPContrato
    Set xlsAplicacion = New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    
    Dim lsArvhivo As String
    Dim lsNombreHoja As String
    Dim lsExisteHoja As Boolean
    Dim Valido As Boolean
    Dim filas As Integer
    Dim filasR As Integer
    Dim filaRS As Integer
    
    Dim lsCtaCod As String
    Dim lsValorOrdenar As String
    Dim lsNroRetasacion As String
    
    Dim respuesta As VbMsgBoxResult
    
    If txtArchivoRetas.Text = "" Then
        MsgBox "No ha seleccionado el archivo a retasar", vbInformation, "Aviso"
        cmdArchivo.SetFocus
    End If
    
    lsNombreHoja = "Retasacion"
    lsArvhivo = Trim(rutaArchivo)
    
    Set xlsLibro = xlsAplicacion.Workbooks.Open(lsArvhivo)
    
    For Each xlHoja In xlsLibro.Worksheets
       If UCase(Trim(xlHoja.Name)) = UCase(Trim(lsNombreHoja)) Then
            xlHoja.Activate
            lsExisteHoja = True
        Exit For
       End If
    Next

    If lsExisteHoja = False Then
        MsgBox "El Nombre de la Hoja debe ser ''" & lsNombreHoja & "''", vbCritical, "Aviso"
        xlsAplicacion.Quit
        xlsAplicacion.Visible = False
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja = Nothing
        Exit Sub
    End If
  
    Valido = False
    
    Select Case ValidarArchio(QuitarExtension(xlHoja.Cells(4, 2)))
        Case 2
            Dim rpt
            rpt = MsgBox("Se reconoce que el archivo de la retasación ya fue cargado al sistema" & _
            ". Si guarda los datos, ud. estará reemplazando el registro de la retasación que fue cargada al sistema con anterioridad." & _
            vbNewLine & "¿Está seguro de continuar?", vbInformation + vbYesNo, "Confirmar")
            If rpt = 7 Then
                txtArchivoRetas.Text = ""
                txtArchivoRetas.SetFocus
                Call ExcelEnd(xlsAplicacion, xlsLibro, xlHoja)
                Exit Sub
            End If
        Case 1
            MsgBox "No se puede cargar el archivo porque no se reconoce el código de preparacion del archivo.", vbInformation, "Aviso"
            Call ExcelEnd(xlsAplicacion, xlsLibro, xlHoja)
            Exit Sub
    End Select
    
    Call CrearRSRetasaciones
    lsCodigoID = ObtenerCodPreparacion(xlHoja.Cells(4, 2))
    
    fg.Rows = 3
    filas = 1
    filasR = 5
    Do While Valido = False
        If filas <> 1 Then
            AdicionaRow fg
        End If
        filas = fg.row
        filasR = filasR + 1
        Me.Caption = "Cargando datos de retasación - [Lote " & xlHoja.Range("B" & filasR).Value & "]"
        fg.TextMatrix(filas, 1) = xlHoja.Range("B" & filasR).Value
        fg.TextMatrix(filas, 2) = xlHoja.Range("C" & filasR).Value & Space(100) & xlHoja.Range("B" & filasR).Value
        fg.TextMatrix(filas, 3) = Space(8) & (xlHoja.Range("E" & filasR).Value & Space(100) & xlHoja.Range("B" & filasR).Value)
        fg.TextMatrix(filas, 4) = xlHoja.Range("F" & filasR).Value
        fg.TextMatrix(filas, 5) = xlHoja.Range("G" & filasR).Value
        fg.TextMatrix(filas, 6) = xlHoja.Range("I" & filasR).Value
        fg.TextMatrix(filas, 7) = xlHoja.Range("H" & filasR).Value & Space(100) & xlHoja.Range("B" & filasR).Value
        fg.TextMatrix(filas, 8) = Space(8) & Format(xlHoja.Range("J" & filasR).Value, gcFormView) & Space(100) & xlHoja.Range("B" & filasR).Value
        fg.TextMatrix(filas, 9) = Format(xlHoja.Range("K" & filasR).Value, gcFormView)
        fg.TextMatrix(filas, 10) = Format(xlHoja.Range("L" & filasR).Value, gcFormView)
        fg.TextMatrix(filas, 11) = xlHoja.Range("M" & filasR).Value
        
        'Datos de la retasación
        lsNroRetasacion = oNColP.ObtenerCodRetasacion(lsCodigoID, Trim$(xlHoja.Range("B" & filasR).Value))
        
        fg.TextMatrix(filas, 12) = xlHoja.Range("N" & filasR).Value
        fg.TextMatrix(filas, 13) = Space(8) & Format(xlHoja.Range("Q" & filasR).Value, gcFormView) & Space(100) & xlHoja.Range("B" & filasR).Value
        fg.TextMatrix(filas, 14) = Format(xlHoja.Range("O" & filasR).Value, gcFormView)
        fg.TextMatrix(filas, 15) = Format(xlHoja.Range("P" & filasR).Value, gcFormView)
        fg.TextMatrix(filas, 16) = xlHoja.Range("R" & filasR).Value
        fg.TextMatrix(filas, 17) = lsNroRetasacion
        fg.TextMatrix(filas, 18) = xlHoja.Range("S" & filasR).Value
        fg.TextMatrix(filas, 19) = xlHoja.Range("T" & filasR).Value
        
        fg.TextMatrix(filas, 20) = "Ver Comite"
        fg.TextMatrix(filas, 21) = lsCodigoID
        fg.TextMatrix(filas, 22) = xlHoja.Range("D" & filasR).Value
        
        Call GuardarRetasacionRecorSet(xlHoja.Range("B" & filasR).Value, _
                                       (xlHoja.Range("C" & filasR).Value & Space(100) & xlHoja.Range("B" & filasR).Value), _
                                       (xlHoja.Range("E" & filasR).Value & Space(100) & xlHoja.Range("B" & filasR).Value), _
                                       xlHoja.Range("F" & filasR).Value, _
                                       xlHoja.Range("G" & filasR).Value, _
                                       xlHoja.Range("I" & filasR).Value, _
                                       (xlHoja.Range("H" & filasR).Value & Space(100) & xlHoja.Range("B" & filasR).Value), _
                                       (Format(xlHoja.Range("J" & filasR).Value, gcFormView) & Space(100) & xlHoja.Range("B" & filasR).Value), _
                                       Format(xlHoja.Range("K" & filasR).Value, gcFormView), _
                                       Format(xlHoja.Range("L" & filasR).Value, gcFormView), _
                                       xlHoja.Range("M" & filasR).Value, _
                                       xlHoja.Range("N" & filasR).Value, _
                                       (Format(xlHoja.Range("Q" & filasR).Value, gcFormView) & Space(100) & xlHoja.Range("B" & filasR).Value), _
                                       Format(xlHoja.Range("O" & filasR).Value, gcFormView), _
                                       Format(xlHoja.Range("P" & filasR).Value, gcFormView), _
                                       xlHoja.Range("R" & filasR).Value, _
                                       lsNroRetasacion, _
                                       xlHoja.Range("S" & filasR).Value, _
                                       xlHoja.Range("T" & filasR).Value, _
                                       lsCodigoID, _
                                       xlHoja.Range("D" & filasR).Value)
        
        
        If Trim(xlHoja.Range("G" & filasR + 1).Value) = "" Then
            Valido = True
        End If
    Loop
    
     lsExisteHoja = True
    
    'Carga de los miembros del cómite
     lsNombreHoja = "Comite"
     
     For Each xlHoja In xlsLibro.Worksheets
       If UCase(Trim(xlHoja.Name)) = UCase(Trim(lsNombreHoja)) Then
            xlHoja.Activate
            lsExisteHoja = True
            Exit For
       End If
    Next

    If lsExisteHoja = False Then
        MsgBox "El nombre de la hoja de trabajo debe ser ''" & lsNombreHoja & "'' obtener los datos del comite de retasación", vbCritical, "Aviso"
        xlsAplicacion.Quit
        xlsAplicacion.Visible = False
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlHoja = Nothing
        Exit Sub
    End If
    
    Valido = False
    filaRS = 1
    Call CrearRSComiteRetas
    Do While Valido = False
        filaRS = filaRS + 1
        Me.Caption = "Cargando datos de comité de retasación - [Item " & filaRS & "]"
        Call GuardarComiteRecorSet(xlHoja.Range("A" & filaRS + 1).Value, xlHoja.Range("B" & filaRS + 1).Value)
        If Trim(xlHoja.Range("A" & filaRS + 1).Value) = "" Or Trim(xlHoja.Range("B" & filaRS + 1).Value) = "" Then
           Valido = True
        End If
    Loop
    
    Call ExcelEnd(xlsAplicacion, xlsLibro, xlHoja)
    Me.Caption = "Retasación Vigente/Diferido/Adjudicado"
End Sub

Private Sub ExcelEnd(ByRef xlAplicacion As Excel.Application, ByRef xlLibro As Excel.Workbook, ByRef xlHoja As Excel.Worksheet)
    xlLibro.Close
    Sleep (800)
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja = Nothing
End Sub
Private Sub ValidarDatosRetasacion()
    lblTot14k.Caption = "0"
    lblTot16k.Caption = "0"
    lblTot18k.Caption = "0"
    lblTot21k.Caption = "0"
    lblTotPiezas.Caption = "0"
    lblTotalRetasado.Caption = "0"
    lblTotLote.Caption = "0.00"
    lblTotLoteRetas.Caption = "0.00"

    Dim lnFila As Integer, lsCredito As String, lsFilaTmp As Integer
    Dim va As Integer
    lnFila = 0
    For lnFila = 2 To fg.Rows - 1
        If fg.TextMatrix(lnFila, 1) <> "" And fg.TextMatrix(lnFila, 12) <> "" _
        And fg.TextMatrix(lnFila, 9) <> "" And fg.TextMatrix(lnFila, 14) <> "" Then
    
            If lsCredito <> fg.TextMatrix(lnFila, 1) Then
                lblTotalRetasado.Caption = Val(lblTotalRetasado.Caption) + 1
            End If
            lblTotLote.Caption = Val(lblTotLote.Caption) + fg.TextMatrix(lnFila, 9)
            lblTotLoteRetas.Caption = Val(lblTotLoteRetas.Caption) + fg.TextMatrix(lnFila, 14)
            lblTotPiezas.Caption = Val(lblTotPiezas.Caption) + fg.TextMatrix(lnFila, 4)
  
            Select Case CStr(Left(fg.TextMatrix(lnFila, 12), 2))
                Case "F"
                    lblTotFant.Caption = Val(lblTotFant.Caption) + 1
                Case "10"
                    lblTot10k.Caption = Val(lblTot10k.Caption) + 1
                Case "12"
                    lblTot12k.Caption = Val(lblTot12k.Caption) + 1
                Case "14"
                     lblTot14k.Caption = Val(lblTot14k.Caption) + 1
                Case "16"
                     lblTot16k.Caption = Val(lblTot16k.Caption) + 1
                Case "18"
                    lblTot18k.Caption = Val(lblTot18k.Caption) + 1
                Case "21"
                    lblTot21k.Caption = Val(lblTot21k.Caption) + 1
            End Select
        
            lsCredito = Trim(fg.TextMatrix(lnFila, 1))
            
        Else

            MsgBox "No se retasó la joya del Cliente" & Chr(13) & "Cliente: " _
                    & fg.TextMatrix(lnFila, 2) & Chr(13) _
                    & "N° Crédito: " & fg.TextMatrix(lnFila, 1) & Chr(13) _
                    & "No se guardará los datos de la retasación si los valores estan vacíos" _
                    , vbInformation, "Aviso"
            fg.SetFocus
            fg.row = lnFila
            fg.Col = 12
            SendKeys "{Enter}"
            Exit Sub
        End If
    Next
End Sub

Private Function QuitarExtension(ByVal psCodPreparacion As String) As String
    Dim punto As Integer
    punto = InStrRev(psCodPreparacion, "-")
    If punto > 0 Then
        QuitarExtension = Left$(psCodPreparacion, punto - 1)
    Else
        QuitarExtension = psCodPreparacion
    End If
End Function

Private Function ObtenerCodPreparacion(ByVal psCodPreparacion As String) As String
    Dim punto As Integer
    punto = InStrRev(psCodPreparacion, "-")
    If punto > 0 Then
        ObtenerCodPreparacion = Mid(psCodPreparacion, punto + 1)
    Else
        ObtenerCodPreparacion = psCodPreparacion
    End If
End Function

'Private Sub FEDatos_OnCellChange(pnRow As Long, pnCol As Long)
'    Dim lnPesoNetoDesc As Double
'    If pnCol = 14 Then 'Peso Bruto
'        If FEDatos.TextMatrix(FEDatos.row, 8) = "" Then
'            MsgBox "Ingrese un Peso Bruto Correcto", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(FEDatos.row, 14)) * 0.1
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(FEDatos.row, 14)) - lnPesoNetoDesc
'        FEDatos.TextMatrix(FEDatos.row, 15) = lnPesoNetoDesc
'    ElseIf pnCol = 15 Then     'Peso Neto
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(pnRow, 14)) * 0.1
'        lnPesoNetoDesc = CDbl(FEDatos.TextMatrix(pnRow, 14)) - lnPesoNetoDesc
'        If FEDatos.TextMatrix(pnRow, 15) <> "" Then
'            If CCur(FEDatos.TextMatrix(pnRow, 15)) < 0 Then
'                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
'                FEDatos.TextMatrix(pnRow, 15) = 0
'            Else
'                If CCur(FEDatos.TextMatrix(pnRow, 15)) > lnPesoNetoDesc Then
'                    MsgBox "Peso Neto " & CCur(FEDatos.TextMatrix(pnRow, 15)) & " debe ser menor a peso neto base " & lnPesoNetoDesc, vbInformation, "Aviso"
'                    FEDatos.TextMatrix(pnRow, 15) = lnPesoNetoDesc
'                End If
'            End If
'        End If
'    End If
'End Sub

Private Sub MostrarMiembrosRetasacion()
    frmColPMiembrosRetasacion.Inicio rsMiembros
End Sub

'Private Sub fg_DblClick()
'    If fg.Col = 12 Then
'        EnfocaTextoCombo cboKilataje, 0, fg
'    End If
'    If fg.Col = 14 Or fg.Col = 15 Then
'        EnfocaTexto txtRetas, 0, fg
'    End If
'    If fg.Col = 19 Then
'        EnfocaTexto txtObservacion, 0, fg
'    End If
'End Sub

'Private Sub fg_KeyPress(KeyAscii As Integer)
'     If fg.Col = 12 Then
'        'cboKilataje.Text = fg.TextMatrix(fg.RowSel, fg.Col)
'        'txtKilataje.Text = fg.TextMatrix(fg.RowSel, fg.Col)
'        'fg.Text = cboKilataje.Text
'        'Right(cboKilataje.Text, 5) =
'        EnfocaTextoCombo txtKilataje, cboKilataje, 0, fg
'    End If
'End Sub

Private Sub Form_Load()
    Call LimpiarFormulario(1)
    Call CargarComboFlex
    Call FormatoRetasacion
    txtBuscar = ""
    Buscar = "cCredito"
    lnTpoBusq = 1
End Sub

'TORE ESR054-2017
Private Sub optCuenta_Click()
    lnTpoBusq = 1
    Buscar = "cCredito"
    txtBuscar.MaxLength = 18
    txtBuscar = ""
    'txtBuscar.SetFocus
End Sub

Private Sub optTasador_Click()
    lnTpoBusq = 2
    Buscar = "cTasador"
    txtBuscar.MaxLength = 4
    txtBuscar = ""
    'txtBuscar.SetFocus
End Sub

Private Sub optNRetasacion_Click()
    lnTpoBusq = 3
    Buscar = "cRNroRetasacion"
    txtBuscar.MaxLength = 15
    txtBuscar = ""
    'txtBuscar.SetFocus
End Sub

Private Sub optCliente_Click()
    lnTpoBusq = 4
    Buscar = "cCliente"
    txtBuscar.MaxLength = 150
    txtBuscar = ""
    txtBuscar.SetFocus
End Sub
'END TORE

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    Dim nResult As Integer
    If KeyAscii = 13 Then
        Select Case lnTpoBusq
            Case 1 'cCredito
               FiltrarRetasaciones (1)
            Case 2 'cTasador
               FiltrarRetasaciones (1)
            Case 3 'cNroRetasacion
                FiltrarRetasaciones (1)
            Case 4 'cCliente
               FiltrarRetasaciones (1)
        End Select
    End If
End Sub

Private Sub FiltrarRetasaciones(Optional ByVal pnTipoFiltro As Integer = 0) 'aqui
    
    If pnTipoFiltro = 1 Then 'Credito
        If Trim(txtBuscar.Text) <> "" Then
            rsRetas.Filter = Buscar & " LIKE '*" + Trim(txtBuscar) + "*'"
        Else
            rsRetas.Filter = ""
        End If
    ElseIf pnTipoFiltro = 2 Then 'Rango de Fechas
       If checkRangoFechas.Value = Checked Then
            If dtpDesde.Value > dtpHasta.Value Then
                dtpDesde.Value = dtpHasta.Value - 1
                MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation, "Aviso"
                Exit Sub
            End If
            rsRetas.Filter = "cRFechaRetas >= " & FechaSQL(dtpDesde.Value) & _
                             " AND cRFechaRetas <= " & FechaSQL(dtpHasta.Value) & ""
        Else
            MsgBox "No se realizó la búsqueda, por favor seleccione el check para habilitar la búsqueda", vbInformation, "Aviso"
            rsRetas.Filter = ""
            Exit Sub
        End If
    Else
        rsRetas.Filter = ""
    End If
    
    Dim lsValorOrdenar As String
    Dim lnItem As Integer
    
    fg.Rows = 3
    lnItem = 1
    Do While Not rsRetas.EOF
         If lnItem <> 1 Then
            AdicionaRow fg
        End If
        lnItem = fg.row
        fg.TextMatrix(lnItem, 1) = rsRetas!cCredito
        fg.TextMatrix(lnItem, 2) = rsRetas!cCliente
        fg.TextMatrix(lnItem, 3) = Space(8) & rsRetas!nTotPiezas
        fg.TextMatrix(lnItem, 4) = rsRetas!nDetPiezas
        fg.TextMatrix(lnItem, 5) = rsRetas!cDescJoya
        fg.TextMatrix(lnItem, 6) = rsRetas!nKilataje
        fg.TextMatrix(lnItem, 7) = rsRetas!cTasador
        fg.TextMatrix(lnItem, 8) = Space(8) & Format(rsRetas!nPTotLote, gcFormView)
        fg.TextMatrix(lnItem, 9) = Format(rsRetas!nPBruto, gcFormView)
        fg.TextMatrix(lnItem, 10) = Format(rsRetas!nPNeto, gcFormView)
        fg.TextMatrix(lnItem, 11) = rsRetas!cHolograma

        'Datos de la retasación
        fg.TextMatrix(lnItem, 12) = rsRetas!nRKilataje
        fg.TextMatrix(lnItem, 13) = Space(8) & Format(rsRetas!nRPTotLote, gcFormView)
        fg.TextMatrix(lnItem, 14) = Format(rsRetas!nRPBruto, gcFormView)
        fg.TextMatrix(lnItem, 15) = Format(rsRetas!nRPNeto, gcFormView)
        fg.TextMatrix(lnItem, 16) = IIf(rsRetas!cRNroHolograma = "", "Sin Holograma", rsRetas!cRNroHolograma)
        fg.TextMatrix(lnItem, 17) = rsRetas!cRNroRetasacion
        fg.TextMatrix(lnItem, 18) = rsRetas!cRFechaRetas
        fg.TextMatrix(lnItem, 19) = rsRetas!cRObservacion

        'fg.TextMatrix(I, 20) = "Ver Comite"
        fg.TextMatrix(lnItem, 21) = rsRetas!nCodigoID
        fg.TextMatrix(lnItem, 22) = rsRetas!nItem

        rsRetas.MoveNext
    Loop



End Sub

Public Function FechaSQL(ByVal vFecha As String) As String
    On Local Error GoTo SQLDateValErr
    If IsDate(vFecha) Then
        FechaSQL = "#" & Format$(vFecha, "yyyy/mm/dd") & "#"
    Else
        FechaSQL = vFecha
    End If
   
    Exit Function
SQLDateValErr:
    Err = 0
    FechaSQL = "#1980/01/01#"
End Function

Private Sub FormatoRetasacion(Optional STIPOoc As String = "")
fg.Cols = 23
fg.TextMatrix(0, 0) = " "
fg.TextMatrix(1, 0) = " "

fg.TextMatrix(0, 1) = "DATOS DE LA JOYA"
fg.TextMatrix(0, 2) = "DATOS DE LA JOYA"
fg.TextMatrix(0, 3) = "DATOS DE LA JOYA"
fg.TextMatrix(0, 4) = "DATOS DE LA JOYA"
fg.TextMatrix(0, 5) = "DATOS DE LA JOYA"
fg.TextMatrix(0, 6) = "DATOS DE LA JOYA"
fg.TextMatrix(0, 7) = "DATOS DE LA JOYA"

fg.TextMatrix(1, 1) = "N° Contrato"
fg.TextMatrix(1, 2) = "Cliente"
fg.TextMatrix(1, 3) = "Tot. Pieza"
fg.TextMatrix(1, 4) = "Det. Pieza"
fg.TextMatrix(1, 5) = "Desc. Pieza"
fg.TextMatrix(1, 6) = "Kilataje"
fg.TextMatrix(1, 7) = "Tasador"

fg.TextMatrix(0, 8) = "VIGENTE/DIFERIDO/ADJUDICADO"
fg.TextMatrix(0, 9) = "VIGENTE/DIFERIDO/ADJUDICADO"
fg.TextMatrix(0, 10) = "VIGENTE/DIFERIDO/ADJUDICADO"
fg.TextMatrix(0, 11) = "VIGENTE/DIFERIDO/ADJUDICADO"

fg.TextMatrix(1, 8) = "P. Tot. Lote"
fg.TextMatrix(1, 9) = "P. Bruto"
fg.TextMatrix(1, 10) = "P. Neto"
fg.TextMatrix(1, 11) = "N° Holograma"

fg.TextMatrix(0, 12) = "RETASADO"
fg.TextMatrix(0, 13) = "RETASADO"
fg.TextMatrix(0, 14) = "RETASADO"
fg.TextMatrix(0, 15) = "RETASADO"
fg.TextMatrix(0, 16) = "RETASADO"
fg.TextMatrix(0, 17) = "RETASADO"
fg.TextMatrix(0, 18) = "RETASADO"
fg.TextMatrix(0, 19) = "RETASADO"
fg.TextMatrix(0, 20) = "RETASADO"

fg.TextMatrix(1, 12) = "Kilataje"
fg.TextMatrix(1, 13) = "Tot. Lote"
fg.TextMatrix(1, 14) = "P. Bruto"
fg.TextMatrix(1, 15) = "P. Neto"
fg.TextMatrix(1, 16) = "N° Holograma"
fg.TextMatrix(1, 17) = "N° Retasación"
fg.TextMatrix(1, 18) = "F. Retasación"
fg.TextMatrix(1, 19) = "Observaciones"
fg.TextMatrix(1, 20) = "Ver Comite"

fg.TextMatrix(0, 21) = "nCodigoID"
fg.TextMatrix(1, 21) = "nCodigoID"
fg.TextMatrix(0, 22) = "nItem"
fg.TextMatrix(1, 22) = "nItem"

fg.RowHeight(-1) = 285
fg.ColWidth(0) = 400
fg.ColWidth(1) = 2000
fg.ColWidth(2) = 4000
fg.ColWidth(3) = 850
fg.ColWidth(4) = 850
fg.ColWidth(5) = 3600
fg.ColWidth(6) = 850
fg.ColWidth(7) = 850
fg.ColWidth(8) = 1200
fg.ColWidth(9) = 1200
fg.ColWidth(10) = 1200
fg.ColWidth(11) = 1500
fg.ColWidth(12) = 1200
fg.ColWidth(13) = 1200
fg.ColWidth(14) = 1200
fg.ColWidth(15) = 1200
fg.ColWidth(16) = 1500
fg.ColWidth(17) = 1500
fg.ColWidth(18) = 1500
fg.ColWidth(19) = 5000
fg.ColWidth(20) = 0
fg.ColWidth(21) = 0
fg.ColWidth(22) = 0

fg.MergeCells = flexMergeRestrictColumns
fg.MergeCol(0) = True
fg.MergeCol(1) = True
fg.MergeCol(2) = True
fg.MergeCol(3) = True
fg.MergeCol(7) = True
fg.MergeCol(8) = True
fg.MergeCol(13) = True
fg.MergeCol(17) = True
fg.MergeCol(21) = True
fg.MergeCol(22) = True


fg.MergeRow(0) = True
fg.MergeRow(1) = True

fg.RowHeight(0) = 200
fg.RowHeight(1) = 200

fg.ColAlignmentFixed(-1) = flexAlignCenterCenter
fg.ColAlignment(1) = flexAlignCenterCenter
fg.ColAlignment(3) = flexAlignLeftCenter
fg.ColAlignment(4) = flexAlignCenterCenter
fg.ColAlignment(6) = flexAlignCenterCenter
fg.ColAlignment(8) = flexAlignLeftCenter
fg.ColAlignment(9) = flexAlignCenterCenter
fg.ColAlignment(10) = flexAlignCenterCenter
fg.ColAlignment(11) = flexAlignCenterCenter
fg.ColAlignment(12) = flexAlignCenterCenter
fg.ColAlignment(13) = flexAlignLeftCenter
fg.ColAlignment(14) = flexAlignCenterCenter
fg.ColAlignment(15) = flexAlignCenterCenter
fg.ColAlignment(16) = flexAlignCenterCenter
fg.ColAlignment(17) = flexAlignCenterCenter
fg.ColAlignment(18) = flexAlignCenterCenter
fg.ColAlignment(20) = flexAlignCenterCenter
End Sub

Public Sub CargarComboFlex()
    Dim rs As New ADODB.Recordset
    Dim oConst As New COMDConstantes.DCOMConstantes
    Set rs = oConst.RecuperaConstantes(3209)
    Call CargaCombo(rs, cboKilataje, 0, 1)
    Set rs = Nothing
    Set oConst = Nothing
End Sub

Public Sub CargaCombo(ByVal prsCombo As ADODB.Recordset, ByVal CtrlCombo As ComboBox, ByVal pnFiel1 As Integer, ByVal pnFiel2 As Integer)
    CtrlCombo.Clear
    While Not prsCombo.EOF
        CtrlCombo.AddItem prsCombo.Fields(pnFiel1) & Space(100) & prsCombo.Fields(pnFiel2) 'CROB20170721
        prsCombo.MoveNext
    Wend
End Sub

Public Function ObtenerTipoCiente(ByVal psPersCod As String) As Integer
    Dim loPigContrato As New COMDColocPig.DCOMColPContrato
    Dim poDR As New ADODB.Recordset
    Set poDR = loPigContrato.dVerificarCredPignoAdjudicado(psPersCod)
    If Not (poDR.BOF And poDR.EOF) Then
        ObtenerTipoCiente = 1
    Else
        Set poDR = Nothing
        Set poDR = loPigContrato.dVerificarCredPignoDesembolso(psPersCod)
        If Not (poDR.BOF And poDR.EOF) Then
            ObtenerTipoCiente = 2
        Else
            ObtenerTipoCiente = 1
        End If
    End If
End Function





