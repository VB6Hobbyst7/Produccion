VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPSubastaProceso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Subasta  de Lotes"
   ClientHeight    =   4560
   ClientLeft      =   2115
   ClientTop       =   2265
   ClientWidth     =   7245
   Icon            =   "frmColPSubastaProceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "Impresión"
      Height          =   585
      Left            =   210
      TabIndex        =   21
      Top             =   3855
      Width           =   2460
      Begin VB.OptionButton optImpresion 
         Caption         =   "Pantalla"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Impresora"
         Height          =   225
         Index           =   1
         Left            =   1230
         TabIndex        =   14
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5790
      TabIndex        =   15
      Top             =   3990
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   3825
      Left            =   210
      TabIndex        =   16
      Top             =   -15
      Width           =   6795
      Begin VB.CommandButton cmdSubasta 
         Caption         =   "Anterior..."
         Height          =   360
         Left            =   5640
         TabIndex        =   30
         Top             =   2910
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Frame Frame4 
         Caption         =   "Precios del Oro "
         Height          =   600
         Left            =   150
         TabIndex        =   23
         Top             =   570
         Width           =   6450
         Begin VB.TextBox txtPreOro14 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   855
            MaxLength       =   6
            TabIndex        =   4
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro16 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            MaxLength       =   6
            TabIndex        =   5
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro18 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3945
            MaxLength       =   6
            TabIndex        =   6
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro21 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   7
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label5 
            Caption         =   "21 Kl. :"
            Height          =   225
            Index           =   4
            Left            =   4995
            TabIndex        =   27
            Top             =   255
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "18 Kl. :"
            Height          =   225
            Index           =   3
            Left            =   3375
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "16 Kl. :"
            Height          =   225
            Index           =   2
            Left            =   1785
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "14 Kl. :"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdActSub 
         Caption         =   "Actas de Subasta"
         Enabled         =   0   'False
         Height          =   360
         Left            =   765
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   4650
      End
      Begin VB.CommandButton cmdPlaVenSub 
         Caption         =   "Planilla de Ventas en Subasta"
         Enabled         =   0   'False
         Height          =   360
         Left            =   765
         TabIndex        =   11
         Top             =   2760
         Visible         =   0   'False
         Width           =   4650
      End
      Begin VB.CommandButton cmdFinSisSub 
         Caption         =   "Finalizar Sistema de Subasta de Lotes"
         Enabled         =   0   'False
         Height          =   360
         Left            =   765
         TabIndex        =   10
         Top             =   2280
         Width           =   4650
      End
      Begin VB.CommandButton cmdRegVenSub 
         Caption         =   "Registrar Ventas en Subasta"
         Enabled         =   0   'False
         Height          =   360
         Left            =   765
         TabIndex        =   9
         Top             =   1800
         Width           =   4650
      End
      Begin VB.CommandButton cmdIniSisSub 
         Caption         =   "Inicializa Sistema para Subasta"
         Height          =   360
         Left            =   750
         TabIndex        =   8
         Top             =   1275
         Width           =   4650
      End
      Begin VB.TextBox txtNumSubasta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   795
         TabIndex        =   0
         Top             =   150
         Width           =   645
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Top             =   165
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtFecSubasta 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHorSubasta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   3900
         TabIndex        =   2
         Top             =   150
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   20
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   1485
         TabIndex        =   19
         Top             =   195
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   0
         Left            =   3405
         TabIndex        =   18
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Número :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   660
      End
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   5850
      TabIndex        =   22
      Top             =   4155
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmColPSubastaProceso.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   330
      Left            =   6345
      TabIndex        =   28
      Top             =   4125
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   582
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmColPSubastaProceso.frx":038A
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   330
      Left            =   2835
      TabIndex        =   29
      Top             =   4005
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmColPSubastaProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* SUBASTA DE CONTRATO PIGNORATICIO
'Archivo:  frmColPSubastaProceso.frm
'LAYG   :  15/05/2001.
'Resumen:  Muestra la Subasta actual y permite:
'   - Iniciar Subasta
'   - Registrar Ventas de Subasta de Lotes
'   - Finalizar Subasta
'   - Generar la planilla de ventas en Subasta.
'   - Generar el acta de Subasta
Option Explicit

Dim fnTasaPreparacionRemate As Double
Dim fnTasaIGV As Double
Dim fsProcesoCadaAgencia As String

Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim pAgeRemSub As String * 2
Dim pBandEOASub As Boolean
Dim pVerCodAnt As Boolean
Dim RegSubasta As New ADODB.Recordset
Dim RegCredPrend As New ADODB.Recordset


Dim RegJoyas As New ADODB.Recordset
Dim RegProcesar As New ADODB.Recordset
Dim sSQL As String
Dim MuestraImpresion As Boolean
Dim vRTFImp As String
Dim vCont As Double

'Muestra el formulario de genración de actas de la subasta
Private Sub cmdActSub_Click()
   'frmActaSubaPrend.Show 1
End Sub


'Procedimiento de Finalización de la subasta
Private Sub cmdFinSisSub_Click()

Dim loFinSub As COMNColoCPig.NCOMColPRecGar
Dim lsFecSubasta As String
Dim x As Integer

On Error GoTo ControlError
    
    lsFecSubasta = Format$(Me.txtFecSubasta, "mm/dd/yyyy")
    
    If MsgBox("Esta seguro de Finalizar el Proceso de Subasta ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
        '********** Finaliza el Remate
        
        Set loFinSub = New COMNColoCPig.NCOMColPRecGar
            Call loFinSub.nFinalizaSubasta(Trim(txtNumSubasta.Text), lsFecSubasta, fsProcesoCadaAgencia)
        Set loFinSub = Nothing
        
        txtEstado = "FINALIZADO"
        cmdIniSisSub.Enabled = False
        cmdRegVenSub.Enabled = False
        cmdFinSisSub.Enabled = False
        'cmdPlaVenSub.Enabled = True    'ARCV 24-06-2007
        'cmdActSub.Enabled = True
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Procedimiento de Inicialización de la Subasta
Private Sub cmdIniSisSub_Click()

On Error GoTo ControlError
Dim loIniSub As COMNColoCPig.NCOMColPRecGar
Dim lsFecSubasta As String
Dim x As Integer

    lsFecSubasta = Format$(Me.txtFecSubasta, "mm/dd/yyyy")
    
    If MsgBox("Esta seguro de Iniciar Subasta de Joyas ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
        '********** Realiza el Inicio
        
        Set loIniSub = New COMNColoCPig.NCOMColPRecGar
            Call loIniSub.nIniciaSubasta(Trim(txtNumSubasta.Text), lsFecSubasta, fnTasaIGV, Val(Me.txtPreOro14.Text), Val(Me.txtPreOro16.Text), Val(Me.txtPreOro18.Text), Val(Me.txtPreOro21.Text), fsProcesoCadaAgencia)
        Set loIniSub = Nothing
        
        txtEstado = "INICIADO"
        cmdIniSisSub.Enabled = False
        cmdRegVenSub.Enabled = True
        cmdFinSisSub.Enabled = True
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Llamada al formulario para registrar las Ventas de la subasta
Private Sub cmdRegVenSub_Click()
    frmColPSubastaRegVenta.Inicio (txtNumSubasta.Text)
End Sub

Private Sub cmdSubasta_Click()
'Dim vNroSub As String
'vNroSub = frmPigVarios.Inicio(2, txtNumSubasta.Text)
'sSql = " SELECT * FROM Subasta WHERE cnrosubas = '" & vNroSub & "'"
'RegSubasta.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'If (RegSubasta.BOF Or RegSubasta.EOF) Then
'    MsgBox " No existe el Remate solicitado ", vbInformation, " Aviso "
'Else
'    With RegSubasta
'        txtNumSubasta = !cnrosubas
'        txtFecSubasta = Format(!dfecsubas, "dd/mm/yyyy")
'        txtHorSubasta = Format(!dfecsubas, "hh:mm")
'        txtPreOro14 = Format(!nPreOro14, "#0.00")
'        txtPreOro16 = Format(!nPreOro16, "#0.00")
'        txtPreOro18 = Format(!nPreOro18, "#0.00")
'        txtPreOro21 = Format(!nPreOro21, "#0.00")
'        txtEstado = Switch(!cEstado = "N", "NO INICIADO", !cEstado = "I", "INICIADO", _
'                !cEstado = "F", "FINALIZADO")
'    End With
'End If
'RegSubasta.Close
'Set RegSubasta = Nothing
End Sub

'Inicializa el formulario actual y carga los datos del último subasta
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Call CargaParametros
    Call CargaDatosUltimaSubasta
End Sub

Private Sub CargaDatosUltimaSubasta()

Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrDatosSub As ADODB.Recordset
Dim lsUltSubasta As String
Dim lsMensaje As String

txtPreOro14 = Format(0, "#0.00")
txtPreOro16 = Format(0, "#0.00")
txtPreOro18 = Format(0, "#0.00")
txtPreOro21 = Format(0, "#0.00")

Set lrDatosSub = New ADODB.Recordset
Set loDatos = New COMNColoCPig.NCOMColPRecGar
    lsUltSubasta = loDatos.nObtieneNroUltimoProceso("S", fsProcesoCadaAgencia, lsMensaje)
    If Trim(lsMensaje) <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set lrDatosSub = loDatos.nObtieneDatosProcesoRGCredPig("S", lsUltSubasta, fsProcesoCadaAgencia, lsMensaje)
    If Trim(lsMensaje) <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    
    If lrDatosSub Is Nothing Then Exit Sub
    
    If lrDatosSub!nRGEstado = geColPRecGarEstNoIniciado And Format(lrDatosSub!dProceso, "dd/mm/yyyy") <> Format(gdFecSis, "dd/mm/yyyy") Then
        Set lrDatosSub = Nothing
        Set lrDatosSub = loDatos.nObtieneDatosProcesoRGCredPig("S", FillNum(Trim(Str(Val(lsUltSubasta) - 1)), 4, "0"), fsProcesoCadaAgencia, lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
Set loDatos = Nothing
'Mostrar Datos
If lrDatosSub Is Nothing Then Exit Sub
    txtNumSubasta = lrDatosSub!cNroProceso
    txtFecSubasta = Format(lrDatosSub!dProceso, "dd/mm/yyyy")
    txtHorSubasta = Format(lrDatosSub!dProceso, "hh:mm")
    txtPreOro14 = Format(lrDatosSub!nPrecioK14, "#0.00")
    txtPreOro16 = Format(lrDatosSub!nPrecioK16, "#0.00")
    txtPreOro18 = Format(lrDatosSub!nPrecioK18, "#0.00")
    txtPreOro21 = Format(lrDatosSub!nPrecioK21, "#0.00")
    
    If lrDatosSub!nRGEstado = gColPRecGarEstNoIniciado Then
        txtEstado = "NO INICIADO"
        If Val(txtPreOro14) = 0 And Val(txtPreOro16) = 0 And Val(txtPreOro18) = 0 And Val(txtPreOro21) = 0 Then
            cmdIniSisSub.Enabled = False
            MsgBox " Regrese a Preparacion de Subasta a ingresar precio del Oro ", vbInformation, " Aviso "
        ElseIf Format(lrDatosSub!dProceso, "dd/mm/yyyy") <> Format(gdFecSis, "dd/mm/yyyy") Then
            cmdIniSisSub.Enabled = False
            MsgBox " No es la fecha para el Remate ", vbInformation, " Aviso "
        End If
    ElseIf lrDatosSub!nRGEstado = gColPRecGarEstIniciado Then
        txtEstado = "INICIADO"
        cmdIniSisSub.Enabled = False
        cmdRegVenSub.Enabled = True
        cmdFinSisSub.Enabled = True
    ElseIf lrDatosSub!nRGEstado = gColPRecGarEstTerminado Then
        txtEstado = "FINALIZADO"
        cmdIniSisSub.Enabled = False
        cmdRegVenSub.Enabled = False
        cmdFinSisSub.Enabled = False
        'cmdPlaVenSub.Enabled = True
        'cmdActSub.Enabled = True
        'cmdSubasta.Visible = True
    End If
    Set lrDatosSub = Nothing
    
    'Inabilita cuando no sea la Agencia  que realiza el Remate
    'If Right(gsCodAge, 2) <> pAgeRemSub Then
    '    pBandEOARem = False
    '    cmdRegVenEOARem.Enabled = False
    'Else
    '    pBandEOARem = True
    'End If
 
End Sub

'Permite salir del formulario actual
Private Sub CmdSalir_Click()
    Unload Me
End Sub


'Valida el campo txtFecSubasta
Private Sub txtFecSubasta_GotFocus()
    fEnfoque txtFecSubasta
End Sub
Private Sub txtFecSubasta_LostFocus()
If Not ValFecha(txtFecSubasta) Then
    txtFecSubasta.SetFocus
End If
End Sub

'Valida el campo txtHorSubasta
Private Sub txtHorSubasta_GotFocus()
    fEnfoque txtHorSubasta
End Sub

'Procedimiento que permite direccionar la planilla de venta de subastas
' a un previo o a la impresora
Private Sub cmdPlaVenSub_Click()
'On Error GoTo ControlError
'    ImprimePlanVenSub
'    If MuestraImpresion And optImpresion(0).value = True Then
'        frmPrevio.Previo rtfImp, "Listado de Planilla", True, 66
'    End If
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

' Procedimiento que genera la planilla de venta del subasta
Public Sub ImprimePlanVenSub()
'    Dim vIndice As Integer  'contador de Item
'    Dim vLineas As Integer
'    Dim vPage As Integer
'    Dim vTelefono As String
'    Dim vSumTasaci As Currency
'    Dim vSumPresta As Currency
'    Dim vSumDeuda As Currency
'    Dim vSumBase As Currency
'    Dim vSumVenta As Currency
'    Dim vVerCodAnt As String
'    MousePointer = 11
'    MuestraImpresion = True
'    vRTFImp = ""
'    sSql = " SELECT cp.cCodCta, cp.dfecvenc, cp.nValTasac, " & _
'        " cp.nPrestamo, cp.nTasaIntVenc, cp.nImpuesto, " & _
'        " cp.nCostCusto, ds.ndeuda, ds.nprebasevta, ds.npreventa " & _
'        " FROM CredPrenda CP JOIN DetSubas DS ON cp.ccodcta = ds.ccodcta " & _
'        " WHERE ds.cestado IN ('V','P') AND ds.cNroSubas = '" & Trim(txtNumSubasta) & "'" & _
'        " ORDER BY ds.ccodcta"
'    RegCredPrend.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        MsgBox " No existen ningún contrato en la Subasta ", vbInformation, " Aviso "
'        MuestraImpresion = False
'        MousePointer = 0
'        Exit Sub
'    Else
'        vPage = 1
'        prgList.Min = 0: vCont = 0
'        prgList.Max = RegCredPrend.RecordCount
'        If optImpresion(0).value = True Then
'            If prgList.Max > pPrevioMax Then
'                RegCredPrend.Close
'                Set RegCredPrend = Nothing
'                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                MuestraImpresion = False
'                MousePointer = 0
'                Exit Sub
'            End If
'            Cabecera "PlanVeSu", vPage
'        Else
'            ImpreBegin True, pHojaFiMax
'            vRTFImp = ""
'            Cabecera "PlanVeSu", vPage
'            Print #ArcSal, ImpreCarEsp(vRTFImp);
'            vRTFImp = ""
'        End If
'        prgList.Visible = True
'        vIndice = 1:        vLineas = 7
'        vSumTasaci = 0:        vSumPresta = 0
'        vSumDeuda = 0:        vSumBase = 0
'        vSumVenta = 0
'        With RegCredPrend
'            Do While Not .EOF
'                'Para Ver el Código Antiguo
'                If pVerCodAnt Then vVerCodAnt = ContratoAntiguo(!cCodCta)
'                If optImpresion(0).value = True Then
'                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & Format(!dFecVenc, "dd/mm/yyyy") & ImpreFormat(Round(!nvaltasac, 2), 10) & ImpreFormat(Round(!nPrestamo, 2), 10) & _
'                        ImpreFormat(Round(!ndeuda, 2), 10) & ImpreFormat(Round(!nprebasevta, 2), 10) & ImpreFormat(Round(!npreventa, 2), 10) & Chr(10)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        vRTFImp = vRTFImp & Space(7) & ImpreFormat(vVerCodAnt, 13, 0) & Chr(10)
'                        vLineas = vLineas + 1
'                    End If
'                Else
'                    Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & Format(!dFecVenc, "dd/mm/yyyy") & ImpreFormat(Round(!nvaltasac, 2), 10) & ImpreFormat(Round(!nPrestamo, 2), 10) & _
'                        ImpreFormat(Round(!ndeuda, 2), 10) & ImpreFormat(Round(!nprebasevta, 2), 10) & ImpreFormat(Round(!npreventa, 2), 10)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        Print #ArcSal, Space(7) & ImpreFormat(vVerCodAnt, 13, 0)
'                        vLineas = vLineas + 1
'                    End If
'                End If
'                vLineas = vLineas + 1
'                vSumTasaci = vSumTasaci + Round(!nvaltasac, 2)
'                vSumPresta = vSumPresta + Round(!nPrestamo, 2)
'                vSumDeuda = vSumDeuda + Round(!ndeuda, 2)
'                vSumBase = vSumBase + Round(!nprebasevta, 2)
'                vSumVenta = vSumVenta + Round(!npreventa, 2)
'                vIndice = vIndice + 1
'                If vLineas >= 55 Then
'                    vPage = vPage + 1
'                    If optImpresion(0).value = True Then
'                        vRTFImp = vRTFImp & Chr(12)
'                        Cabecera "PlanVeSu", vPage
'                    Else
'                        If vPage Mod 5 = 0 Then
'                            ImpreEnd
'                            ImpreBegin True, pHojaFiMax
'                        Else
'                            ImpreNewPage
'                        End If
'                        vRTFImp = ""
'                        Cabecera "PlanVeSu", vPage
'                        Print #ArcSal, ImpreCarEsp(vRTFImp);
'                        vRTFImp = ""
'                    End If
'                    vLineas = 7
'                End If
'                vCont = vCont + 1
'                prgList.value = vCont
'                .MoveNext
'            Loop
'            If optImpresion(0).value = True Then
'                vRTFImp = vRTFImp & Chr(10)
'                vRTFImp = vRTFImp & Space(12) & "RESUMEN " & ImpreFormat(vSumTasaci, 20) & ImpreFormat(vSumPresta, 10) & _
'                    ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumBase, 10) & ImpreFormat(vSumVenta, 10) & Chr(10)
'                rtfImp.Text = vRTFImp
'            Else
'                Print #ArcSal, ""
'                Print #ArcSal, Space(12) & "RESUMEN " & ImpreFormat(vSumTasaci, 20) & ImpreFormat(vSumPresta, 10) & _
'                    ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumBase, 10) & ImpreFormat(vSumVenta, 10)
'                ImpreEnd
'            End If
'        End With
'        prgList.Visible = False
'        prgList.value = 0
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'    End If
'    MousePointer = 0
End Sub

'Cabecera de las Impresiones
Private Sub Cabecera(ByVal vOpt As String, ByVal vPagina As Integer, Optional ByVal pPagCorta As Boolean = True)
    Dim vTitulo As String
    Dim vSubTit As String
    Dim vArea As String * 30
    Dim vNroLineas As Integer
    Select Case vOpt
        Case "PlanVeSu"
            vTitulo = "PLANILLA DE PRENDAS DE LA SUBASTA N°: " & Format(txtNumSubasta, "@@@@") & " DEL " & txtFecSubasta
    End Select
    vSubTit = "(Después de la Subasta)"
    vArea = "Crédito Pignoraticio"
    vNroLineas = IIf(pPagCorta = True, 105, 162)
    'Centra Título
    vTitulo = String(Round((vNroLineas - Len(Trim(vTitulo))) / 2), " ") & vTitulo
    'Centra SubTítulo
    vSubTit = String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ") & vSubTit & String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ")

    vRTFImp = vRTFImp & Chr(10)
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(gsNomAge, 25, 0) & Space(vNroLineas - 40) & "Página: " & Format(vPagina, "@@@@") & Chr(10)
    vRTFImp = vRTFImp & Space(1) & vTitulo & Chr(10)
    vRTFImp = vRTFImp & Space(1) & vArea & vSubTit & Space(7) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & Chr(10)
    vRTFImp = vRTFImp & String(vNroLineas, "-") & Chr(10)
    Select Case vOpt
        Case "PlanVeSu"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO     FECHA        TASACION     CAPITAL        DEUDA        BASE        VENTA" & Chr(10)
            vRTFImp = vRTFImp & Space(1) & "                    VENCIMI.                   PRESTADO                    VENTA        NETA" & Chr(10)
    End Select
    vRTFImp = vRTFImp & String(vNroLineas, "-") & Chr(10)
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loParam = New COMDColocPig.DCOMColPCalculos
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
Set loParam = Nothing
    pPrevioMax = 5000
    pLineasMax = 56
    pHojaFiMax = 66
    'pAgeRemSub = Right(ReadVarSis("CPR", "cAgeRemSub"), 2)
    'pVerCodAnt = IIf(Left(ReadVarSis("CPR", "cVerCodAnt"), 1) = "S", True, False)
    
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsProcesoCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsProcesoCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing
End Sub
