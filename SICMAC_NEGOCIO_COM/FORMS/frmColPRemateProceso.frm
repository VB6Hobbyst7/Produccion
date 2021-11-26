VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPRemateProceso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio : Remate"
   ClientHeight    =   6330
   ClientLeft      =   1335
   ClientTop       =   1875
   ClientWidth     =   7185
   Icon            =   "frmColPRemateProceso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraImpresion 
      Caption         =   "Impresión"
      Height          =   540
      Left            =   210
      TabIndex        =   18
      Top             =   5700
      Width           =   2160
      Begin VB.OptionButton optImpresion 
         Caption         =   "Excel"
         Height          =   195
         Index           =   2
         Left            =   1125
         TabIndex        =   21
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Pantalla"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optImpresion 
         Caption         =   "Impresora"
         Height          =   225
         Index           =   1
         Left            =   3300
         TabIndex        =   19
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6030
      TabIndex        =   17
      Top             =   5805
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   5640
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   75
      Width           =   6795
      Begin VB.CommandButton CmdActa 
         Caption         =   "Acta de Remate"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   37
         Top             =   5160
         Width           =   4890
      End
      Begin VB.CommandButton cmdListAnt 
         Caption         =   "Con Cod Ant."
         Height          =   315
         Left            =   5580
         TabIndex        =   31
         Top             =   1815
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias..."
         Height          =   345
         Left            =   5640
         TabIndex        =   35
         Top             =   4560
         Width           =   1020
      End
      Begin VB.CommandButton cmdRemate 
         Caption         =   "Anterior..."
         Height          =   360
         Left            =   5760
         TabIndex        =   34
         Top             =   3525
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPlanillaNoVendido 
         Caption         =   "Planilla de Contratos No Vendidos en Remate"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   14
         Top             =   3720
         Width           =   4890
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Precios del Oro "
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
         Height          =   585
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Top             =   600
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
            TabIndex        =   26
            Top             =   210
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
            TabIndex        =   25
            Top             =   210
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
            TabIndex        =   24
            Top             =   210
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
            TabIndex        =   23
            Top             =   225
            Width           =   750
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "21 Kl. :"
            Height          =   225
            Index           =   4
            Left            =   4995
            TabIndex        =   30
            Top             =   255
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "18 Kl. :"
            Height          =   225
            Index           =   3
            Left            =   3375
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "16 Kl. :"
            Height          =   225
            Index           =   2
            Left            =   1785
            TabIndex        =   28
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "14 Kl. :"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   27
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdCartaSobrante 
         Caption         =   "Cartas de Sobrantes"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   16
         Top             =   4680
         Width           =   4890
      End
      Begin VB.CommandButton cmdActaRemate 
         Caption         =   "Listado al Detalle de  Remates "
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   15
         Top             =   4200
         Width           =   4890
      End
      Begin VB.CommandButton cmdPlanillaVentasRemate 
         Caption         =   "Planilla de Ventas en Remate"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   13
         Top             =   3240
         Width           =   4890
      End
      Begin VB.CommandButton cmdFinalizaRemate 
         Caption         =   "Finalizar Sistema de Remate de Joyas"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   12
         Top             =   2760
         Width           =   4890
      End
      Begin VB.CommandButton cmdRegistraVentaRemate 
         Caption         =   "Registrar Ventas en Remate"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   11
         Top             =   2280
         Width           =   4890
      End
      Begin VB.CommandButton cmdPlanillaContEntranRemate 
         Caption         =   "Planilla de Contratos que entran en Remate"
         Enabled         =   0   'False
         Height          =   360
         Left            =   405
         TabIndex        =   10
         Top             =   1800
         Width           =   4890
      End
      Begin VB.CommandButton cmdIniciaRemate 
         Caption         =   "Inicializa Sistema para Remate de Joyas"
         Height          =   360
         Left            =   405
         TabIndex        =   9
         Top             =   1320
         Width           =   4890
      End
      Begin VB.TextBox txtNumRemate 
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
         Left            =   900
         TabIndex        =   6
         Top             =   195
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
         Left            =   5370
         TabIndex        =   4
         Top             =   195
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txtFecRemate 
         Height          =   315
         Left            =   2175
         TabIndex        =   7
         Top             =   195
         Width           =   1170
         _ExtentX        =   2064
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
      Begin MSMask.MaskEdBox txtHorRemate 
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
         TabIndex        =   8
         Top             =   195
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.Label lblEtiqueta 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   5
         Left            =   4725
         TabIndex        =   5
         Top             =   210
         Width           =   645
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha :"
         Height          =   255
         Index           =   6
         Left            =   1590
         TabIndex        =   3
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   0
         Left            =   3405
         TabIndex        =   2
         Top             =   225
         Width           =   525
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Número :"
         Height          =   255
         Index           =   7
         Left            =   225
         TabIndex        =   1
         Top             =   225
         Width           =   660
      End
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   330
      Left            =   2775
      TabIndex        =   33
      Top             =   5910
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   360
      Left            =   225
      TabIndex        =   36
      Top             =   4770
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   635
      _Version        =   393217
      TextRTF         =   $"frmColPRemateProceso.frx":030A
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   870
      Left            =   210
      OleObjectBlob   =   "frmColPRemateProceso.frx":038D
      TabIndex        =   32
      Top             =   150
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmColPRemateProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REMATE - PROCESO.
'Archivo:  frmColPRemateProceso.frm
'LAYG   :  05/07/2001.
'Resumen:  Proceso de Remate de Creditos Pignoraticios
'          Inicializa / Imprime Planillas / Venta de JoyasRegistra precios de oro // Genera Cartas de Aviso Vencimiento // Aviso Remate
'   El formulario de Remate nos muestra el remate actual y nos permite
'   cerrar las operaciones con los contratos en remate y da opciones como:
'   - Iniciar el Remate
'   - Generar la planilla de contratos que entran en el remate
'   - Registrar las ventas del remate
'   - Finalizar el remate.
'   - Generar la planilla de ventas del remate cerrado.
'   - Generar el acta del remate.
'   - Generar las cartas de sobrantes.

'Variables del formulario
Option Explicit
Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim pAgeRemSub As String * 2
Dim pBandEOARem As Boolean

Dim fnTasaPreparacionRemate As Double
Dim fnFactorPrecioBaseRemate As Double
Dim fnDiasVctoParaRemate As Double
Dim fnTasaImpuesto As Double
Dim fnTasaCustodiaVencida As Double

Dim fsRemateCadaAgencia As String

Dim fnJoyasDet As Integer

'Dim pVerCodAnt As Boolean
'Dim pCtaAhoSob As String
'Dim pDifeDiasRema As Integer
'Dim RegRemate As New ADODB.Recordset
'Dim RegCredPrend As New ADODB.Recordset
'Dim RegPersona As New ADODB.Recordset
'Dim RegJoyas As New ADODB.Recordset
'Dim RegProcesar As New ADODB.Recordset
'Dim MuestraImpresion As Boolean
Dim vRTFImp As String
'Dim vCont As Double
Dim vFecAviso As Date, vFecRemNue As Date, vFecSis As Date

Private Sub CmdActa_Click()
    ImprimeActa
End Sub



Private Sub cmdAgencia_Click()
    'Selec Age. a realizar Remate
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdCartaSobrante_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    lsCadImprimir = ""
    
    rtfCartas.FileName = App.path & cPlantillaAvisoSobrante
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeAvisoSobranteRemate(Me.txtNumRemate.Text, Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), _
                        rtfCartas.Text, 66, gdFecSis, lsmensaje)
                        
                If Trim(lsmensaje) <> "" Then
                     MsgBox lsmensaje, vbInformation, "Aviso"
                     Exit Sub
                End If

            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Sobrante de Remate", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdFinalizaRemate_Click()
Dim loFinRem As COMNColoCPig.NCOMColPRecGar

Dim lsFecRemate As String
Dim X As Integer
Dim lsmensaje As String

'Dim PObjConec As DConstante
'Dim loConstSis As NConstSistemas
'Dim lrCtaAho As ADODB.Recordset
Dim lsCtaSobranteRemate As String

Dim loSobranteR As COMNColoCPig.NCOMColPRecGar
Dim lnMontoSobrante As Currency

On Error GoTo ControlError
    
    lsFecRemate = Format$(Me.txtFecRemate, "mm/dd/yyyy")
    ' Obtiene la cuenta de Ahorros ****
    Dim loDatRem As COMNColoCPig.NCOMColPRecGar
    Set loDatRem = New COMNColoCPig.NCOMColPRecGar
        lsCtaSobranteRemate = loDatRem.nObtieneCtaSobranteRemate(fsRemateCadaAgencia, lsmensaje)
        If Trim(lsmensaje) <> "" Then
             MsgBox lsmensaje, vbInformation, "Aviso"
             Exit Sub
        End If
    Set loDatRem = Nothing
        
    If lsCtaSobranteRemate = "" Then
         MsgBox "No se encuentra configurada la Cta de Ahorros de Sobrante", vbInformation, "Aviso"
         Exit Sub
    End If
    
    ' Obtiene el Monto del sobrante
    Set loSobranteR = New COMNColoCPig.NCOMColPRecGar
        lnMontoSobrante = loSobranteR.nObtieneSobranteRemateProceso(Trim(txtNumRemate.Text), fsRemateCadaAgencia, lsmensaje)
    Set loSobranteR = Nothing
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    lnMontoSobrante = Format(lnMontoSobrante, "#0.00")
    
    If MsgBox("Esta seguro de Finalizar el Proceso de Remate ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then
        '********** Finaliza el Remate
        Set loFinRem = New COMNColoCPig.NCOMColPRecGar
            Call loFinRem.nFinalizaRemate(Trim(txtNumRemate.Text), lsFecRemate, lsCtaSobranteRemate, _
                        lnMontoSobrante, fsRemateCadaAgencia, gsCodUser, sLpt)
        
        
    
       Set loFinRem = New COMNColoCPig.NCOMColPRecGar
       Call loFinRem.nRecGarGrabaDatosPreparaCredPignoraticio("R", txtNumRemate.Text, gColPRecGarEstTerminado, _
                Format(txtFecRemate.Text, "mm/dd/yyyy hh:mm"), fsRemateCadaAgencia, , , , , , , , False, _
                fnTasaCustodiaVencida, fnTasaPreparacionRemate, fnTasaImpuesto, fnFactorPrecioBaseRemate, True)
      Set loFinRem = Nothing
        
        
        txtEstado = "FINALIZADO"
        cmdIniciaRemate.Enabled = False
        cmdPlanillaContEntranRemate.Enabled = False
        cmdRegistraVentaRemate.Enabled = False
        'If pBandEOARem Then cmdRegVenEOARem.Enabled = False
        cmdFinalizaRemate.Enabled = False
        cmdPlanillaVentasRemate.Enabled = True
        cmdPlanillaNoVendido.Enabled = True
        cmdActaRemate.Enabled = True
        cmdCartaSobrante.Enabled = True
        CmdActa.Enabled = True
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub


' Inicializa el Proceso de Remate en la CMAC - T
Private Sub cmdIniciaRemate_Click()

Dim loIniRem As COMNColoCPig.NCOMColPRecGar
Dim lsFecRemate As String
Dim X As Integer

On Error GoTo ControlError
    'Selec Age. a realizar Remate
    'frmSelectAgencias.Inicio Me
    'frmSelectAgencias.Show 1
    If Not IsDate(txtFecRemate) Then
        MsgBox "Ingrese una Fecha correcta", vbInformation, "Aviso"
        Exit Sub
    End If

    
    lsFecRemate = Format$(Me.txtFecRemate, "mm/dd/yyyy")
    
    If MsgBox("Esta seguro de Iniciar Remate ? ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then

        'Remate debe figurar con la fecha anterior
        'vFecRemNue = DateAdd("d", -pDifeDiasRema, txtFecRemate.Text)
        vFecRemNue = txtFecRemate.Text
        
        vFecAviso = DateAdd("d", -fnDiasVctoParaRemate, vFecRemNue)
        
        'For X = 1 To frmSelectAgencias.List1.ListCount
        '    If frmSelectAgencias.List1.Selected(X - 1) = True Then
                '********** Realiza el Inicio
                Set loIniRem = New COMNColoCPig.NCOMColPRecGar
                    Call loIniRem.nIniciaRemate(Trim(txtNumRemate.Text), lsFecRemate, fsRemateCadaAgencia, fnDiasVctoParaRemate, fnFactorPrecioBaseRemate, Val(Me.txtPreOro14.Text), Val(Me.txtPreOro16.Text), Val(Me.txtPreOro18.Text), Val(Me.txtPreOro21.Text), fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate)
                Set loIniRem = Nothing
                
        '    End If
        'Next X
        
        txtEstado = "INICIADO"
        cmdIniciaRemate.Enabled = False
        cmdPlanillaContEntranRemate.Enabled = True
        cmdRegistraVentaRemate.Enabled = True
        cmdFinalizaRemate.Enabled = True
        CmdActa.Enabled = True
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdListAnt_Click()

On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    lsCadImprimir = ""
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaParaRemateConSiaf(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasVctoParaRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, _
                        CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, lsmensaje, gImpresora)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

' Genera la planilla de Contratos que entran al remate
Private Sub cmdPlanillaContEntranRemate_Click()

On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    If Not IsDate(txtFecRemate) Then
        MsgBox "Ingrese una Fecha correcta", vbInformation, "Aviso"
        Exit Sub
    End If

    lsCadImprimir = ""
    Dim CCONTHOJAS As Integer
    
     If optImpresion(2).value = True Then
    
            For lnAge = 1 To frmSelectAgencias.List1.ListCount
                If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                             Call ExportarExcel(1)
                             Exit Sub
                End If
            Next lnAge
      
    
    End If
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then

            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaParaRemate(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasVctoParaRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, _
                        CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, lsmensaje, gImpresora)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Exit Sub
            End If

            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub ExportarExcel(ByVal nTpoReporte As Integer)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim nFila As Long, i As Long, lnAge As Integer


Dim oRep As COMNColoCPig.NCOMColPRecGar, rsTemp As ADODB.Recordset, lrCredPigJoyasDet As ADODB.Recordset
Dim lsRep As String, loMuestraContrato As DColPContrato
Dim psAge As String, psagenom As String


If MsgBox("Este reporte puede demorar unos minutos..." & vbCrLf & "¿Desea procesar la información ?", vbOKOnly + vbQuestion, "AVISO") = vbNo Then
    Exit Sub
End If

Set oRep = New COMNColoCPig.NCOMColPRecGar
'(lsRep, Me.TxtFecha, Me.txtFechaF, Me.txtMonto.value, Me.txtMontoF.value, gsNomAge, gsNomCmac, gdFecSis, Me.TxtAgencia.Text, TxtBuscarUser.Text, rtfCartas.Text, Val(EditMoney3.Text), lsEstadosCheques, lsOptionsCheques, lsOrden, lscheck, lscmacllamada, lscmacrecepcion, pspersoneria)
'  lscadena = lscadena & GetRepCapNMejoresClientes(CLng(pnMontoIni), pnTipoCambio, psCodAge, psEmpresa, psNomAge, pdFecSis)
Select Case nTpoReporte
       Case 2
        For lnAge = 1 To frmSelectAgencias.List1.ListCount

           If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then

             psAge = Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2)
             psagenom = Trim(Mid(frmSelectAgencias.List1.List(lnAge - 1), 3))
              Set rsTemp = oRep.nListadoParaRemConSIAF(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                           Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, 30, gdFecSis, _
                           fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), _
                           CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, IIf(fnJoyasDet = 1, True, False))


                GoTo Continua
           End If
       Next

End Select

Continua:

If rsTemp.EOF Then
    MsgBox "No se encontro información para este reporte", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
End If

Dim lsArchivoN As String, lbLibroOpen As Boolean

Dim Item As Integer

   lsArchivoN = App.path & "\Spooler\RemDet" & lsRep & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"

   OleExcel.Class = "ExcelWorkSheet"
   lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
   If lbLibroOpen Then
            Set xlHoja1 = xlLibro.Worksheets(1)

            ExcelAddHoja Format(gdFecSis, "yyyymmdd") & psagenom, xlLibro, xlHoja1

            nFila = 1

            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
            xlHoja1.Range("F2:H2").MergeCells = True
            xlHoja1.Cells(nFila, 6) = Format(gdFecSis, "Long Date")

             'prgBar.value = 2


            nFila = 3
            xlHoja1.Cells(nFila, 1) = "LISTADO DE CONTRATOS PARA REMATE NRO " & IIf(Me.txtEstado = "NO INICIADO", "0000", Me.txtNumRemate) & "DEL " & Format(Me.txtFecRemate, "dd/mm/yyyy") & " -  " & psagenom


            xlHoja1.Range("A1:M5").Font.Bold = True

            xlHoja1.Range("A3:M3").MergeCells = True
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
            xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter

            'xlHoja1.Range("A5:H5").AutoFilter

            nFila = 5

            xlHoja1.Cells(nFila, 1) = "ITEM "
            xlHoja1.Cells(nFila, 2) = "CUENTA SIAFC   "
            xlHoja1.Cells(nFila, 3) = "CUENTA SICMACI "
            xlHoja1.Cells(nFila, 4) = "CLIENTE "
            xlHoja1.Cells(nFila, 5) = "PIEZAS "
            xlHoja1.Cells(nFila, 6) = "DESCRIPCION "
            xlHoja1.Cells(nFila, 7) = "               "
            xlHoja1.Cells(nFila, 8) = "               "
            xlHoja1.Cells(nFila, 9) = "14K "
            xlHoja1.Cells(nFila, 10) = "16K "
            xlHoja1.Cells(nFila, 11) = "18K "
            xlHoja1.Cells(nFila, 12) = "21K "
            xlHoja1.Cells(nFila, 13) = "VAL REGISTRO "




            i = 0
            While Not rsTemp.EOF
                nFila = nFila + 1

                'prgBar.value = ((i) / RSTEMP.RecordCount) * 100

                i = i + 1


                xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                'xlHoja1.Cells(nFila, 2) = RSTEMP!CnOMcLIENTE
                xlHoja1.Cells(nFila, 2) = rsTemp!codesiaf
                xlHoja1.Cells(nFila, 3) = rsTemp!cCtaCod
                xlHoja1.Cells(nFila, 4) = rsTemp!cNomCliente
                xlHoja1.Cells(nFila, 5) = rsTemp!npiezas
                
                xlHoja1.Cells(nFila, 9) = Format(rsTemp!nK14, "#0.00")
                xlHoja1.Cells(nFila, 10) = Format(rsTemp!nK16, "#0.00")
                xlHoja1.Cells(nFila, 11) = Format(rsTemp!nK18, "#0.00")
                xlHoja1.Cells(nFila, 12) = Format(rsTemp!nK21, "#0.00")
                xlHoja1.Cells(nFila, 13) = Format(rsTemp!nRemSubBaseVta, "#,##0.00")


                Set loMuestraContrato = New DColPContrato
                        Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(rsTemp!cCtaCod)
                    Set loMuestraContrato = Nothing

                    'Item = 0
                    Do While Not lrCredPigJoyasDet.EOF

                        nFila = nFila + 1

                        'cKilataje,nItem, nPiezas, nPesoBruto, nPesoNeto, nValTasac, cDescrip

                        'Item = Item + 1
                        xlHoja1.Cells(nFila, 5) = Format(lrCredPigJoyasDet!nItem, "00")
                        xlHoja1.Cells(nFila, 6) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0)
                        xlHoja1.Cells(nFila, 7) = ImpreFormat(lrCredPigJoyasDet!cdescrip, 27)
                        xlHoja1.Cells(nFila, 8) = lrCredPigJoyasDet!ckilataje & "K "
                        xlHoja1.Cells(nFila, 9) = ImpreFormat(lrCredPigJoyasDet!npesobruto, 4, 2) & "Gr"
                        xlHoja1.Cells(nFila, 10) = ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"




                        'lmDetalle(Item) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0) & " " & ImpreFormat(lrCredPigJoyasDet!cDescrip, 27) & " " _
                                    & lrCredPigJoyasDet!cKilataje & "K " & ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"


                        lrCredPigJoyasDet.MoveNext
                    Loop



                rsTemp.MoveNext

            Wend

           ' xlHoja1.Columns.AutoFit

            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit



            'Cierro...
            OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
            OleExcel.SourceDoc = lsArchivoN
            OleExcel.Verb = 1
            OleExcel.Action = 1
            OleExcel.DoVerb -1

           ' prgBar.value = 100

   End If

   Set rsTemp = Nothing


    
End Sub


'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Private Sub cmdPlanillaNoVendido_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    lsCadImprimir = ""
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaNoVendidosRemate(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasVctoParaRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, _
                        CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
                
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdPlanillaVentasRemate_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    lsCadImprimir = ""
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaVentaRemate(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasVctoParaRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, _
                        CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, lsmensaje, gImpresora)
                 If Trim(lsmensaje) <> "" Then
                     MsgBox lsmensaje, vbInformation, "Aviso"
                     Exit Sub
                End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub



'Llama al formulario para registrar las Ventas del remate
Private Sub cmdRegistraVentaRemate_Click()
    frmColPRemateRegVenta.Inicio (txtNumRemate.Text)
End Sub


'Inicializa el formulario actual y carga los datos del último remate
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaParametros
    Call CargaDatosUltimoRemate
End Sub

Private Sub CargaDatosUltimoRemate()

Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrdatosrem As ADODB.Recordset
Dim lsUltRemate As String
Dim lsmensaje As String

txtPreOro14 = Format(0, "#0.00")
txtPreOro16 = Format(0, "#0.00")
txtPreOro18 = Format(0, "#0.00")
txtPreOro21 = Format(0, "#0.00")

Set lrdatosrem = New ADODB.Recordset
Set loDatos = New COMNColoCPig.NCOMColPRecGar
    lsUltRemate = loDatos.nObtieneNroUltimoProceso("R", fsRemateCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    Set lrdatosrem = loDatos.nObtieneDatosProcesoRGCredPig("R", lsUltRemate, fsRemateCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    If lrdatosrem Is Nothing Then
        Exit Sub
    End If
    If lrdatosrem!nRGEstado = gColPRecGarEstNoIniciado And Format(lrdatosrem!dProceso, "dd/mm/yyyy") <> Format(gdFecSis, "dd/mm/yyyy") Then
        Set lrdatosrem = Nothing
        Set lrdatosrem = loDatos.nObtieneDatosProcesoRGCredPig("R", FillNum(Trim(Str(Val(lsUltRemate) - 1)), 4, "0"), fsRemateCadaAgencia, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
Set loDatos = Nothing
'Mostrar Datos
If lrdatosrem Is Nothing Then Exit Sub

txtNumRemate = lrdatosrem!cNroProceso
txtFecRemate = Format(lrdatosrem!dProceso, "dd/mm/yyyy")
txtHorRemate = Format(lrdatosrem!dProceso, "hh:mm")
txtPreOro14 = Format(lrdatosrem!nPrecioK14, "#0.00")
txtPreOro16 = Format(lrdatosrem!nPrecioK16, "#0.00")
txtPreOro18 = Format(lrdatosrem!nPrecioK18, "#0.00")
txtPreOro21 = Format(lrdatosrem!nPrecioK21, "#0.00")

If lrdatosrem!nRGEstado = gColPRecGarEstNoIniciado Then
    txtEstado = "NO INICIADO"
    If Val(txtPreOro14) = 0 And Val(txtPreOro16) = 0 And Val(txtPreOro18) = 0 And Val(txtPreOro21) = 0 Then
        cmdIniciaRemate.Enabled = False
        MsgBox " Regrese a Preparacion de Remate a ingresar precio del Oro ", vbInformation, " Aviso "
    ElseIf Format(lrdatosrem!dProceso, "dd/mm/yyyy") <> Format(gdFecSis, "dd/mm/yyyy") Then
        cmdIniciaRemate.Enabled = False
        MsgBox " No es la fecha para el Remate ", vbInformation, " Aviso "
    End If
ElseIf lrdatosrem!nRGEstado = gColPRecGarEstIniciado Then
    txtEstado = "INICIADO"
    cmdIniciaRemate.Enabled = False
    cmdPlanillaContEntranRemate.Enabled = True
    cmdRegistraVentaRemate.Enabled = True
    'cmdRegVenEOARem.Enabled = True
    cmdFinalizaRemate.Enabled = True
    cmdActaRemate.Enabled = True
    CmdActa.Enabled = True
ElseIf lrdatosrem!nRGEstado = gColPRecGarEstTerminado Then
    txtEstado = "FINALIZADO"
    cmdIniciaRemate.Enabled = False
    cmdPlanillaContEntranRemate.Enabled = False
    cmdRegistraVentaRemate.Enabled = False
    
    cmdFinalizaRemate.Enabled = False
    cmdPlanillaVentasRemate.Enabled = True
    cmdPlanillaNoVendido.Enabled = True
    cmdActaRemate.Enabled = True
    cmdCartaSobrante.Enabled = True
    cmdRemate.Visible = False
    CmdActa.Enabled = True
End If
Set lrdatosrem = Nothing

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
'Valida el campo txtfecremate
Private Sub txtFecRemate_GotFocus()
fEnfoque txtFecRemate
End Sub

Private Sub txtFecRemate_LostFocus()
If Not ValFecha(txtFecRemate) Then
    txtFecRemate.SetFocus
End If
End Sub

'Valida el campo txthorremate
Private Sub txtHorRemate_GotFocus()
fEnfoque txtHorRemate
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loParam = New COMDColocPig.DCOMColPCalculos

    fnDiasVctoParaRemate = loParam.dObtieneColocParametro(gConsColPDiasAtrasoParaRemate)
    fnFactorPrecioBaseRemate = loParam.dObtieneColocParametro(gConsColPFactorPrecioBaseRemate)
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    
    
Set loParam = Nothing
    pPrevioMax = 5000
    pLineasMax = 56
    pHojaFiMax = 66
    'pAgeRemSub = Right(ReadVarSis("CPR", "cAgeRemSub"), 2)
    'pVerCodAnt = IIf(Left(ReadVarSis("CPR", "cVerCodAnt"), 1) = "S", True, False)
    'pCtaAhoSob = ReadVarSis("AHO", "cCtaSobPre")
    'pDifeDiasRema = Val(ReadVarSis("CPR", "nDifeDiasRema"))
    
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    fnJoyasDet = loConstSis.LeeConstSistema(109)
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsRemateCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsRemateCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing
End Sub


Private Sub ImprimeActa()
    Dim loPrevio As previo.clsPrevio
    Dim lsCadImprimir  As String
    Dim lsCartaModelo As String
  
    lsCadImprimir = ""
    rtfCartas.FileName = App.path & cPlantillaActaRema
     
    lsCartaModelo = rtfCartas.Text
    lsCartaModelo = Replace(lsCartaModelo, "<<NROREMATE>>", txtNumRemate.Text, , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<FECHAC>>", Format(gdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<FECHAL>>", Format(gdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<MARTILLERO>>", "ANGEL DEL CASTILLO", , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<GERENTE>>", "CARLOS CASTRO CASTRO", , 2, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<ADMINISTRADOR>>", "JUVENAL VARGAS TRUJILLO", , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<VEEDOR>>", "MERCEDES SALAZAR", , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<AGENCIA>>", gsNomAge, , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<AGENCIA>>", gsNomAge, , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<HORA>>", Format(Now(), "HH:MM:SS"), , 1, vbTextCompare)
    lsCartaModelo = Replace(lsCartaModelo, "<<NROREMATE>>", txtNumRemate.Text, , 1, vbTextCompare)
    lsCadImprimir = lsCadImprimir & lsCartaModelo
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Sobrante de Remate", True
    Set loPrevio = Nothing
    
   
End Sub

