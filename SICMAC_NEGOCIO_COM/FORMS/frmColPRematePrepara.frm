VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPRematePrepara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Preparación de Remate"
   ClientHeight    =   5145
   ClientLeft      =   1065
   ClientTop       =   2205
   ClientWidth     =   6900
   Icon            =   "frmColPRematePrepara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   4950
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   6675
      Begin VB.Frame fraContenedor 
         Height          =   2160
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   6360
         Begin VB.CommandButton cmdListAnt 
            Caption         =   "Listado de Contratos para Remate con SIAF"
            Height          =   345
            Left            =   255
            TabIndex        =   38
            Top             =   2370
            Visible         =   0   'False
            Width           =   3645
         End
         Begin VB.CommandButton cmdAntiguos 
            Caption         =   "Planilla para remate con SIAF"
            Height          =   345
            Left            =   240
            TabIndex        =   37
            Top             =   2160
            Visible         =   0   'False
            Width           =   3660
         End
         Begin VB.CheckBox chkCobrarGasto 
            Caption         =   "Cobrar Gasto"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4920
            TabIndex        =   36
            Top             =   1800
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdAgencia 
            Caption         =   "A&gencias..."
            Height          =   345
            Left            =   4920
            TabIndex        =   20
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Frame fraImpresion 
            Caption         =   "Impresión"
            Height          =   1290
            Left            =   4890
            TabIndex        =   35
            Top             =   120
            Width           =   1320
            Begin VB.OptionButton optImpresion 
               Caption         =   "Excel"
               Height          =   270
               Index           =   3
               Left            =   165
               TabIndex        =   39
               Top             =   975
               Width           =   990
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Pantalla"
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   17
               Top             =   285
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Impresora"
               Height          =   270
               Index           =   1
               Left            =   165
               TabIndex        =   18
               Top             =   480
               Width           =   990
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Archivo"
               Height          =   270
               Index           =   2
               Left            =   165
               TabIndex        =   19
               Top             =   720
               Width           =   990
            End
         End
         Begin VB.CommandButton cmdImpListRema 
            Caption         =   "Listado de Contratos para Remate"
            Enabled         =   0   'False
            Height          =   360
            Left            =   240
            TabIndex        =   16
            Top             =   1540
            Width           =   3645
         End
         Begin VB.CommandButton cmdImpPlanRema 
            Caption         =   "Planilla para remate"
            Enabled         =   0   'False
            Height          =   360
            Left            =   255
            TabIndex        =   15
            Top             =   1110
            Width           =   3645
         End
         Begin VB.CommandButton cmdImpAvisVenc 
            Caption         =   "Cartas de Aviso de Vencimiento"
            Height          =   360
            Left            =   255
            TabIndex        =   13
            Top             =   255
            Width           =   3645
         End
         Begin VB.CommandButton cmdImpAvisRema 
            Caption         =   "Cartas de Aviso de Remate"
            Height          =   360
            Left            =   255
            TabIndex        =   14
            Top             =   675
            Width           =   3645
         End
         Begin VB.Label lblCargo 
            Caption         =   "Cargo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3960
            TabIndex        =   43
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblNombre 
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   3960
            TabIndex        =   42
            Top             =   960
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image ImgFirma 
            Height          =   615
            Left            =   3960
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   390
         Left            =   3120
         TabIndex        =   9
         Top             =   1800
         Width           =   990
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
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   300
         Width           =   1305
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
         Left            =   765
         TabIndex        =   0
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   5280
         TabIndex        =   11
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   4200
         TabIndex        =   10
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   5400
         TabIndex        =   12
         Top             =   4440
         Width           =   990
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Precios del Oro  (Sin  IGV) "
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
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   6375
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
            Left            =   5400
            MaxLength       =   6
            TabIndex        =   8
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
            Left            =   3795
            MaxLength       =   6
            TabIndex        =   7
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
            Left            =   2280
            MaxLength       =   6
            TabIndex        =   6
            Top             =   225
            Width           =   750
         End
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
            Left            =   675
            MaxLength       =   6
            TabIndex        =   5
            Top             =   225
            Width           =   750
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "14 K :"
            Height          =   225
            Index           =   1
            Left            =   165
            TabIndex        =   26
            Top             =   255
            Width           =   540
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "16 K :"
            Height          =   225
            Index           =   2
            Left            =   1755
            TabIndex        =   25
            Top             =   240
            Width           =   510
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "18 K :"
            Height          =   225
            Index           =   3
            Left            =   3270
            TabIndex        =   24
            Top             =   240
            Width           =   510
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "21 K :"
            Height          =   225
            Index           =   4
            Left            =   4920
            TabIndex        =   23
            Top             =   240
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox txtFecRemate 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   300
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
         Left            =   3840
         TabIndex        =   2
         Top             =   300
         Width           =   705
         _ExtentX        =   1244
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
      Begin RichTextLib.RichTextBox rtfImp 
         Height          =   360
         Left            =   0
         TabIndex        =   32
         Top             =   2880
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmColPRematePrepara.frx":030A
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
         Height          =   360
         Left            =   0
         TabIndex        =   33
         Top             =   3600
         Visible         =   0   'False
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   635
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmColPRematePrepara.frx":038A
      End
      Begin MSComctlLib.ProgressBar prgList 
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   4440
         Visible         =   0   'False
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSMask.MaskEdBox TxtFechaCorte 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   720
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
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha de Corte de Remate:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   750
         Width           =   3255
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Estado"
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   31
         Top             =   345
         Width           =   510
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Número"
         Height          =   255
         Index           =   7
         Left            =   105
         TabIndex        =   30
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Hora"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   29
         Top             =   330
         Width           =   405
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   6
         Left            =   1485
         TabIndex        =   28
         Top             =   330
         Width           =   630
      End
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   240
      Top             =   4650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   870
      Left            =   210
      OleObjectBlob   =   "frmColPRematePrepara.frx":040D
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmColPRematePrepara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************
'* REMATE - PREPARACION.
'Archivo:  frmColPRematePrepara.frm
'LAYG   :  01/07/2001.
'Resumen:  Preparacion para Remate de Creditos Pignoraticios
'          Registra precios de oro // Genera Cartas de Aviso Vencimiento // Aviso Remate

Option Explicit
'Variables el formulario
Dim pCorte As Variant
Dim fnPrevioMax As Double
Dim fnLineasMax As Long
Dim fnHojaFiMax As Integer
Dim fbVerCodAnt As Boolean


Dim fnDiasAtrasoAvisoVencimiento As Double
Dim fnDiasAtrasoAvisoRemate As Double

Dim fnFactorPrecioBaseRemate As Double
Dim fnTasaPreparacionRemate As Double
Dim fnTasaImpuesto As Double
Dim fnTasaCustodiaVencida As Double
Dim fsRemateCadaAgencia As String
Dim fnJoyasDet As Integer
Dim fnCostoCorrespondencia As Double

Dim RegCredPrend As New adodb.Recordset
Dim RegJoyas As New adodb.Recordset
Dim RegProcesar As New adodb.Recordset
Dim sSQL As String
Dim MuestraImpresion As Boolean
Dim vRTFImp As String, vBuffer As String
Dim vCont As Double
Dim vNomAge As String

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub HabilitaControles(ByVal pbEditar As Boolean, ByVal pbGrabar As Boolean, ByVal pbSalir As Boolean, _
    ByVal pbCancelar As Boolean, ByVal pbFecRemate As Boolean, ByVal pbHorRemate As Boolean, _
    ByVal pbPreOro14 As Boolean, ByVal pbPreOro16 As Boolean, ByVal pbPreOro18 As Boolean, ByVal pbPreOro21 As Boolean, _
    ByVal pbImpAvisRema As Boolean, ByVal pbImpAvisVenc As Boolean, ByVal pbImpPlanRema As Boolean, ByVal pbImpListRema As Boolean)

    cmdEditar.Enabled = pbEditar
    cmdGrabar.Enabled = pbGrabar
    cmdSalir.Enabled = pbSalir
    cmdCancelar.Enabled = pbCancelar
    txtFecRemate.Enabled = pbFecRemate
    TxtFechaCorte.Enabled = pbFecRemate
    txtHorRemate.Enabled = pbHorRemate
    txtPreOro14.Enabled = pbPreOro14
    txtPreOro16.Enabled = pbPreOro16
    txtPreOro18.Enabled = pbPreOro18
    txtPreOro21.Enabled = pbPreOro21
    cmdImpAvisRema.Enabled = pbImpAvisRema
    cmdImpAvisVenc.Enabled = pbImpAvisVenc
    cmdImpPlanRema.Enabled = pbImpPlanRema
    cmdImpListRema.Enabled = pbImpListRema
    
End Sub

Private Sub cmdAntiguos_Click()
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
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasAtrasoAvisoRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), _
                        CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, lsmensaje, gImpresora)
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
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, lsCadImprimir
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If
End Sub

'Permite no reconocer los datos ingresados
Private Sub cmdCancelar_Click()
On Error GoTo ControlError
 Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, True, False, False)
 Limpiar
 VeriDatRem
 If txtPreOro14 > 0 And txtPreOro16 > 0 And txtPreOro18 > 0 _
        And txtPreOro21 > 0 Then
    cmdImpListRema.Enabled = True
    cmdImpPlanRema.Enabled = True
 End If
 Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite editar los campos editables de un remate
Private Sub CmdEditar_Click()
On Error GoTo ControlError
Call HabilitaControles(False, True, False, True, True, True, True, True, True, True, False, False, False, False)

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'permite grabar los cambios ingresados
Private Sub cmdGrabar_Click()
On Error GoTo ControlError
Dim loGrabRem As COMNColoCPig.NCOMColPRecGar

Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, True, True, True)

If txtPreOro14 > 0 And txtPreOro16 > 0 And txtPreOro18 > 0 _
  And txtPreOro21 > 0 Then
    cmdImpListRema.Enabled = True
    cmdImpPlanRema.Enabled = True
End If

Set loGrabRem = New COMNColoCPig.NCOMColPRecGar
    Call loGrabRem.nRecGarGrabaDatosPreparaCredPignoraticio("R", txtNumRemate.Text, gColPRecGarEstNoIniciado, _
                Format(Me.txtFecRemate.Text, "mm/dd/yyyy hh:mm"), fsRemateCadaAgencia, Val(Me.txtPreOro14.Text), Val(Me.txtPreOro16.Text), Val(Me.txtPreOro18.Text), Val(Me.txtPreOro21.Text), , , , False, fnTasaCustodiaVencida, fnTasaPreparacionRemate, fnTasaImpuesto, fnFactorPrecioBaseRemate, , Format(Me.TxtFechaCorte.Text, "dd/mm/yyyy"))
Set loGrabRem = Nothing
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
        'ImprimeContRema
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeListadoParaRemateConSIAF(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasAtrasoAvisoRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), _
                        CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, IIf(fnJoyasDet = 1, True, False), lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                            MsgBox lsmensaje, vbInformation, "Aviso"
                            Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Listado Contratos para Remate", True
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, vBuffer
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If

End Sub

'Permite salir del formulario actual
Private Sub cmdSalir_Click()
    Unload frmSelectAgencias
    Unload Me
End Sub

Private Sub Command1_Click()
   
Dim RStream As adodb.Stream
Dim oDCred As COMDCredito.DCOMCredDoc
Dim pR As adodb.Recordset
Dim sRutaTmp As String
Dim sRutaImg As String

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set pR = oDCred.RecuperaJefeAgencia(gsCodAge)
    Set oDCred = Nothing
    
    Me.lblNombre = pR!cPersNombre
    Me.lblCargo = pR!Cargo
    
    Set RStream = New adodb.Stream
    RStream.Type = adTypeBinary
    RStream.Open
    
    RStream.Write pR.Fields("Firma").value

    sRutaTmp = App.path & "\spooler\FirJefe.bmp"
    RStream.SaveToFile App.path & "\spooler\FirJefe.bmp", adSaveCreateOverWrite
    sRutaImg = sRutaTmp
    If Len(Trim(sRutaImg)) > 0 Then
        ImgFirma.Picture = LoadPicture(sRutaImg)
    End If
    Kill (App.path & "\spooler\FirJefe.bmp")
    RStream.Close
    Set RStream = Nothing
'    CargarFirmadeBD = True
End Sub

'Permite inicializar el formulario actual
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaParametros
    Limpiar
End Sub

Private Sub Limpiar()

Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrdatosrem As adodb.Recordset
Dim lsUltRemate As String
Dim lsmensaje As String

txtPreOro14 = Format(0, "#0.00")
txtPreOro16 = Format(0, "#0.00")
txtPreOro18 = Format(0, "#0.00")
txtPreOro21 = Format(0, "#0.00")

Set lrdatosrem = New adodb.Recordset
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
Set loDatos = Nothing
'Mostrar Datos
If (lrdatosrem Is Nothing) Then
    Exit Sub
End If
txtNumRemate = lrdatosrem!cNroProceso
txtFecRemate = Format(lrdatosrem!dProceso, "dd/mm/yyyy")
TxtFechaCorte = Format(lrdatosrem!dProceso, "dd/mm/yyyy")
txtHorRemate = Format(lrdatosrem!dProceso, "hh:mm")
txtPreOro14 = Format(lrdatosrem!nPrecioK14, "#0.00")
txtPreOro16 = Format(lrdatosrem!nPrecioK16, "#0.00")
txtPreOro18 = Format(lrdatosrem!nPrecioK18, "#0.00")
txtPreOro21 = Format(lrdatosrem!nPrecioK21, "#0.00")

If lrdatosrem!nRGEstado = 0 Then
    txtEstado = "NO INICIADO"
    If Val(txtPreOro18) > 0 And Val(txtPreOro21) > 0 Then 'MPBR 2004/09/23
        cmdImpListRema.Enabled = True
        cmdImpPlanRema.Enabled = True
    End If
ElseIf lrdatosrem!nRGEstado = 1 Then
    txtEstado = "INICIADO"
    cmdEditar.Enabled = False
Else
    MsgBox " No existe el remate generado", vbCritical, " Error de Sistema "
    cmdEditar.Enabled = False
    txtNumRemate = ""
    txtFecRemate = Format("01/01/2000", "dd/mm/yyyy")
    TxtFechaCorte = Format("01/01/2000", "dd/mm/yyyy")
    txtHorRemate = Format("00:00", "hh:mm")
End If
Set lrdatosrem = Nothing

End Sub


Private Sub fraContenedor_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

End Sub

Private Sub TxtFechaCorte_Change()
    fEnfoque TxtFechaCorte
End Sub


Private Sub TxtFechaCorte_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtPreOro14.SetFocus
  End If
End Sub

'Valida el campo txtfecremate
Private Sub txtFecRemate_GotFocus()
    fEnfoque txtFecRemate
End Sub
Private Sub txtFecRemate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtHorRemate.SetFocus
End If
End Sub
Private Sub txtFecRemate_LostFocus()
If Not ValFecha(txtFecRemate) Then
    txtFecRemate.SetFocus
ElseIf DateDiff("d", txtFecRemate, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    'txtFecRemate.Enabled = True
    txtFecRemate.SetFocus
End If
'VeriDatRem
End Sub
Private Sub txtFecRemate_Validate(Cancel As Boolean)
If Not ValFecha(txtFecRemate) Then
    Cancel = True
ElseIf DateDiff("d", txtFecRemate, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    Cancel = True
End If
End Sub

'Valida el campo txthorremate
Private Sub txtHorRemate_GotFocus()
fEnfoque txtHorRemate
End Sub
Private Sub txtHorRemate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtFechaCorte.SetFocus
End If
End Sub
Private Sub txtHorRemate_LostFocus()
If Not ValidaHora(txtHorRemate) Then
    txtHorRemate.SetFocus
End If
'VeriDatRem
End Sub

'Valida el campo txtpreoro14
Private Sub txtPreOro14_GotFocus()
fEnfoque txtPreOro14
End Sub
Private Sub txtPreOro14_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro14, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro14 = Format(txtPreOro14, "#0.00")
    txtPreOro16.SetFocus
End If
End Sub
Private Sub txtPreOro14_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro16
Private Sub txtPreOro16_GotFocus()
fEnfoque txtPreOro16
End Sub
Private Sub txtPreOro16_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro16, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro16 = Format(txtPreOro16, "#0.00")
    txtPreOro18.SetFocus
End If
End Sub
Private Sub txtPreOro16_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro18
Private Sub txtPreOro18_GotFocus()
fEnfoque txtPreOro18
End Sub
Private Sub txtPreOro18_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro18, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro18 = Format(txtPreOro18, "#0.00")
    txtPreOro21.SetFocus
End If
End Sub
Private Sub txtPreOro18_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro21
Private Sub txtPreOro21_GotFocus()
fEnfoque txtPreOro21
End Sub
Private Sub txtPreOro21_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro21, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro21 = Format(txtPreOro21, "#0.00")
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtPreOro21_LostFocus()
    VeriPreOro
End Sub

' Permite activar la opción de procesar solo cuando están ingresados los campos
' fecha y hora de remate
Private Sub VeriDatRem()
    If Len(txtNumRemate) > 0 And Len(txtFecRemate) = 10 And Len(txtHorRemate) = 5 Then
    End If
End Sub

' Permite activar la opción de grabar solo cuando están ingresados los campos
' del precio del oro
Private Sub VeriPreOro()
    If Val(txtPreOro14) > 0 And Val(txtPreOro16) > 0 And Val(txtPreOro18) > 0 _
        And Val(txtPreOro21) > 0 Then
        cmdGrabar.Enabled = True
    End If
End Sub

'Permite imprimir en pantalla o directo a la impresora los avisos de vencimiento
Private Sub cmdImpAvisVenc_Click()
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
    rtfCartas.FileName = App.path & cPlantillaAvisoVencimiento
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nRemImprimeAvisoVencimiento(rtfCartas.Text, Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), fnDiasAtrasoAvisoVencimiento, _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, gdFecSis, _
                        IIf(chkCobrarGasto.value = 1, fnCostoCorrespondencia, 0), gsCodAge, gsCodUser, lsmensaje, gImpresora)
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
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", False
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, lsCadImprimir
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If
End Sub

'Permite imprimir en pantalla o directo a la impresora los avisos de remate
Private Sub cmdImpAvisRema_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String

'Dim loPrevio As previo.clsPrevio
Dim dFechaCorte As Date

Dim lnAge As Integer
    
    If Not IsDate(txtFecRemate) Then
        MsgBox "Ingrese una Fecha correcta", vbInformation, "Aviso"
        Exit Sub
    End If
    dFechaCorte = CDate(TxtFechaCorte.Text)
    
    '*** PEAC 20080305
    CargaFirmaJefeAge

    'Pig_CartasNotariales dFechaCorte
    Pig_CartasNotariales dFechaCorte, Me.lblNombre, Me.lblCargo, Me.ImgFirma.Picture
    

    '*** end PEAC

'    lsCadImprimir = ""
'    rtfCartas.FileName = App.path & cPlantillaAvisoRemate
'
'    For lnAge = 1 To frmSelectAgencias.List1.ListCount
'        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
'
'            Set loImprime = New COMNColoCPig.NCOMColPRecGar
'                lsCadImprimir = lsCadImprimir & loImprime.nRemImprimeAvisoRemate(rtfCartas.Text, Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), fnDiasAtrasoAvisoRemate, _
'                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, gdFecSis, _
'                        IIf(chkCobrarGasto.value = 1, fnCostoCorrespondencia, 0), gsCodAge, gsCodUser, lsmensaje, gImpresora)
'                If Trim(lsmensaje) <> "" Then
'                    MsgBox lsmensaje, vbInformation, "Aviso"
'                    Exit Sub
'                End If
'            Set loImprime = Nothing
'
'        End If
'    Next lnAge
'
'    If Len(Trim(lsCadImprimir)) = 0 Then
'        MsgBox "No hay datos para mostrar en el reporte", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    If Me.optImpresion(0).value = True Then
'        Set loPrevio = New previo.clsPrevio
'            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", False
'        Set loPrevio = Nothing
'    Else
'        dlgGrabar.CancelError = True
'        dlgGrabar.InitDir = App.path
'        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'        dlgGrabar.ShowSave
'        If dlgGrabar.FileName <> "" Then
'           Open dlgGrabar.FileName For Output As #1
'            Print #1, lsCadImprimir
'            Close #1
'        End If
'    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If


End Sub


'Permite imprimir en pantalla o directo a la impresora la planilla de remate
Private Sub cmdImpPlanRema_Click()
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
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaParaRemate(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasAtrasoAvisoRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), _
                        CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, lsmensaje, gImpresora)
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
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, lsCadImprimir
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If
End Sub

'Permite imprimir en pantalla o directo a la impresora los listados de remate
Private Sub cmdImpListRema_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    lsCadImprimir = ""
    
    If Not IsDate(txtFecRemate) Then
        MsgBox "Ingrese una Fecha correcta", vbInformation, "Aviso"
        Exit Sub
    End If

    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
        'ImprimeContRema
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeListadoParaRemate(Format(Me.txtFecRemate.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnDiasAtrasoAvisoRemate, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, fnFactorPrecioBaseRemate, CCur(Me.txtPreOro14), CCur(Me.txtPreOro16), _
                        CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumRemate.Text, IIf(fnJoyasDet = 1, True, False), lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                            MsgBox lsmensaje, vbInformation, "Aviso"
                            Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Listado Contratos para Remate", True
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, lsCadImprimir
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If

End Sub

Private Sub CargaParametros()
Dim loparam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loparam = New COMDColocPig.DCOMColPCalculos
    fnDiasAtrasoAvisoVencimiento = Int(loparam.dObtieneColocParametro(gConsColPDiasAtrasoCartaVenc))
    fnDiasAtrasoAvisoRemate = Int(loparam.dObtieneColocParametro(gConsColPDiasAtrasoParaRemate))
    fnFactorPrecioBaseRemate = loparam.dObtieneColocParametro(gConsColPFactorPrecioBaseRemate)
    fnTasaPreparacionRemate = loparam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnTasaImpuesto = loparam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnTasaCustodiaVencida = loparam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnCostoCorrespondencia = loparam.dObtieneColocParametro(3035)
    
Set loparam = Nothing
    fnPrevioMax = 5000
    fnLineasMax = 55
    fnHojaFiMax = 66

Set loConstSis = New COMDConstSistema.NCOMConstSistema
    fnJoyasDet = loConstSis.LeeConstSistema(109) ' Joyas en Detalle
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsRemateCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsRemateCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing

End Sub

Private Sub CargaFirmaJefeAge()
   
Dim RStream As adodb.Stream
Dim oDCred As COMDCredito.DCOMCredDoc
Dim pR As adodb.Recordset
Dim sRutaTmp As String
Dim sRutaImg As String
Dim sApMaterno As String, sApPaterno As String, sNombres As String

    Set oDCred = New COMDCredito.DCOMCredDoc
    Set pR = oDCred.RecuperaJefeAgencia(gsCodAge)
    Set oDCred = Nothing
    
    sApPaterno = PstaNombre(pR!cPersNombre, True)
    If InStr(1, pR!cPersNombre, "/", vbTextCompare) <> 0 Then
        sApPaterno = Trim(Mid(pR!cPersNombre, 1, InStr(1, pR!cPersNombre, "/", vbTextCompare) - 1))
    End If
    
    If InStr(1, pR!cPersNombre, "\", vbTextCompare) <> 0 Then
        sApMaterno = Mid(pR!cPersNombre, InStr(1, pR!cPersNombre, "\", vbTextCompare) + 1, InStr(1, pR!cPersNombre, ",", vbTextCompare) - InStr(1, pR!cPersNombre, "\", vbTextCompare) - 1)
    Else
        sApMaterno = Mid(pR!cPersNombre, InStr(1, pR!cPersNombre, "/", vbTextCompare) + 1, InStr(1, pR!cPersNombre, ",", vbTextCompare) - InStr(1, pR!cPersNombre, "/", vbTextCompare) - 1)
    End If
    sNombres = LTrim(Mid(pR!cPersNombre, InStr(1, pR!cPersNombre, ",", vbTextCompare) + 1, 100))
    
    
    Me.lblNombre = sNombres & " " & sApPaterno & " " & sApMaterno 'pR!cPersNombre
    Me.lblCargo = pR!Cargo
    
    Set RStream = New adodb.Stream
    RStream.Type = adTypeBinary
    RStream.Open
    
    If Len(Trim(pR!firma)) = 0 Then
        MsgBox "Escanear y registrar la firma del Responsable de la Agencia.", vbInformation, " Aviso "
        Exit Sub
    End If
    
    RStream.Write pR.Fields("Firma").value

    sRutaTmp = App.path & "\spooler\FirJefe.bmp"
    RStream.SaveToFile App.path & "\spooler\FirJefe.bmp", adSaveCreateOverWrite
    sRutaImg = sRutaTmp
    If Len(Trim(sRutaImg)) > 0 Then
        ImgFirma.Picture = LoadPicture(sRutaImg)
    End If
    Kill (App.path & "\spooler\FirJefe.bmp")
    RStream.Close
    Set RStream = Nothing
End Sub
