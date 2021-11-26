VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInversionesCancelacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   2235
   ClientTop       =   2985
   ClientWidth     =   11820
   Icon            =   "frmInversionesCancelacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11820
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   8205
      TabIndex        =   16
      Top             =   5160
      Width           =   3495
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   2100
         TabIndex        =   18
         Top             =   360
         Width           =   1245
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.Frame fraTransferencia 
      Caption         =   "Transferencia a :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   945
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   7965
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   315
         Left            =   915
         TabIndex        =   12
         Top             =   225
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDesCtaIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   915
         TabIndex        =   15
         Top             =   600
         Width           =   6840
      End
      Begin VB.Label lblDescIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3510
         TabIndex        =   14
         Top             =   225
         Width           =   4245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta N° :"
         Height          =   210
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.Frame FraConcepto 
      Height          =   1395
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   11580
      Begin VB.TextBox txtMovDesc 
         Height          =   750
         Left            =   750
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   195
         Width           =   10755
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   285
         Left            =   9795
         TabIndex        =   8
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   1020
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL :"
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
         Height          =   195
         Left            =   8640
         TabIndex        =   10
         Top             =   1065
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         FillColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   8460
         Top             =   990
         Width           =   3045
      End
   End
   Begin VB.Frame fradatosGen 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      Begin VB.Frame fraImpRenta 
         Height          =   630
         Left            =   8520
         TabIndex        =   29
         Top             =   2985
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtImpRenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   1435
            TabIndex        =   30
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Imp. Renta :"
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
            Height          =   195
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Width           =   1065
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   240
            Top             =   165
            Width           =   2415
         End
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   218
         Width           =   2055
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   345
         Left            =   10080
         TabIndex        =   5
         Top             =   180
         Width           =   1410
      End
      Begin VB.Frame fraCapital 
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
         Height          =   630
         Left            =   180
         TabIndex        =   2
         Top             =   2985
         Width           =   8220
         Begin VB.TextBox txtValorCuota 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   6810
            TabIndex        =   28
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtNroCuotas 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   3915
            TabIndex        =   27
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtCalculado 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            Height          =   285
            Left            =   6810
            TabIndex        =   25
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1215
         End
         Begin VB.TextBox txtInteres 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   3915
            TabIndex        =   23
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1215
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
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
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1005
            TabIndex        =   3
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1440
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Calculado :"
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
            Height          =   195
            Left            =   5490
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   5400
            Top             =   165
            Width           =   2655
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Interes :"
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
            Height          =   195
            Left            =   2880
            TabIndex        =   24
            Top             =   240
            Width           =   960
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   2760
            Top             =   165
            Width           =   2385
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capital :"
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
            Height          =   195
            Left            =   210
            TabIndex        =   4
            Top             =   240
            Width           =   720
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   150
            Top             =   165
            Width           =   2325
         End
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   870
         TabIndex        =   1
         Top             =   210
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.FlexEdit fgIF 
         Height          =   2265
         Left            =   45
         TabIndex        =   22
         Top             =   600
         Width           =   11445
         _extentx        =   20188
         _extenty        =   2302
         cols0           =   21
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   $"frmInversionesCancelacion.frx":030A
         encabezadosanchos=   "350-1000-1800-1800-1200-900-800-900-900-500-1000-500-900-0-0-0-1200-0-0-0-0"
         font            =   "frmInversionesCancelacion.frx":03C8
         font            =   "frmInversionesCancelacion.frx":03F0
         font            =   "frmInversionesCancelacion.frx":0418
         font            =   "frmInversionesCancelacion.frx":0440
         font            =   "frmInversionesCancelacion.frx":0468
         fontfixed       =   "frmInversionesCancelacion.frx":0490
         backcolorcontrol=   11861226
         backcolorcontrol=   11861226
         backcolorcontrol=   11861226
         lbultimainstancia=   -1
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-10-X-X-X-X-X-X-X-X-X-X"
         textstylefixed  =   3
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-R-R-R-C-L-L-R-L-C-C-C-C-R-C-C-C-C"
         formatosedit    =   "0-0-0-0-2-2-2-0-0-0-2-0-0-0-0-0-2-0-0-0-0"
         textarray0      =   "N°"
         lbeditarflex    =   -1
         lbformatocol    =   -1
         lbpuntero       =   -1
         lbordenacol     =   -1
         colwidth0       =   345
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interes al:"
         Height          =   210
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmInversionesCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaIntereses As String
Dim lsCtaProvision As String
Dim lsCtaHaberCap As String
Dim lsCtaDebe As String
Dim lsCtaCapPerdido As String
Dim lbCapPerdido As Boolean
Dim lnIntFndMutuo As Currency
Dim lsAntiCtaImpRenta As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdAceptar_Click()
        
        Dim oFun As New NContFunciones
        Dim lbEliminaMov As Boolean
        
       lbEliminaMov = oFun.PermiteModificarAsiento(Format(Me.txtFecha.Text, "yyyymmdd"), False)
       If Not lbEliminaMov Then
          MsgBox "Fecha de Cancelacion corresponde a un mes ya Cerrado.? ", vbInformation, "¡Confirmación!"
          Exit Sub
       End If
       
       If Not Valida Then Exit Sub
       
       Dim sMovNro As String
       Dim lsMovNroACAdd    As String 'PASI20150921 ERS0472015
       Dim oCaja As New nCajaGeneral
       Dim oCon As New NContFunciones
         
         
        If MsgBox("Esta Seguro de Guardar los Datos", vbYesNo, "!Aviso¡") = vbYes Then
            If Not obtenerCuentasCont Then Exit Sub
            'sMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            With Me.fgIF
                sMovNro = oCon.GeneraMovNro(Me.txtFecha.Text, gsCodAge, gsCodUser)
                If Trim(Right(.TextMatrix(.row, 1), 2)) = 6 Then 'PASI20150921 ERS0712015
                    Sleep (1000)
                    lsMovNroACAdd = oCon.GeneraMovNro(txtFecha.Text, gsCodAge, gsCodUser)
                End If
                oCaja.GrabaCancelInversion sMovNro, gsOpeCod, Me.txtMovDesc.Text, lsCtaDebe, _
                                           txtImporte.Text, Me.txtBuscaEntidad, lsCtaProvision, IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", txtInteres.Text, 0), _
                                           IIf(lbCapPerdido, lsCtaCapPerdido, lsCtaIntereses), IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", txtCalculado.Text, lnIntFndMutuo), lsCtaHaberCap, txtCapital.Text, _
                                           .TextMatrix(.row, 13), .TextMatrix(.row, 14), .TextMatrix(.row, 15), Me.txtFecha.Text, _
                                           .TextMatrix(.row, 17), Right(.TextMatrix(.row, 1), 1), IIf(Mid(gsOpeCod, 3, 1) = "1", "1", "2"), _
                                           .TextMatrix(.row, 19), IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", .TextMatrix(.row, 11), 0), .TextMatrix(.row, 7), IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", .TextMatrix(.row, 12), Me.txtFecha.Text), _
                                           IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", .TextMatrix(.row, 5), 0), IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", .TextMatrix(.row, 16), 0), .TextMatrix(.row, 2), .TextMatrix(.row, 9), .TextMatrix(.row, 4), _
                                           Me.txtValorCuota.Text, Me.txtNroCuotas.Text, Me.txtImpRenta.Text, lsAntiCtaImpRenta, lsMovNroACAdd
                                           
                  
                ImprimeAsientoContable sMovNro, "", "", "", True, False
                If Trim(Right(.TextMatrix(.row, 1), 2)) = 6 Then  'PASI20150921 ERS0712015
                    ImprimeAsientoContable lsMovNroACAdd
                End If
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
                If MsgBox("Desea Realizar otra Cancelacion ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                        .EliminaFila .row
                        Me.txtBuscaEntidad.Text = ""
                        Limpiar
                        Me.lblDescIfTransf = ""
                        Me.lblDesCtaIfTransf = ""
                       
                        
                Else
                    Unload Me
                End If
            End With
        End If
End Sub
Private Function Valida() As Boolean
    Valida = True
    If Me.fgIF.TextMatrix(Me.fgIF.row, 1) = "" Then
        MsgBox "El Dato seleccionado de la Lista no es Valida", vbInformation, "Aviso!"
        Valida = False
        Exit Function
    End If
    If Val(Me.txtCapital.Text) = 0 Or Me.txtCapital.Text = "" Then
        MsgBox "El Capital no es Valido", vbInformation, "Aviso!"
        Valida = False
        Exit Function
    End If
    
    If Trim(Right(fgIF.TextMatrix(Me.fgIF.row, 1), 2)) <> "3" Then
        If (Val(Me.txtInteres.Text) = 0 Or Me.txtInteres.Text = "") Then
            If MsgBox("El Interes es 0, Desea continuar", vbYesNo, "Aviso!") = vbNo Then
                Valida = False
                Exit Function
            End If
         End If
         If (Val(Me.txtCalculado.Text) = 0 Or Me.txtCalculado.Text = "") Then
            If MsgBox("El Interes Calculado es 0, Desea continuar", vbYesNo, "Aviso!") = vbNo Then
                Valida = False
                Exit Function
            End If
        ElseIf Val(Me.txtCalculado.Text) < 0 Then
           MsgBox "El Interes Calculado es menor que 0", vbInformation, "Aviso!"
           Valida = False
               Exit Function
        End If
    End If
    
    If Val(Me.txtImporte.Text) = 0 Or Me.txtImporte.Text = "" Then
        MsgBox "El Importe Total No es Valido", vbInformation, "Aviso!"
        Valida = False
        Exit Function
    End If
    
    If Me.txtMovDesc.Text = "" Then
        MsgBox "Ingrese una Descripcion o Glosa de la Operacion", vbInformation, "Aviso!"
        Valida = False
        txtMovDesc.SetFocus
        Exit Function
    End If
    
    If Me.txtBuscaEntidad.Text = "" Then
        MsgBox "La Cuenta Contable no es Valida", vbInformation, "Aviso!"
        Valida = False
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
    
   If Trim(Right(fgIF.TextMatrix(Me.fgIF.row, 1), 2)) = "3" Then
        If Me.txtValorCuota.Text = "" Or Val(Me.txtValorCuota.Text) = 0 Then
            MsgBox "Debe Ingresar el Valor Cuota", vbInformation, "Aviso!"
            Valida = False
            txtValorCuota.SetFocus
            Exit Function
        End If
     End If
    
End Function
Private Function obtenerCuentasCont() As Boolean
    obtenerCuentasCont = True
    Dim sCtaConf As String
    Dim oOpe As New DOperacion
    Dim nTpoInv As Integer 'PASI20150922 ERS0712015
   
       sCtaConf = Me.fgIF.TextMatrix(Me.fgIF.row, 18)
       nTpoInv = CInt(Trim(Right(Me.fgIF.TextMatrix(Me.fgIF.row, 1), 2)))
       
       'Si NO existe capital perdido
        If Not lbCapPerdido Then
            'ALPA20130719*********************************************************************************
            'lsCtaProvision = "13" + Mid(gsOpeCod, 3, 1) + "80" + Right(sCtaConf, Len(sCtaConf) - 3)
            'Modificado PASI20150922 ERS0712015
            'lsCtaProvision = "13" + Mid(gsOpeCod, 3, 1) + Right(sCtaConf, Len(sCtaConf) - 3)
            
            'PASI*******
            'Select Case nTpoInv
            '    Case 2
            '        lsCtaProvision = "13" + Mid(gsOpeCod, 3, 1) + "40201"
            '    Case 6
            '        lsCtaProvision = "15" + Mid(gsOpeCod, 3, 1) + "70911"
            '    Case 7
            '        lsCtaProvision = "13" + Mid(gsOpeCod, 3, 1) + "4010101"
            '    Case 1, 3, 4, 5
            '        lsCtaProvision = "13" + Mid(gsOpeCod, 3, 1) + Right(sCtaConf, Len(sCtaConf) - 3)
            'End Select
            '************
            
            Select Case nTpoInv
                Case 2, 6, 7
                    lsCtaProvision = oOpe.ObtieneCtaProvxCancelInversion(nTpoInv)
                    lsCtaProvision = Replace(lsCtaProvision, "M", Mid(gsOpeCod, 3, 1))
                Case 1, 3, 4, 5
                    lsCtaProvision = "13" + Mid(gsOpeCod, 3, 1) + Right(sCtaConf, Len(sCtaConf) - 3)
            End Select
            
            'end pasi
            '*********************************************************************************************
            
            If Not oOpe.ValidaCtaCont(lsCtaProvision) Then
               MsgBox "Falta definir Cuenta Contable de Interes Provisionado Valida: " & lsCtaProvision, vbInformation, "¡Aviso!"
               obtenerCuentasCont = False
               Exit Function
            End If

        Else
            lsCtaCapPerdido = "43" + Mid(gsOpeCod, 3, 1) + "101030705" + Right(sCtaConf, 2)
            If Not oOpe.ValidaCtaCont(lsCtaCapPerdido) Then
               MsgBox "Falta definir Cuenta Contable de Perdida de Inversion Valida: " & lsCtaCapPerdido, vbInformation, "¡Aviso!"
               obtenerCuentasCont = False
               Exit Function
            End If
        
        End If
        
        'Modificado PASI20150922 ERS0712015
        'lsCtaIntereses = "51" + Mid(gsOpeCod, 3, 1) + "30" + Right(sCtaConf, Len(sCtaConf) - 3)
        
        'PASI**************
        ' Select Case nTpoInv
        '    Case 2
        '        lsCtaIntereses = "51" + Mid(gsOpeCod, 3, 1) + "304020101"
        '    Case 6
        '        lsCtaIntereses = "51" + Mid(gsOpeCod, 3, 1) + "305180107"
        '    Case 7
        '        lsCtaIntereses = "51" + Mid(gsOpeCod, 3, 1) + "3040101"
        '    Case 1, 3, 4, 5
        '        lsCtaIntereses = "51" + Mid(gsOpeCod, 3, 1) + "30" + Right(sCtaConf, Len(sCtaConf) - 3)
        'End Select
        '******************
        
        Select Case nTpoInv
            Case 2, 6, 7
                lsCtaIntereses = oOpe.ObtieneCtaInteresesxCancelInversion(nTpoInv)
                lsCtaIntereses = Replace(lsCtaIntereses, "M", Mid(gsOpeCod, 3, 1))
            Case 1, 3, 4, 5
                lsCtaIntereses = "51" + Mid(gsOpeCod, 3, 1) + "30" + Right(sCtaConf, Len(sCtaConf) - 3)
        End Select
        'end pasi
        lsAntiCtaImpRenta = lsCtaIntereses
        If Not oOpe.ValidaCtaCont(lsCtaIntereses) Then
           MsgBox "Falta definir Cuenta Contable de Interes Calculado Valida: " & lsCtaIntereses, vbInformation, "¡Aviso!"
           obtenerCuentasCont = False
           Exit Function
        End If
                    
        'Modificado PASI20150922 ERS0712015
        'lsCtaHaberCap = sCtaConf
        
        'PASI*****
        'Select Case nTpoInv
        '    Case 2
        '        lsCtaHaberCap = "13" + Mid(gsOpeCod, 3, 1) + "40201"
        '    Case 6
        '        lsCtaHaberCap = "15" + Mid(gsOpeCod, 3, 1) + "70911"
        '   Case 7
        '        lsCtaHaberCap = "13" + Mid(gsOpeCod, 3, 1) + "4010101"
        '    Case 1, 3, 4, 5
        '        lsCtaHaberCap = sCtaConf
        'End Select
        '*********
        
        Select Case nTpoInv
            Case 2, 6, 7
                lsCtaHaberCap = oOpe.ObtieneCtaHaberCapxCancelInversion(nTpoInv)
                lsCtaHaberCap = Replace(lsCtaHaberCap, "M", Mid(gsOpeCod, 3, 1))
            Case 1, 3, 4, 5
                lsCtaHaberCap = sCtaConf
        End Select
        
        'end pasi
        If Not oOpe.ValidaCtaCont(lsCtaHaberCap) Then
           MsgBox "Falta definir Cuenta Contable de Capital: " + lsCtaHaberCap, vbInformation, "¡Aviso!"
           obtenerCuentasCont = False
           Exit Function
        End If
            
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", 0, txtBuscaEntidad, ObjEntidadesFinancieras)
        If Not oOpe.ValidaCtaCont(lsCtaDebe) Then
                MsgBox "Falta definir Cuenta Contable para Abonar Capital en Orden: " & lsCtaDebe, vbInformation, "¡Aviso!"
               obtenerCuentasCont = False
                Exit Function
        End If
End Function
Private Sub cmdProcesar_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo ProcesarErr
    
    If ValidaFecha(Me.txtFecha.Text) <> "" Then
        MsgBox "Debe Ingresar una Fecha Valida", vbInformation, "Aviso!"
        txtFecha.SetFocus
        Exit Sub
    End If
 
    Dim oCaja As New nCajaGeneral
     Set rs = oCaja.obtenerInversionesVigentes(IIf(gsOpeCod = "421303", "421302", "422302"), Me.txtFecha.Text, Trim(Right(Me.cmbTipo.Text, 2)))
       
           
    
    fgIF.Clear
    fgIF.FormaCabecera
    fgIF.Rows = 2
    If Not rs.EOF And Not rs.BOF Then
        Set fgIF.Recordset = rs
        fgIF.SetFocus
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ProcesarErr:
        MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgIF_Click()
    With Me.fgIF
            Limpiar
        If .TextMatrix(.row, 1) <> "" Then
            If Trim(Right(.TextMatrix(.row, 1), 2)) <> "3" Then
                Me.txtInteres.Visible = True
                Me.txtCalculado.Visible = True
                Me.txtNroCuotas.Visible = False
                Me.txtValorCuota.Visible = False
                fraImpRenta.Visible = False
                Me.Label3 = "Interes :"
                Me.Label4 = "Calculado :"
                Me.txtCapital.Text = .TextMatrix(.row, 4)
                Me.txtInteres.Text = Format(.TextMatrix(.row, 6), "##,##0.00")
                Me.txtCalculado.Text = Format(.TextMatrix(.row, 10), "##,##0.00")
                Me.txtImporte.Text = Format(CCur(.TextMatrix(.row, 4)) + CCur(.TextMatrix(.row, 6)) + CCur(.TextMatrix(.row, 10)), "##,##0.00")
                          
            Else
                Me.txtInteres.Visible = False
                Me.txtCalculado.Visible = False
                Me.txtNroCuotas.Visible = True
                Me.txtValorCuota.Visible = True
                fraImpRenta.Visible = True
                Me.Label3 = "Nro.Cuotas:"
                Me.Label4 = "Valor Cuota 2:"
                Me.txtCapital.Text = Format(.TextMatrix(.row, 4), "##,##0.00")
                Me.txtNroCuotas.Text = Format(.TextMatrix(.row, 20), "##,##0.0000")
                
            
            End If
        End If
    
    End With
End Sub
Private Sub Limpiar()
    Me.txtInteres.Visible = True
    Me.txtCalculado.Visible = True
    Me.txtNroCuotas.Visible = False
    Me.txtValorCuota.Visible = False
    fraImpRenta.Visible = False
    Me.txtInteres.Text = "0.00"
    Me.txtCalculado.Text = "0.00"
    Me.txtImporte.Text = "0.00"
    Me.txtCapital.Text = "0.00"
    Me.txtNroCuotas.Text = "0.00"
    Me.txtValorCuota.Text = "0.00"
    Me.txtImpRenta.Text = "0.00"
    Me.txtMovDesc.Text = ""
End Sub
Private Sub Form_Load()
    Dim oOpe As New DOperacion
    If gsOpeCod = "421303" Then
        Me.Caption = "Cancelacion de Inversiones MN"
    ElseIf gsOpeCod = "422303" Then
        Me.Caption = "Cancelacion de Inversiones ME"
    End If
    Me.txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtBuscaEntidad.psRaiz = "Cuentas de Entidades Financieras"
    txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "1")
    cargarTipoInversion
End Sub
Private Sub cargarTipoInversion()
    Dim rsTpoInversion As ADODB.Recordset
    Dim oCons As DConstante
    
    Set rsTpoInversion = New ADODB.Recordset
    Set oCons = New DConstante
    
    Set rsTpoInversion = oCons.CargaConstante(9990)
    If Not (rsTpoInversion.EOF And rsTpoInversion.BOF) Then
        cmbTipo.Clear
        Do While Not rsTpoInversion.EOF
            cmbTipo.AddItem Trim(rsTpoInversion(2)) & Space(100) & Trim(rsTpoInversion(1))
            rsTpoInversion.MoveNext
        Loop
        cmbTipo.AddItem "Todos los Tipos" + Space(70) + "%", 0
        cmbTipo.ListIndex = 0
    End If
    
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
   Dim oCtaIf As New NCajaCtaIF
    lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lblDesCtaIfTransf = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
End Sub


Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub txtImpRenta_Click()
    If CCur(Me.txtImpRenta.Text) = 0# Then
        Me.txtImpRenta.Text = ""
    End If
End Sub

Private Sub txtImpRenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        calcularImporte
        
        Me.txtMovDesc.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtImpRenta.Text) > 0 Then
            txtImpRenta.Text = Mid(txtImpRenta.Text, 1, Len(txtImpRenta.Text) - 1)
            txtImpRenta.SelStart = Len(txtImpRenta.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtValorCuota_Click()
    If CCur(Me.txtValorCuota.Text) = 0# Then
        Me.txtValorCuota.Text = ""
    End If
End Sub

Private Sub txtValorCuota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        calcularImporte
        
        Me.txtImpRenta.Text = ""
        Me.txtImpRenta.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtValorCuota.Text) > 0 Then
            txtValorCuota.Text = Mid(txtValorCuota.Text, 1, Len(txtValorCuota.Text) - 1)
            txtValorCuota.SelStart = Len(txtValorCuota.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtValorCuota_LostFocus()
    If Me.txtValorCuota.Text = "" Then
        Me.txtValorCuota.Text = "0.00"
    ElseIf Me.txtValorCuota.Text <> "" Then
        calcularImporte
    End If

End Sub
Private Sub txtImpRenta_LostFocus()
    If Me.txtImpRenta.Text = "" Then
        Me.txtImpRenta.Text = "0.00"
    ElseIf Me.txtImpRenta.Text <> "" Then
        calcularImporte
    End If
End Sub
Private Sub calcularImporte()
        Me.txtImporte.Text = Format(CCur(Me.txtNroCuotas.Text) * CDbl(IIf(Me.txtValorCuota.Text = "", "0.00", Me.txtValorCuota.Text)) - CDbl(IIf(Me.txtImpRenta.Text = "", "0.00", Me.txtImpRenta.Text)), "##,##0.00")
        If CCur(txtImporte.Text) < CCur(Me.txtCapital.Text) Then
            lbCapPerdido = True
            lnIntFndMutuo = CCur(Me.txtCapital.Text) - CCur(txtImporte.Text)
        Else
            lbCapPerdido = False
            lnIntFndMutuo = CCur(txtImporte.Text) - CCur(Me.txtCapital.Text)
        End If
End Sub
