VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInversionesProvision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   1710
   ClientTop       =   2640
   ClientWidth     =   11835
   Icon            =   "frmInversionesProvision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   11835
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   17
      Top             =   5040
      Width           =   11580
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   8760
         TabIndex        =   19
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   345
         Left            =   10260
         TabIndex        =   18
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame FraConcepto 
      Height          =   1395
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   11580
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
         TabIndex        =   14
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   1020
         Width           =   1680
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   750
         Left            =   870
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   195
         Width           =   10635
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
         TabIndex        =   16
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   495
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
      Height          =   3510
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      Begin VB.Frame fraFndMutuos 
         Height          =   570
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   11280
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
            Left            =   5760
            TabIndex        =   23
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1590
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
            Left            =   1545
            TabIndex        =   22
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1830
         End
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
            Left            =   9450
            TabIndex        =   21
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1590
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Valor Cuota :"
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
            Left            =   8250
            TabIndex        =   24
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nro Cuotas :"
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
            Left            =   4560
            TabIndex        =   25
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label Label6 
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
            Left            =   735
            TabIndex        =   26
            Top             =   225
            Width           =   720
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   4470
            Top             =   165
            Width           =   2895
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   630
            Top             =   165
            Width           =   2760
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   8160
            Top             =   165
            Width           =   2895
         End
      End
      Begin VB.Frame fraInteres 
         Height          =   570
         Left            =   4395
         TabIndex        =   3
         Top             =   2865
         Width           =   7080
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
            Left            =   5400
            TabIndex        =   5
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1590
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
            Left            =   1545
            TabIndex        =   4
            Tag             =   "2"
            Text            =   "0.00"
            Top             =   195
            Width           =   1830
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
            Left            =   4320
            TabIndex        =   7
            Top             =   225
            Width           =   975
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
            Left            =   735
            TabIndex        =   6
            Top             =   225
            Width           =   720
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   4110
            Top             =   165
            Width           =   2895
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            FillColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   630
            Top             =   165
            Width           =   2760
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   345
         Left            =   10080
         TabIndex        =   2
         Top             =   180
         Width           =   1410
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   218
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   870
         TabIndex        =   8
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
         TabIndex        =   9
         Top             =   585
         Width           =   11445
         _extentx        =   20188
         _extenty        =   3995
         cols0           =   21
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   $"frmInversionesProvision.frx":030A
         encabezadosanchos=   "350-1000-1800-1800-1200-900-800-900-900-500-1000-500-900-0-0-0-1200-0-0-0-0"
         font            =   "frmInversionesProvision.frx":03C8
         font            =   "frmInversionesProvision.frx":03F0
         font            =   "frmInversionesProvision.frx":0418
         font            =   "frmInversionesProvision.frx":0440
         font            =   "frmInversionesProvision.frx":0468
         fontfixed       =   "frmInversionesProvision.frx":0490
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Interes al:"
         Height          =   210
         Left            =   105
         TabIndex        =   11
         Top             =   270
         Width           =   705
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
         TabIndex        =   10
         Top             =   255
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmInversionesProvision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaHaber As String
Dim lsCtaDebe As String
Dim lbCapPerdido As Boolean
Dim lnIntFndMutuo As Currency
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Function Valida() As Boolean
    Valida = True
    If Me.fgIF.TextMatrix(Me.fgIF.row, 1) = "" Then
        MsgBox "El Dato seleccionado de la Lista no es Valida", vbInformation, "Aviso!"
        Valida = False
        Exit Function
    End If

    If Trim(Right(fgIF.TextMatrix(Me.fgIF.row, 1), 2)) <> "3" Then
        If Val(Me.txtCalculado.Text) = 0 Or Me.txtCalculado.Text = "" Then
            MsgBox "El Interes Calculado es 0", vbYesNo, "Aviso!"
               Valida = False
               Exit Function
        ElseIf Val(Me.txtCalculado.Text) < 0 Then
               MsgBox "El Interes Calculado es menor que 0", vbYesNo, "Aviso!"
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
        MsgBox "Ingrese un Descripcion o Glosa de la Operacion", vbInformation, "Aviso!"
        Valida = False
        txtMovDesc.SetFocus
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
    Dim nTpoInv As Integer 'PASI20150921 ERS0712015
    Dim oOpe As New DOperacion
   
       sCtaConf = Me.fgIF.TextMatrix(Me.fgIF.row, 18)
       nTpoInv = CInt(Trim(Right(Me.fgIF.TextMatrix(Me.fgIF.row, 1), 2)))
        If Not lbCapPerdido Then
            'ALPA 20140601--Incidente INC1406020035
            If sCtaConf = "1315010101" Or sCtaConf = "1325010101" Then
                lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "30" + Left(Right(sCtaConf, Len(sCtaConf) - 3), Len(sCtaConf) - 5)
            Else
                'Modificado PASI20150921 ERS0712015
                'lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "30" + Right(sCtaConf, Len(sCtaConf) - 3)
                
                'PASI***************
                'Select Case nTpoInv
                '    Case 2 ' Certificado de deposito BCRP
                '        lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "304020101"
                '    Case 6 'Op. Reportes
                '        lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "305180107"
                '    Case 7 'Letras de Tesoro, Papeles
                '        If sCtaConf = "13" + Mid(gsOpeCod, 3, 1) + "4071901" Then 'Papeles
                '            lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "304071901"
                '        Else 'Letras de Tesoro
                '            lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "3040101"
                '        End If
                '    Case 1, 3, 4, 5
                '        lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "30" + Right(sCtaConf, Len(sCtaConf) - 3)
                'End Select
                '********************
                Select Case nTpoInv
                    Case 2 ' Certificado de deposito BCRP
                        lsCtaHaber = oOpe.ObtieneCtaHaberxProvInversion(nTpoInv, 0)
                        lsCtaHaber = Replace(lsCtaHaber, "M", Mid(gsOpeCod, 3, 1))
                    Case 6 'Op. Reportes
                        lsCtaHaber = oOpe.ObtieneCtaHaberxProvInversion(nTpoInv, 0)
                        lsCtaHaber = Replace(lsCtaHaber, "M", Mid(gsOpeCod, 3, 1))
                    Case 7 'Letras de Tesoro, Papeles
                        If sCtaConf = "13" + Mid(gsOpeCod, 3, 1) + "4071901" Then 'Papeles
                            lsCtaHaber = oOpe.ObtieneCtaHaberxProvInversion(nTpoInv, 0)
                            lsCtaHaber = Replace(lsCtaHaber, "M", Mid(gsOpeCod, 3, 1))
                        Else 'Letras de Tesoro
                            lsCtaHaber = oOpe.ObtieneCtaHaberxProvInversion(nTpoInv, 1)
                        lsCtaHaber = Replace(lsCtaHaber, "M", Mid(gsOpeCod, 3, 1))
                        End If
                    Case 1, 3, 4, 5
                        lsCtaHaber = "51" + Mid(gsOpeCod, 3, 1) + "30" + Right(sCtaConf, Len(sCtaConf) - 3)
                End Select
                'end pasi
            End If
            If Not oOpe.ValidaCtaCont(lsCtaHaber) Then
               MsgBox "Falta definir Cuenta Contable (H) de la Provision Valida: " & lsCtaHaber, vbInformation, "¡Aviso!"
               obtenerCuentasCont = False
               Exit Function
            End If
            
            'ALPA 20130509*************************************************************************
            'lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + "80" + Right(sCtaConf, Len(sCtaConf) - 3)
            
            'Modificado PASI20150921 ERS0712015
            'lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + Right(sCtaConf, Len(sCtaConf) - 3)
            
            'PASI***
            'Select Case nTpoInv
            '    Case 2
            '        lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + "40201"
            '    Case 6
            '        lsCtaDebe = "15" + Mid(gsOpeCod, 3, 1) + "70911"
            '    Case 7
            '        If sCtaConf = "13" + Mid(gsOpeCod, 3, 1) + "4071901" Then 'Papeles
            '            lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + "4071901"
            '        Else 'Letras de Tesoro
            '            lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + "4010101"
            '        End If
            '    Case 1, 3, 4, 5
            '        lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + Right(sCtaConf, Len(sCtaConf) - 3)
            'End Select
            '*******
            
             Select Case nTpoInv
                Case 2
                    lsCtaDebe = oOpe.ObtieneCtaDebexProvInversion(nTpoInv, 0)
                    lsCtaDebe = Replace(lsCtaDebe, "M", Mid(gsOpeCod, 3, 1))
                Case 6
                    lsCtaDebe = oOpe.ObtieneCtaDebexProvInversion(nTpoInv, 0)
                    lsCtaDebe = Replace(lsCtaDebe, "M", Mid(gsOpeCod, 3, 1))
                Case 7
                    If sCtaConf = "13" + Mid(gsOpeCod, 3, 1) + "4071901" Then 'Papeles
                        lsCtaDebe = oOpe.ObtieneCtaDebexProvInversion(nTpoInv, 0)
                        lsCtaDebe = Replace(lsCtaDebe, "M", Mid(gsOpeCod, 3, 1))
                    Else 'Letras de Tesoro
                        lsCtaDebe = oOpe.ObtieneCtaDebexProvInversion(nTpoInv, 1)
                        lsCtaDebe = Replace(lsCtaDebe, "M", Mid(gsOpeCod, 3, 1))
                    End If
                Case 1, 3, 4, 5
                    lsCtaDebe = "13" + Mid(gsOpeCod, 3, 1) + Right(sCtaConf, Len(sCtaConf) - 3)
            End Select
            
            'end pasi
            
            '**************************************************************************************
            If Not oOpe.ValidaCtaCont(lsCtaDebe) Then
               MsgBox "Falta definir Cuenta Contable (D) de la Provision Valida: " & lsCtaDebe, vbInformation, "¡Aviso!"
               obtenerCuentasCont = False
               Exit Function
            End If
        Else
            lsCtaDebe = "43" + Mid(gsOpeCod, 3, 1) + "101030705" + Right(sCtaConf, 2)
            If Not oOpe.ValidaCtaCont(lsCtaDebe) Then
               MsgBox "Falta definir Cuenta Contable (D) de la Provision Valida: " & lsCtaDebe, vbInformation, "¡Aviso!"
               obtenerCuentasCont = False
               Exit Function
            End If
        
        End If
End Function

Private Sub cmdAceptar_Click()
    Dim oFun As New NContFunciones
    Dim lbModificaMov As Boolean
        
    lbModificaMov = oFun.PermiteModificarAsiento(Format(Me.txtFecha.Text, "yyyymmdd"), False)
    If Not lbModificaMov Then
          MsgBox "Fecha de Provision corresponde a un mes ya Cerrado.? ", vbInformation, "¡Confirmación!"
          Exit Sub
    End If
    
    If Not Valida Then Exit Sub
    'GrabaProvInversion
    Dim sMovNro As String
       Dim oCaja As New nCajaGeneral
       Dim oCon As New NContFunciones
         
         
       If MsgBox("Esta Seguro de Guardar los Datos", vbYesNo, "!Aviso¡") = vbYes Then
            If Not obtenerCuentasCont Then Exit Sub
            'sMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            sMovNro = oCon.GeneraMovNro(Me.txtFecha.Text, gsCodAge, gsCodUser)
            With Me.fgIF
                oCaja.GrabaProvInversion sMovNro, gsOpeCod, Me.txtMovDesc.Text, lsCtaDebe, _
                                           lsCtaHaber, Me.txtImporte.Text, _
                                           IIf(Trim(Right(.TextMatrix(.row, 1), 2)) <> "3", Me.txtCalculado.Text, lnIntFndMutuo), _
                                           .TextMatrix(.row, 13), .TextMatrix(.row, 14), .TextMatrix(.row, 15), Me.txtFecha.Text, .TextMatrix(.row, 17), _
                                           getFinMes, Trim(Right(.TextMatrix(.row, 1), 2)), Me.txtCapital.Text, Me.txtValorCuota.Text, Me.txtNroCuotas.Text
                                           'lbCapPerdido, Trim(Right(.TextMatrix(.Row, 1), 2))
                                           
                                  
                ImprimeAsientoContable sMovNro, "", "", "", True, False
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grado la Operación "
                Set objPista = Nothing
                '****
                If MsgBox("Desea Realizar otra Operacion ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                        .EliminaFila .row
                                             
                        Me.txtInteres.Text = "0.00"
                        Me.txtCalculado.Text = "0.00"
                        Me.txtImporte.Text = "0.00"
                        Me.txtMovDesc.Text = ""
                                                
                Else
                    Unload Me
                End If
             End With
        End If
End Sub
Private Function getFinMes() As Boolean
    getFinMes = False
    Dim dFinMes As Date
    dFinMes = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" + Mid(Me.txtFecha.Text, 4, 2) + "/" + Right(Me.txtFecha.Text, 2))))
    If DateDiff("D", Me.txtFecha.Text, dFinMes) = 0 Then
        getFinMes = True
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
     Set rs = oCaja.obtenerInversionesVigentes(IIf(gsOpeCod = "421304", "421302", "422302"), Me.txtFecha.Text, Trim(Right(Me.cmbTipo.Text, 2)))
       
           
    
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
                Me.fraFndMutuos.Visible = False
                Me.fraInteres.Visible = True
                Me.txtInteres.Text = Format(.TextMatrix(.row, 6), "##,##0.00")
                Me.txtCalculado.Text = Format(.TextMatrix(.row, 10), "##,##0.00")
                Me.txtImporte.Text = Format(CCur(.TextMatrix(.row, 6)) + CCur(.TextMatrix(.row, 10)), "##,##0.00")
            Else
                Me.fraInteres.Visible = False
                Me.fraFndMutuos.Visible = True
                Me.txtCapital.Text = Format(.TextMatrix(.row, 4), "##,##0.00")
                Me.txtNroCuotas.Text = Format(.TextMatrix(.row, 20), "##,##0.0000")
                'Me.txtImporte.Text = Format(CCur(.TextMatrix(.Row, 6)) + CCur(.TextMatrix(.Row, 10)), "##,##0.00")
            
            End If
        End If
    
    End With
End Sub
Private Sub Limpiar()
    Me.fraFndMutuos.Visible = False
    Me.fraInteres.Visible = True
    Me.txtInteres.Text = "0.00"
    Me.txtCalculado.Text = "0.00"
    Me.txtImporte.Text = "0.00"
    Me.txtCapital.Text = "0.00"
    Me.txtNroCuotas.Text = "0.00"
    Me.txtValorCuota.Text = "0.00"
    Me.txtMovDesc.Text = ""
End Sub
Private Sub Form_Load()
    Dim oOpe As New DOperacion
    If gsOpeCod = "421304" Then
        Me.Caption = "Provision de Inversiones MN"
    ElseIf gsOpeCod = "422304" Then
        Me.Caption = "Provision de Inversiones ME"
    End If
    Me.txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
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
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub
Private Sub txtValorCuota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtImporte.Text = Format(CCur(Me.txtNroCuotas.Text) * CDbl(Me.txtValorCuota.Text), "##,##0.00")
        If CCur(txtImporte.Text) < CCur(Me.txtCapital.Text) Then
            lbCapPerdido = True
            lnIntFndMutuo = CCur(Me.txtCapital.Text) - CCur(txtImporte.Text)
        Else
            lbCapPerdido = False
            lnIntFndMutuo = CCur(txtImporte.Text) - CCur(Me.txtCapital.Text)
        End If
        
        Me.txtMovDesc.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtValorCuota.Text) > 0 Then
            txtValorCuota.Text = Mid(txtValorCuota.Text, 1, Len(txtValorCuota.Text) - 1)
            txtValorCuota.SelStart = Len(txtValorCuota.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If

End Sub
