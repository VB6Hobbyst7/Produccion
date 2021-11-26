VERSION 5.00
Begin VB.Form FrmOpePagosRFA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos RFA"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Pago"
      Height          =   6495
      Left            =   0
      TabIndex        =   5
      Top             =   1110
      Width           =   7035
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   5580
         TabIndex        =   45
         Top             =   6030
         Width           =   1275
      End
      Begin VB.CommandButton cmdmora 
         Caption         =   "&Mora"
         Enabled         =   0   'False
         Height          =   345
         Left            =   4260
         TabIndex        =   44
         Top             =   6030
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   2970
         TabIndex        =   43
         Top             =   6030
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   345
         Left            =   300
         TabIndex        =   42
         Top             =   6030
         Width           =   1275
      End
      Begin VB.Frame Frame4 
         Height          =   1530
         Left            =   300
         TabIndex        =   27
         Top             =   4410
         Width           =   6570
         Begin VB.ComboBox CmbForPag 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmOpePagosRFA.frx":0000
            Left            =   1335
            List            =   "FrmOpePagosRFA.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   210
            Width           =   1785
         End
         Begin VB.TextBox TxtMonPag 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1320
            MaxLength       =   9
            TabIndex        =   28
            Top             =   585
            Width           =   1170
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Left            =   3210
            TabIndex        =   40
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label LblProxfec 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4575
            TabIndex        =   39
            Top             =   975
            Width           =   45
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Prox. fecha Pag :"
            Height          =   195
            Left            =   3180
            TabIndex        =   38
            Top             =   945
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Monto a Pagar"
            Height          =   195
            Left            =   135
            TabIndex        =   37
            Top             =   615
            Width           =   1050
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Saldo de Capital"
            Height          =   195
            Left            =   90
            TabIndex        =   36
            Top             =   915
            Width           =   1680
         End
         Begin VB.Label LblNewSalCap 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1890
            TabIndex        =   35
            Top             =   900
            Width           =   45
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Cuota Pendiente"
            Height          =   195
            Left            =   90
            TabIndex        =   34
            Top             =   1230
            Width           =   1710
         End
         Begin VB.Label LblNewCPend 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1905
            TabIndex        =   33
            Top             =   1230
            Width           =   45
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Estado Credito"
            Height          =   195
            Left            =   3180
            TabIndex        =   32
            Top             =   1215
            Width           =   1035
         End
         Begin VB.Label LblEstado 
            AutoSize        =   -1  'True
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
            TabIndex        =   31
            Top             =   1215
            Width           =   75
         End
         Begin VB.Label LblNumDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4425
            TabIndex        =   30
            Top             =   225
            Width           =   1665
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   240
         TabIndex        =   7
         Top             =   195
         Width           =   6570
         Begin VB.Label LblMontoCuota 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4290
            TabIndex        =   26
            Top             =   1410
            Width           =   840
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Monto Cuota :"
            Height          =   195
            Left            =   2835
            TabIndex        =   25
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label LblTotCuo 
            AutoSize        =   -1  'True
            Caption         =   "Cuotas"
            Height          =   195
            Left            =   1650
            TabIndex        =   24
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label LblForma 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   23
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label Lbl2 
            AutoSize        =   -1  'True
            Caption         =   "Forma Pago"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1395
            Width           =   870
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Deuda a la Fecha : "
            Height          =   195
            Left            =   2835
            TabIndex        =   21
            Top             =   1095
            Width           =   1410
         End
         Begin VB.Label LblTotDeuda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4290
            TabIndex        =   20
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   780
            Width           =   585
         End
         Begin VB.Label LblMoneda 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   18
            Top             =   750
            Width           =   1155
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Linea Credito"
            Height          =   195
            Left            =   105
            TabIndex        =   17
            Top             =   510
            Width           =   930
         End
         Begin VB.Label LblLinCred 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1110
            TabIndex        =   16
            Top             =   495
            Width           =   4950
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   1095
            Width           =   930
         End
         Begin VB.Label LblSalCap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1110
            TabIndex        =   14
            Top             =   1065
            Width           =   1155
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Monto del Credito"
            Height          =   195
            Left            =   2835
            TabIndex        =   13
            Top             =   810
            Width           =   1245
         End
         Begin VB.Label LblMonCred 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4290
            TabIndex        =   12
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   480
         End
         Begin VB.Label LblNomCli 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1110
            TabIndex        =   10
            Top             =   195
            Width           =   4950
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Calificacion :"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1695
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label LblCalMiViv 
            Appearance      =   0  'Flat
            Caption         =   "Mal Pagador"
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
            Height          =   195
            Left            =   1110
            TabIndex        =   8
            Top             =   1710
            Visible         =   0   'False
            Width           =   1395
         End
      End
      Begin VB.CommandButton CmdPlanPagos 
         Caption         =   "&Plan Pagos"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1620
         TabIndex        =   6
         Top             =   6030
         Width           =   1275
      End
      Begin VB.Label lblGastos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4650
         TabIndex        =   68
         Top             =   4080
         Width           =   1275
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Gastos:"
         Height          =   195
         Left            =   3480
         TabIndex        =   67
         Top             =   4200
         Width           =   540
      End
      Begin VB.Label lblComisionCofide 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4650
         TabIndex        =   66
         Top             =   3780
         Width           =   1275
      End
      Begin VB.Label lblComisionCofide1 
         AutoSize        =   -1  'True
         Caption         =   "Comision Cofide:"
         Height          =   195
         Left            =   3480
         TabIndex        =   65
         Top             =   3900
         Width           =   1170
      End
      Begin VB.Label lblMora 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4650
         TabIndex        =   64
         Top             =   3480
         Width           =   1275
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Mora:"
         Height          =   195
         Left            =   3480
         TabIndex        =   63
         Top             =   3630
         Width           =   405
      End
      Begin VB.Label lblDiasAtraso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4650
         TabIndex        =   62
         Top             =   3180
         Width           =   1275
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dias de Atraso:"
         Height          =   195
         Left            =   3450
         TabIndex        =   61
         Top             =   3270
         Width           =   1080
      End
      Begin VB.Label lblDifInt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2130
         TabIndex        =   60
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label lblDifCap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   59
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "DIF:"
         Height          =   195
         Left            =   360
         TabIndex        =   58
         Top             =   3930
         Width           =   300
      End
      Begin VB.Label lblRFCInt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2130
         TabIndex        =   57
         Top             =   3570
         Width           =   1275
      End
      Begin VB.Label lblRFCCap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   56
         Top             =   3570
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
         Height          =   195
         Left            =   330
         TabIndex        =   55
         Top             =   3630
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Interes"
         Height          =   195
         Left            =   2160
         TabIndex        =   54
         Top             =   2970
         Width           =   480
      End
      Begin VB.Label lblRfaInt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2130
         TabIndex        =   53
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Capital"
         Height          =   195
         Left            =   810
         TabIndex        =   52
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label lblRFACap 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   810
         TabIndex        =   51
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Line Line1 
         X1              =   300
         X2              =   6750
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RFA:"
         Height          =   195
         Left            =   330
         TabIndex        =   50
         Top             =   3300
         Width           =   360
      End
      Begin VB.Label lblDeudaPendiente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1620
         TabIndex        =   49
         Top             =   2250
         Width           =   1545
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Deuda Pendiente"
         Height          =   195
         Left            =   270
         TabIndex        =   48
         Top             =   2310
         Width           =   1245
      End
      Begin VB.Label LblCPend 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1605
         TabIndex        =   47
         Top             =   2580
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Pendiente"
         Height          =   210
         Left            =   270
         TabIndex        =   46
         Top             =   2610
         Width           =   1185
      End
   End
   Begin VB.Frame FramaAgencias 
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   7035
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
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
         Left            =   3660
         TabIndex        =   4
         Top             =   180
         Width           =   900
      End
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4770
         TabIndex        =   2
         Top             =   150
         Width           =   2115
         Begin VB.ListBox LstCred 
            Height          =   450
            ItemData        =   "FrmOpePagosRFA.frx":0004
            Left            =   60
            List            =   "FrmOpePagosRFA.frx":0006
            TabIndex        =   3
            Top             =   270
            Width           =   1980
         End
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblAgencias 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   69
         Top             =   600
         Width           =   3465
      End
   End
End
Attribute VB_Name = "FrmOpePagosRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOperacion As String
Dim bPrepago As Integer
Dim nCalendDinamTipo As Integer
Dim bCalenDinamic As Boolean
Private bCalenCuotaLibre As Boolean
Public Sub Inicia(sCodOpe As String)
    sOperacion = sCodOpe
    Me.Show 1
End Sub


Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    End If
End Sub

Private Sub CmdBuscar_Click()
    Dim oCredito As DCredito
    Dim rs As ADODB.Recordset
    Dim oPers As UPersona
    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        Set oCredito = New DCredito
        Set rs = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigMor, gColocEstVigVenc, gColocEstVigNorm, gColocEstRefMor, gColocEstRefVenc, gColocEstRefNorm, 2031, 2032))
        Do While Not rs.EOF
            LstCred.AddItem rs!cCtaCod
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El cliente no tiene creditos vigentes", vbInformation, "Aviso"
    End If
    
End Sub

Private Function CargarDatos(ByVal psCtaCod As String) As Boolean
    Dim oCredito As DCredito
    Dim rs As ADODB.Recordset
    Dim oNegCredito As NCredito
    Dim oGastos As nGasto
    Dim dParam As DParametro
    Dim nAnios As Integer
    Dim oAge As DAgencias
    
    'On Error GoTo ErroCargarDatos
     Set oCredito = New DCredito
     Set rs = oCredito.RecuperaDatosCreditoVigente(psCtaCod, gdFecSis)
     Set oCredito = Nothing
     
     Set DParametro = New DParametro
     nAnios = dParam.RecuperaValorParametro(3053)
     Set dParam = Nothing
     
     If Not rs.EOF And Not rs.BOF Then
        If Mid(psCtaCod, 4, 2) <> gsCodAge Then
            Set Age = New DAgencias
            lblAgencias.Caption = oAge.NombreAgencia(Mid(psCtaCod, 4, 2))
            Set oAge = Nothing
        Else
            lblAgencias.Caption = ""
        End If
     End If
     LblMontoCuota.Caption = Format(IIf(IsNull(rs!CuotaAprobada), 0, rs!CuotaAprobada), "###,##0.00")
     bPrepago = IIf(rs!bPrepago = True, 1, 0)
     nCalendDinamTipo = rs!nCalendDinamTipo
     Set oNegCredito = New NCredito
     
     nCalendDinamTipo = rs!nCalendDinamico
     
     If IsNull(rs!nCalendDinamico) Then
        bCalenDinamic = False
     Else
        If rs!nCalendDinamico = 1 Then
            bCalenDinamic = True
        Else
            bCalenDinamic = False
        End If
     End If
     
     If rs!nColocCalend = gColocCalendCodCL Then
        bCalenCuotaLibre = True
     Else
        bCalenCuotaLibre = False
     End If
     
     
End Function

