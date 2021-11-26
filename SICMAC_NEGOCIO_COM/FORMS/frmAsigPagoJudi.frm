VERSION 5.00
Begin VB.Form frmAsigPagoJudi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pago de Cobranza Judicial"
   ClientHeight    =   4650
   ClientLeft      =   1635
   ClientTop       =   1950
   ClientWidth     =   7065
   Icon            =   "frmAsigPagoJudi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCalculos 
      Height          =   2985
      Left            =   7110
      TabIndex        =   35
      Top             =   990
      Visible         =   0   'False
      Width           =   2355
      Begin VB.Frame Frame2 
         Caption         =   "Deuda "
         Height          =   1425
         Left            =   90
         TabIndex        =   45
         Top             =   1530
         Width           =   2085
         Begin VB.Label lblDeuCap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   300
            Left            =   990
            TabIndex        =   53
            Top             =   165
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Capital "
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   52
            Top             =   210
            Width           =   645
         End
         Begin VB.Label lblDeuMor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   990
            TabIndex        =   51
            Top             =   765
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Int. Morat"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   50
            Top             =   810
            Width           =   825
         End
         Begin VB.Label lblDeuCom 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   990
            TabIndex        =   49
            Top             =   465
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Int. Comp"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   48
            Top             =   495
            Width           =   750
         End
         Begin VB.Label lblDeuGas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   990
            TabIndex        =   47
            Top             =   1065
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Gastos"
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   46
            Top             =   1110
            Width           =   705
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pagos "
         Height          =   1425
         Left            =   90
         TabIndex        =   36
         Top             =   90
         Width           =   1995
         Begin VB.Label lblPagCap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   915
            TabIndex        =   44
            Top             =   165
            Width           =   990
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Capital "
            Height          =   195
            Index           =   13
            Left            =   75
            TabIndex        =   43
            Top             =   210
            Width           =   750
         End
         Begin VB.Label lblPagMor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   915
            TabIndex        =   42
            Top             =   765
            Width           =   990
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Int. Morat"
            Height          =   195
            Index           =   14
            Left            =   75
            TabIndex        =   41
            Top             =   810
            Width           =   885
         End
         Begin VB.Label lblPagCom 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   915
            TabIndex        =   40
            Top             =   465
            Width           =   990
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Int. Comp"
            Height          =   195
            Index           =   15
            Left            =   75
            TabIndex        =   39
            Top             =   495
            Width           =   885
         End
         Begin VB.Label lblPagGas 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   915
            TabIndex        =   38
            Top             =   1065
            Width           =   990
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "Gastos"
            Height          =   195
            Index           =   16
            Left            =   75
            TabIndex        =   37
            Top             =   1110
            Width           =   795
         End
      End
   End
   Begin VB.Frame fraListado 
      Caption         =   "Relación  de Cuentas"
      Height          =   945
      Left            =   4500
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   2505
      Begin VB.ListBox lstContratos 
         Height          =   645
         Left            =   90
         TabIndex        =   31
         Top             =   225
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar ..."
      Height          =   405
      Left            =   3240
      TabIndex        =   28
      Top             =   180
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   390
      Left            =   915
      TabIndex        =   4
      Top             =   4140
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2085
      TabIndex        =   5
      Top             =   4140
      Width           =   990
   End
   Begin VB.Frame fradatos 
      Height          =   705
      Left            =   180
      TabIndex        =   19
      Top             =   3240
      Width           =   6705
      Begin VB.TextBox txtNumDoc 
         Height          =   285
         Left            =   2970
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ComboBox cboForPag 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1485
      End
      Begin VB.TextBox txtMonPag 
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
         Left            =   5130
         MaxLength       =   13
         TabIndex        =   3
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label lblNumDoc 
         Caption         =   "Documento "
         Height          =   195
         Left            =   2970
         TabIndex        =   27
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Forma Pago "
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   21
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label7 
         Caption         =   "Monto  "
         Height          =   195
         Index           =   4
         Left            =   4500
         TabIndex        =   20
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame fratitular 
      Caption         =   "Credito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   180
      TabIndex        =   7
      Top             =   990
      Width           =   6675
      Begin VB.Label lblDemanda 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   5670
         TabIndex        =   54
         Top             =   1710
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbogado 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   1440
         TabIndex        =   34
         Top             =   1710
         Width           =   3255
      End
      Begin VB.Label lblCondicion 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Condición"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblTipCom 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   5055
         TabIndex        =   26
         Top             =   1350
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Comisión "
         Height          =   195
         Index           =   18
         Left            =   3780
         TabIndex        =   25
         Top             =   1365
         Width           =   1260
      End
      Begin VB.Label lblTipCob 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   990
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Cobranza "
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   23
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label7 
         Caption         =   "Cliente "
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblMetJudi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   5055
         TabIndex        =   18
         Top             =   990
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Met Liquidación "
         Height          =   195
         Index           =   17
         Left            =   3780
         TabIndex        =   17
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label7 
         Caption         =   "Estudio Juridico"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   2700
         TabIndex        =   15
         Top             =   270
         Width           =   3615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCodPers 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   5055
         TabIndex        =   13
         Top             =   630
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda "
         Height          =   195
         Index           =   11
         Left            =   3780
         TabIndex        =   12
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Ing. a Judicial "
         Height          =   210
         Index           =   10
         Left            =   180
         TabIndex        =   11
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Demanda"
         Height          =   195
         Index           =   9
         Left            =   4860
         TabIndex        =   10
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label lblPreJudi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   2790
         TabIndex        =   9
         Top             =   1890
         Visible         =   0   'False
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFecJudi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   630
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5760
      TabIndex        =   6
      Top             =   4140
      Width           =   1020
   End
   Begin SICMACT.ActxCtaCred txtCodigo 
      Height          =   570
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   1005
      enabledage      =   -1  'True
      enabledprod     =   -1  'True
      enabled         =   -1  'True
      Caption         =   "  Credito "
   End
   Begin VB.Image imgInterconex2 
      Height          =   480
      Left            =   5760
      Picture         =   "frmAsigPagoJudi.frx":030A
      Top             =   270
      Width           =   480
   End
   Begin VB.Image imgInterconex 
      Height          =   480
      Left            =   6300
      Picture         =   "frmAsigPagoJudi.frx":0614
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Crédito Antiguo"
      Height          =   285
      Left            =   180
      TabIndex        =   29
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmAsigPagoJudi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Centralizacion
'usuario : NSSE
'Fecha : 14/12/2000

'****************************************************************************
'* ARCHIVO : frmAsigPagoJudi
'* Registrar Pagos de Cobranza Judicial por Ventanilla, (Recepcion de Dinero)
'****************************************************************************
Option Explicit
Dim pAgSede As String * 2
Dim pDias As Integer
Dim lbPoseeGastos As Boolean
Dim lbCambioEstado As Boolean
Dim lsEstado As String
Dim lnMonto As Currency
Dim lbVacio As Boolean
Dim lsCodCta As String
Dim lsMoneda As String
Dim lsMonto As String
Dim lnMontoPagado As Currency
Dim lbVacioCred As Boolean
Dim vGAS As Currency, vCAP As Currency, vCOM As Currency, vMOR As Currency
Dim vIntComGenerado As Currency
Dim vComision As Currency
Dim vSobDisPag As Currency
Dim EstJud As String
Dim lsRefinan As String
' Para Codigos Operacion
Dim lsCodOpeKARDEX As String
Dim lsCodOpeCAP As String
Dim lsCodOpeICOM As String
Dim lsCodOpeIMOR As String
Dim lsCodOpeGAST As String
Dim lsCodOpeCOMI As String
Dim pConexion As New ADODB.Connection
Public PagoJudInterconex As String  ' S = Con Interconex / N = Sin Intercone
Public PagoJudOtraCMAC As String   ' S = Con Otra Cmac  / N = Sin OtraCmac
Public vCodAgeRem As String ' Para Operacion con otra Ag / Cmac
Public vCodUsuRem As String ' Para Operacion con otra Ag / Cmac

Private Sub cboForPag_Click()
txtNumDoc.Visible = IIf(Right(cboForPag, 2) = "EF", False, True)
LblNumDoc.Visible = IIf(Right(cboForPag, 2) = "EF", False, True)
End Sub

Private Sub cboForPag_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If txtNumDoc.Visible = True Then
        txtNumDoc.SetFocus
    Else
        TxtMonPag.SetFocus
    End If
 End If
End Sub

'Permite buscar un cliente por su nombre y/o documento
Private Sub cmdBuscar_Click()
On Error GoTo ControlError
Dim RegCJ As New ADODB.Recordset
Dim SSQL As String
Dim vCodigo As String
frmBuscaCli.Inicia frmAsigGastJudi, False
vCodigo = CodGrid
If Trim(vCodigo) <> "" Then
    SSQL = " SELECT PJ.ccodcta, CJ.cestado, R.ccodant " & _
        " FROM PersCJudi PJ INNER JOIN CredCJudi CJ ON pj.ccodcta = cj.ccodcta " & _
        " LEFT JOIN RelCtaMigJudi R ON CJ.ccodcta = R.ccodcta " & _
        " WHERE PJ.cCodpers = '" & vCodigo & "' AND CJ.cEstado IN('V') " & _
        " ORDER BY PJ.ccodcta"
       If Right(Trim(gsCodAge), 2) = pAgSede Then
          RegCJ.Open SSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
       Else
          If AbreConeccion(pAgSede) Then
             RegCJ.Open SSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
          End If
       End If
        
    If (RegCJ.BOF Or RegCJ.EOF) Then
        MsgBox " Cliente no posee Contratos ", vbInformation, " Aviso "
    Else
        lstContratos.Clear
        fraListado.Visible = True
        lstContratos.Visible = True
        Do While Not RegCJ.EOF
            lstContratos.AddItem RegCJ!cCodCta & " - " & RegCJ!cCodAnt
            RegCJ.MoveNext
        Loop
        lstContratos.SetFocus
    End If
    RegCJ.Close
    Set RegCJ = Nothing
Else
    MsgBox " Cliente no selecionado ", vbInformation, " Aviso "
    lstContratos.Clear
    txtCodigo.Text = Right(gsCodAge, 2)
    Limpiar
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdcancelar_Click()
    Limpiar
    txtCodigo.Text = Right(gsCodAge, 2)
    fraListado.Visible = False
    lstContratos.Clear
    cboForPag.Enabled = False
    TxtMonPag.Enabled = False
    cmdGrabar.Enabled = False
    txtCodigo.Enabled = True
    txtCodigo.Enfoque (2)
End Sub

Private Sub CmdGrabar_Click()
Dim rsPlan As New ADODB.Recordset
Dim SSQL As String
Dim vNumTran As Double
Dim vNumGasCta As Double
Dim lsCodGas As String
Dim lsComenta As String
'Dim lsfecha As String
Dim Item As ListItem
Dim lbAbrioTransaccion As Boolean
Dim VPago As Currency
Dim vSalGasAnt As Currency
Dim vSalGasNue As Currency
Dim lsCodOpe As String
Dim lnRefDemanda As String  ' Refinan - Demanda
Dim vAgenciaGraba As String

On Error GoTo GrabarPagJudLinea

Call CodigosOpe ' Codigos de Operacion

If vCodAgeRem = "" Then
   vAgenciaGraba = gsCodAge
Else
   vAgenciaGraba = vCodAgeRem
End If
' (1) Con Demand/Norm  (2) Sin Demand/Norm (3) Con Demand/Ref  (4) Sin Demand/Ref
lnRefDemanda = IIf(lsRefinan = "N", IIf(lblDemanda = "S", 1, 2), IIf(lblDemanda = "S", 3, 4))

If Validar = True Then
   If MsgBox("Desea Grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        'Carga fecha y hora de grabación
        gdHoraGrab = Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss")
        txtNumDoc.Text = Replace(txtNumDoc.Text, "'", " ", , , vbTextCompare)
        If Right(Trim(gsCodAge), 2) = pAgSede Then
          ' Si Agencia local / Recepciona LLamada Agencia / Recepciona Llamada CMAC
            'CONEXION LOCAL
            vNumTran = NumUltJudi(lsCodCta, dbCmact) + 1
            vNumGasCta = NumGasCta(lsCodCta)
            dbCmact.BeginTrans  ' INICIA TRANSACCION - LOCAL
            lbAbrioTransaccion = True
            'Actualiza en CredCJudi
            SSQL = "UPDATE CredCJudi SET " & _
                " nSaldCap =" & CCur(lblDeuCap) & ", nSaldIntCom = " & CCur(lblDeuCom) & ", " & _
                " nSaldIntMor = " & CCur(lblDeuMor) & ", nSaldGast = " & CCur(lblDeuGas) & ", " & _
                " nCapPag =" & CCur(lblPagCap) & ", nIntComPag = " & CCur(lblPagCom) & ", " & _
                " nIntMorPag = " & CCur(lblPagMor) & ", nGastoPag = " & CCur(lblPagGas) & ", " & _
                " cCodUsu = '" & gsCodUser & "',  " & _
                " dFecMod = '" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "'," & _
                " dFecUltPago = '" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "'," & _
                " nIntComGen = nIntComGen + " & vIntComGenerado & "," & _
                " nNumTranCta = " & vNumTran & _
                " WHERE cCodCta = '" & Trim(lsCodCta) & "' "
            dbCmact.Execute SSQL
            'Actualiza en KardexJudi
            SSQL = "INSERT INTO KardexJudi (cCodCta, dFecTran, " & _
                " nNumTranCta, nMonTran, nCapital, nIntComp, nIntMorat, nGastos, " & _
                " cCodOpe, cNumDoc, cCodAge, cCodUsu, " & _
                " cMetLiquid, cTipTrans, cModPag) " & _
                " VALUES ('" & Trim(lsCodCta) & "','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "', " & _
                vNumTran & "," & CCur(TxtMonPag) & "," & vCAP & "," & vCOM & "," & vMOR & "," & vGAS & ",'" & _
                lsCodOpeKARDEX & "','" & txtNumDoc.Text & "','" & vAgenciaGraba & "','" & gsCodUser & "','" & _
                Trim(lblMetJudi) & "','1','" & IIf(Right(cboForPag, 2) = "EF", "EF", "CH") & "')"
            dbCmact.Execute SSQL
            'Inserta Gasto Pagado por Comision de Abogado
            If vComision > 0 Then
              SSQL = "INSERT INTO PLANGASTOSJUD(cCodCta, cCodGas, cOrigGas, " & _
                " nMonGas, nMonNeg,cEstado, cComenta, " & _
                " cMotivoGas, dFecAsig, dFecMod, " & _
                " cCodUsu, nNumTranCta, nNumGasCta) " & _
                " VALUES('" & lsCodCta & "','01807000','D'," & vComision & _
                "," & vComision & ",'P','..'," & _
                "'COMISION','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & _
                gsCodUser & "'," & vNumTran & "," & vNumGasCta & ")"
              dbCmact.Execute SSQL
              'Graba en TranDiaria Comision
              If vGAS >= vComision Then  ' Si Pago afecta la Comision
                SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                   " cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                   " nNroCuota, cCodUsuRem ) " & _
                   "VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeCOMI & "','" & Trim(lsCodCta) & "','" & _
                   txtNumDoc.Text & "'," & vComision & ",'" & gsCodUser & "','" & vAgenciaGraba & "'," & _
                   vNumTran & ",'" & vCodUsuRem & "')"
                dbCmact.Execute SSQL
              End If
            End If
            
            If (vGAS - vComision) > 0 Then   '  Gastos
            SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                " cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                " nNroCuota, cCodUsuRem ) " & _
                "VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeGAST & "','" & Trim(lsCodCta) & "','" & _
                txtNumDoc.Text & "'," & (vGAS - vComision) & ",'" & gsCodUser & "','" & vAgenciaGraba & "'," & _
                vNumTran & ",'" & vCodUsuRem & "')"
            dbCmact.Execute SSQL
            End If
            
            If vCAP > 0 Then   ' Capital
                SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                    " cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                    " nNroCuota,nSaldCnt,cCodUsuRem) " & _
                    " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeCAP & "','" & Trim(lsCodCta) & "','" & _
                    txtNumDoc.Text & "'," & vCAP & ",'" & gsCodUser & "','" & vAgenciaGraba & "'," & _
                    vNumTran & "," & lnRefDemanda & ",'" & vCodUsuRem & "')"
                dbCmact.Execute SSQL
            End If
            
          
            If vCOM > 0 Then  ' INT COMP
                SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                    " cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                    " nNroCuota,nSaldCnt,cCodUsuRem) " & _
                    " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeICOM & "','" & Trim(lsCodCta) & "','" & _
                    txtNumDoc.Text & "'," & vCOM & ",'" & gsCodUser & "','" & vAgenciaGraba & "'," & _
                    vNumTran & "," & lnRefDemanda & ",'" & vCodUsuRem & "')"
                dbCmact.Execute SSQL
            End If
            
            If vMOR > 0 Then  ' INT MOR
                SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                    " cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                    " nNroCuota,nSaldCnt,cCodUsuRem) " & _
                    "VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeIMOR & "','" & Trim(lsCodCta) & "','" & _
                    txtNumDoc.Text & "'," & vMOR & ",'" & gsCodUser & "','" & vAgenciaGraba & "'," & _
                    vNumTran & "," & lnRefDemanda & ",'" & vCodUsuRem & "')"
                dbCmact.Execute SSQL
            End If
            
            If vSobDisPag > 0 Then
                'Inserta Sobrante como Gasto Administrativo
                SSQL = "INSERT INTO PLANGASTOSJUD(cCodCta, cCodGas, cOrigGas, " & _
                    " nMonGas, nMonNeg, dFecPag, nMonPag, cEstado, cComenta, " & _
                    " cMotivoGas, dFecAsig, dFecMod, " & _
                    " cCodUsu, nNumTranCta, nNumGasCta) " & _
                    " VALUES('" & lsCodCta & "','01808001','D'," & vSobDisPag & _
                    "," & vSobDisPag & ",'" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "'," _
                    & vSobDisPag & ",'P','..'," & _
                    "'Gasto Administ','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" _
                    & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & _
                    gsCodUser & "'," & vNumTran & "," & vNumGasCta + 1 & ")"
                dbCmact.Execute SSQL
                SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                    " cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                    " nNroCuota, cCodUsuRem) " & _
                    "VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeGAST & "','" & Trim(lsCodCta) & "','" & _
                    txtNumDoc.Text & "'," & vSobDisPag & ",'" & gsCodUser & "','" & vAgenciaGraba & "'," & _
                    vNumTran & ",'" & vCodUsuRem & "')"
                dbCmact.Execute SSQL
                
            End If
            
            dbCmact.CommitTrans  ' CIERRA TRANSACCION - LOCAL
            
            ImprimirReciboPago
            Do While True
                If MsgBox("Desea Reimprimir Boleta ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    ImprimirReciboPago
                Else
                    Exit Do
                End If
            Loop

        Else ' **************************************************************************
            'CONEXION REMOTA
            If AbreConeccion(pAgSede, True, True) Then
                vNumTran = NumUltJudi(lsCodCta, dbCmactN) + 1
                dbCmactN.BeginTrans  ' INICIA TRANSACCION - REMOTA
                lbAbrioTransaccion = True
                'Actualiza en CredCJudi - REMOTO
                SSQL = "UPDATE CredCJudi SET " & _
                    " nSaldCap =" & Val(lblDeuCap) & ", nSaldIntCom = " & Val(lblDeuCom) & ", " & _
                    " nSaldIntMor = " & Val(lblDeuMor) & ", nSaldGast = " & Val(lblDeuGas) & ", " & _
                    " nCapPag =" & Val(lblPagCap) & ", nIntComPag = " & Val(lblPagCom) & ", " & _
                    " nIntMorPag = " & Val(lblPagMor) & ", nGastoPag = " & Val(lblPagGas) & ", " & _
                    " cCodUsu = '" & gsCodUser & "', " & _
                    " dFecMod = '" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "'," & _
                    " dFecUltPago = '" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "'," & _
                    " nIntComGen = nIntComGen  + " & vIntComGenerado & "," & _
                    " nNumTranCta = " & vNumTran & _
                    " WHERE cCodCta = '" & Trim(lsCodCta) & "' "
                dbCmactN.Execute SSQL
                'Actualiza en KardexJudi - REMOTO
                SSQL = "INSERT INTO KardexJudi (cCodCta, dFecTran, nNumTranCta," & _
                    " nMonTran, nCapital, nIntComp, nIntMorat, nGastos, " & _
                    " cCodOpe, cNumDoc, cCodAge, cCodUsu, " & _
                    " cMetLiquid, cTipTrans, cModPag) " & _
                    " VALUES ('" & Trim(lsCodCta) & "','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "', " & _
                    vNumTran & "," & CCur(TxtMonPag) & "," & vCAP & "," & vCOM & "," & vMOR & "," & vGAS & ",'" & _
                    lsCodOpeKARDEX & "','" & txtNumDoc.Text & "','" & gsCodAge & "','" & gsCodUser & "','" & _
                    Trim(lblMetJudi) & "','1','" & IIf(Right(cboForPag, 2) = "EF", "EF", "CH") & "')"
               dbCmactN.Execute SSQL
                'Inserta Gasto pagado por Comision de Abogado - REMOTO
                If vComision > 0 Then
                   SSQL = "INSERT INTO PLANGASTOSJUD(cCodCta, cCodGas, cOrigGas, " & _
                     " nMonGas, nMonNeg, cEstado, cComenta, " & _
                     " cMotivoGas, dFecAsig, dFecMod,  " & _
                     " cCodUsu, nNumTranCta, nNumGasCta) " & _
                     " VALUES('" & lsCodCta & "','01807000','D'," & vComision & _
                     "," & vComision & ",'P','..'," & _
                     "'COMISION','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & " ','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & _
                     gsCodUser & "'," & vNumTran & "," & vNumGasCta & ")"
                   dbCmactN.Execute SSQL
                   'Graba en TranDiaria Comision  - REMOTO
                   If vGAS >= vComision Then ' Si Pago afecta la comision
                      SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                        " cNumDoc, nMonTran, cCodUsuRem, cCodAge, " & _
                        " nNroCuota) " & _
                        "VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeCOMI & "','" & Trim(lsCodCta) & "','" & _
                        txtNumDoc.Text & "'," & vComision & ",'" & gsCodUser & "','" & gsCodAge & "'," & _
                        vNumTran & ")"
                      dbCmactN.Execute SSQL
                   End If
                End If
                
                If (vGAS - vComision) > 0 Then   '  Gastos - REMOTO
                  SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                    " cNumDoc, nMonTran, cCodUsuRem, cCodAge, " & _
                    " nNroCuota) " & _
                    " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeGAST & "','" & Trim(lsCodCta) & "','" & _
                    txtNumDoc.Text & "'," & (vGAS - vComision) & ",'" & gsCodUser & "','" & gsCodAge & "'," & _
                    vNumTran & ")"
                  dbCmactN.Execute SSQL
                End If
                
                If vCAP > 0 Then   ' Capital - REMOTO
                    SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                        " cNumDoc, nMonTran, cCodUsuRem, cCodAge, " & _
                        " nNroCuota,nSaldCnt) " & _
                        " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeCAP & "','" & Trim(lsCodCta) & "','" & _
                        txtNumDoc.Text & "'," & vCAP & ",'" & gsCodUser & "','" & gsCodAge & "'," & _
                        vNumTran & "," & lnRefDemanda & ")"
                    dbCmactN.Execute SSQL
                End If
                
                If vCOM > 0 Then ' Int Comp  - REMOTO
                    SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                        " cNumDoc, nMonTran, cCodUsuRem, cCodAge, " & _
                        " nNroCuota,nSaldCnt) " & _
                        " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeICOM & "','" & Trim(lsCodCta) & "','" & _
                        txtNumDoc.Text & "'," & vCOM & ",'" & gsCodUser & "','" & gsCodAge & "'," & _
                        vNumTran & "," & lnRefDemanda & ")"
                    dbCmactN.Execute SSQL
                End If
                
                If vMOR > 0 Then  ' Int Mor - REMOTO
                    SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                        " cNumDoc, nMonTran, cCodUsuRem, cCodAge, " & _
                        " nNroCuota,nSaldCnt) " & _
                        " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeIMOR & "','" & Trim(lsCodCta) & "','" & _
                        txtNumDoc.Text & "'," & vMOR & ",'" & gsCodUser & "','" & gsCodAge & "'," & _
                        vNumTran & "," & lnRefDemanda & ")"
                    dbCmactN.Execute SSQL
                End If
                
                If vSobDisPag > 0 Then
                    'Inserta Sobrante como Gasto Administrativo -  REMOTO
                   SSQL = "INSERT INTO PLANGASTOSJUD(cCodCta, cCodGas, cOrigGas, " & _
                     " nMonGas, nMonNeg, dFecPag, nMonPag, cEstado, cComenta, " & _
                     " cMotivoGas, dFecAsig, dFecMod, " & _
                     " cCodUsu, nNumTranCta, nNumGasCta) " & _
                     " VALUES('" & lsCodCta & "','01808001','D'," & vSobDisPag & _
                     "," & vSobDisPag & ",'" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "'," _
                     & vSobDisPag & ",'P','..'," & _
                     "'Gasto Administrat','" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" _
                     & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & _
                     gsCodUser & "'," & vNumTran & "," & vNumGasCta + 1 & ")"
                  dbCmactN.Execute SSQL
                   SSQL = "INSERT INTO TranDiaria (dFecTran, cCodOpe, cCodCta, " & _
                        " cNumDoc, nMonTran, cCodUsuRem, cCodAge, " & _
                        " nNroCuota) " & _
                        "VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & lsCodOpeGAST & "','" & Trim(lsCodCta) & "','" & _
                        txtNumDoc.Text & "'," & vSobDisPag & ",'" & gsCodUser & "','" & gsCodAge & "'," & _
                        vNumTran & ")"
                    dbCmactN.Execute SSQL
                End If
               
                dbCmactN.CommitTrans  ' CIERRA TRANSACCION - REMOTA
                lbAbrioTransaccion = False
    
                'Grabar en Trandiaria - LOCAL  ANTES
                SSQL = "INSERT INTO TranDiaria ( dFectran, cCodCta, cCodOpe, " & _
                    "cNumDoc, nMonTran, cCodUsu, cCodAge, " & _
                    " nNroCuota) " & _
                    " VALUES ('" & Format(gdHoraGrab, "mm/dd/yyyy hh:mm:ss") & "','" & Trim(lsCodCta) & "','" & gsPagJudEfDOA & "','" & _
                    txtNumDoc.Text & "'," & CCur(TxtMonPag) & ",'" & gsCodUser & "','112" & pAgSede & "'," & _
                    vNumTran & ")"
                dbCmact.Execute SSQL
                
                ImprimirReciboPago
                Do While True
                    If MsgBox("Desea Reimprimir Boleta ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                        ImprimirReciboPago
                    Else
                        Exit Do
                    End If
                Loop
                
            End If  ' Si hay conexion con Sede
            
        
        End If   ' Fin Local /Remoto
        
        
        
        If lbEstConexion Then
            CierraConeccion
        End If
        
        cmdGrabar.Enabled = False
        txtCodigo.Enabled = True
        'cmdExaminar.Enabled = True
        Limpiar
        txtCodigo.Text = Right(gsCodAge, 2)
   Else
        MsgBox "Grabación cancelada", vbInformation, " Aviso "
   End If  ' Desea Grabar
   
End If  ' Validar
Exit Sub
GrabarPagJudLinea:
   MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    If lbAbrioTransaccion Then
        dbCmact.RollbackTrans
    End If
    If dbCmactN.BeginTrans = True Then
       dbCmactN.RollbackTrans
    End If
    'MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    Exit Sub
End Sub

Private Sub CodigosOpe()
 lsCodOpeKARDEX = ""
 lsCodOpeCAP = ""
 lsCodOpeICOM = ""
 lsCodOpeIMOR = ""
 lsCodOpeGAST = ""
 lsCodOpeCOMI = ""
 If PagoJudOtraCMAC = "N" Then   ' Si es de la Caja de trujillo
    If PagoJudInterconex = "S" Then ' Con Interconexion
       If Right(Trim(gsCodAge), 2) = pAgSede Then  ' *********En la Sede
          If Right(cboForPag, 2) = "EF" Then  ' Efectiv
             
             If EstJud = "J" Then
                lsCodOpeKARDEX = gsPagJudEfecLin
                lsCodOpeCAP = gsPagJudEfKLin
                lsCodOpeICOM = gsPagJudEfICLin
                lsCodOpeIMOR = gsPagJudEfIMLin
                lsCodOpeGAST = gsPagJudGasELin
                lsCodOpeCOMI = gsPagJudGasComiELin
             ElseIf EstJud = "A" Then
                lsCodOpeKARDEX = gsPagCasEfecLin
                lsCodOpeCAP = gsPagCasEfKLin
                lsCodOpeICOM = gsPagCasEfICLin
                lsCodOpeIMOR = gsPagCasEfIMLin
                lsCodOpeGAST = gsPagJudGasELin
                lsCodOpeCOMI = gsPagJudGasComiELin
             End If
          
          ElseIf Right(cboForPag, 2) = "CH" Then  ' Cheque
             
             If EstJud = "J" Then
                lsCodOpeKARDEX = gsPagJudChequeLin
                lsCodOpeCAP = gsPagJudChKLin
                lsCodOpeICOM = gsPagJudChICLin
                lsCodOpeIMOR = gsPagJudChIMLin
                lsCodOpeGAST = gsPagJudGasCLin
                lsCodOpeCOMI = gsPagJudGasComiCLin
             ElseIf EstJud = "A" Then
                lsCodOpeKARDEX = gsPagCasChequeLin
                lsCodOpeCAP = gsPagCasChKLin
                lsCodOpeICOM = gsPagCasChICLin
                lsCodOpeIMOR = gsPagCasChIMLin
                lsCodOpeGAST = gsPagJudGasCLin
                lsCodOpeCOMI = gsPagJudGasComiCLin
             End If
          End If
       
       Else  ' *************************************** En Otra Agencia
          
          If Right(cboForPag, 2) = "EF" Then  ' Efectiv
             
             If EstJud = "J" Then
                lsCodOpeKARDEX = gsPagJudEfEOALin
                lsCodOpeCAP = gsPagJudEfKEOALin
                lsCodOpeICOM = gsPagJudEfICEOALin
                lsCodOpeIMOR = gsPagJudEfIMEOALin
                lsCodOpeGAST = gsPagJudGasEEOALin
                lsCodOpeCOMI = gsPagJudGasComiEEOALin
             ElseIf EstJud = "A" Then
                lsCodOpeKARDEX = gsPagCasEfEOALin
                lsCodOpeCAP = gsPagCasEfKEOALin
                lsCodOpeICOM = gsPagCasEfICEOALin
                lsCodOpeIMOR = gsPagCasEfIMEOALin
                lsCodOpeGAST = gsPagJudGasEEOALin
                lsCodOpeCOMI = gsPagJudGasComiEEOALin
             End If
          ElseIf Right(cboForPag, 2) = "CH" Then  ' Cheque
             
             If EstJud = "J" Then
                lsCodOpeKARDEX = gsPagJudChEOALin
                lsCodOpeCAP = gsPagJudChKEOALin
                lsCodOpeICOM = gsPagJudChICEOALin
                lsCodOpeIMOR = gsPagJudChIMEOALin
                lsCodOpeGAST = gsPagJudGasCEOALin
                lsCodOpeCOMI = gsPagJudGasComiCEOALin
             ElseIf EstJud = "A" Then
                lsCodOpeKARDEX = gsPagCasChEOALin
                lsCodOpeCAP = gsPagCasChKEOALin
                lsCodOpeICOM = gsPagCasChICEOALin
                lsCodOpeIMOR = gsPagCasChIMEOALin
                lsCodOpeGAST = gsPagJudGasCEOALin
                lsCodOpeCOMI = gsPagJudGasComiCEOALin
             End If
          End If
       
       End If
    Else ' Sin Interconexion ******************
    
          If Right(cboForPag, 2) = "EF" Then  ' Efectiv
             
             If EstJud = "J" Then
                lsCodOpeKARDEX = gsPagJudEfEOALin
                lsCodOpeCAP = gsPagJudEfKEOALin
                lsCodOpeICOM = gsPagJudEfICEOALin
                lsCodOpeIMOR = gsPagJudEfIMEOALin
                lsCodOpeGAST = gsPagJudGasEEOALin
                lsCodOpeCOMI = gsPagJudGasComiEEOALin
             ElseIf EstJud = "A" Then
                lsCodOpeKARDEX = gsPagCasEfEOALin
                lsCodOpeCAP = gsPagCasEfKEOALin
                lsCodOpeICOM = gsPagCasEfICEOALin
                lsCodOpeIMOR = gsPagCasEfIMEOALin
                lsCodOpeGAST = gsPagJudGasEEOALin
                lsCodOpeCOMI = gsPagJudGasComiEEOALin
             End If
          ElseIf Right(cboForPag, 2) = "CH" Then  ' Cheque
             
             If EstJud = "J" Then
                lsCodOpeKARDEX = gsPagJudChEOALin
                lsCodOpeCAP = gsPagJudChKEOALin
                lsCodOpeICOM = gsPagJudChICEOALin
                lsCodOpeIMOR = gsPagJudChIMEOALin
                lsCodOpeGAST = gsPagJudGasCEOALin
                lsCodOpeCOMI = gsPagJudGasComiCEOALin
             ElseIf EstJud = "A" Then
                lsCodOpeKARDEX = gsPagCasChEOALin
                lsCodOpeCAP = gsPagCasChKEOALin
                lsCodOpeICOM = gsPagCasChICEOALin
                lsCodOpeIMOR = gsPagCasChIMEOALin
                lsCodOpeGAST = gsPagJudGasCEOALin
                lsCodOpeCOMI = gsPagJudGasComiCEOALin
             End If
          End If
    
    End If ' Con Interconexion
 '************************ Para Pago en Otra CMAC *************
 Else   ' Desde  otra Caja
    If Right(cboForPag, 2) = "EF" Then  ' Efectiv
       
       If EstJud = "J" Then
          lsCodOpeKARDEX = gsPagJudEfecEOCLin
          lsCodOpeCAP = gsPagJudEfKEOCLin
          lsCodOpeICOM = gsPagJudEfICEOCLin
          lsCodOpeIMOR = gsPagJudEfIMEOCLin
          lsCodOpeGAST = gsPagJudGasEEOCLin
          lsCodOpeCOMI = gsPagJudGasComiEEOCLin
       ElseIf EstJud = "A" Then
          lsCodOpeKARDEX = gsPagCasEfecEOCLin
          lsCodOpeCAP = gsPagCasEfKEOCLin
          lsCodOpeICOM = gsPagCasEfICEOCLin
          lsCodOpeIMOR = gsPagCasEfIMEOCLin
          lsCodOpeGAST = gsPagJudGasEEOCLin
          lsCodOpeCOMI = gsPagJudGasComiEEOCLin
       End If
    ElseIf Right(cboForPag, 2) = "CH" Then  ' Cheque (SE GRABA COMO CON EFECTIVO)
       
       If EstJud = "J" Then
          lsCodOpeKARDEX = gsPagJudEfecEOCLin
          lsCodOpeCAP = gsPagJudEfKEOCLin
          lsCodOpeICOM = gsPagJudEfICEOCLin
          lsCodOpeIMOR = gsPagJudEfIMEOCLin
          lsCodOpeGAST = gsPagJudGasEEOCLin
          lsCodOpeCOMI = gsPagJudGasComiEEOCLin
       ElseIf EstJud = "A" Then
          lsCodOpeKARDEX = gsPagCasEfecEOCLin
          lsCodOpeCAP = gsPagCasEfKEOCLin
          lsCodOpeICOM = gsPagCasEfICEOCLin
          lsCodOpeIMOR = gsPagCasEfIMEOCLin
          lsCodOpeGAST = gsPagJudGasEEOCLin
          lsCodOpeCOMI = gsPagJudGasComiEEOCLin
       End If
    End If
 
 End If
End Sub


Private Sub cmdSalir_Click()
    If cmdGrabar.Enabled = True Then
        MsgBox "Antes de Culminar Guarde o Cancele los Cambios", vbInformation, "Aviso"
        cmdGrabar.SetFocus
    Else
        If lbEstConexion Then CierraConeccion
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    cmdSalir.Value = True
End If
End Sub

'carga formulario de busqueda de contrato antiguo
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then  'And Len(Trim(vNroContrato)) <> 12
        gsNueCod = ""
        frmRelaCredJudi.Show 1
        If Len(Trim(gsNueCod)) = 12 Then
            fraListado.Visible = False
            lstContratos.Clear
            Limpiar
            txtCodigo.Text = Right(gsCodAge, 2)
            txtCodigo.Enabled = True
            txtCodigo.Text = gsNueCod
            txtCodigo.Enfoque (3)
            gsNueCod = ""
        Else
            MsgBox " Búsqueda cancelada ", vbInformation, " ! Aviso ! "
        End If
    End If
End Sub

Private Sub Limpiar()
    'txtCodigo.Text = Right(gsCodAge, 2)
    lblCodPers = "":    lblNomPers = ""
    lblFecJudi = "":    lblPreJudi = ""
    lblMetJudi = "":    lblPreJudi = ""
    lblMoneda = "":    lblTipCob = ""
    lblTipCom = ""
    lblDeuCap = "":    lblDeuCom = ""
    lblDeuMor = "":    lblDeuGas = ""
    lblPagCap = "":    lblPagCom = ""
    lblPagMor = "":    lblPagGas = ""
    lblDemanda = "":   lblCondicion = ""
    lblAbogado = ""
    TxtMonPag = ""
    txtNumDoc = ""
    'lblDeudaTotal = ""
    cboForPag.ListIndex = 0
End Sub

'******************************************************************
'funcion para validar la información ingresada el momento de grabar
'******************************************************************
Private Function Validar() As Boolean
   If lsCodCta = 0 Then
        Validar = False
        MsgBox "Número de Cuenta de Crédito no valido", vbInformation, "Aviso"
        txtCodigo.Enfoque 2
        Exit Function
   Else
        Validar = True
   End If
   If Len(TxtMonPag) = 0 Then
        Validar = False
        MsgBox "Monto no válido para el plan", vbInformation, "Aviso"
        TxtMonPag.SetFocus
        Exit Function
   Else
        Validar = True
   End If

End Function

Private Sub Form_Load()
gcIntCentra = CentraSdi(Me)
AbreConexion

CargaParametros
cboForPag.AddItem "Efectivo                EF"
cboForPag.AddItem "Cheque                  CH"
txtCodigo.Text = Right(gsCodAge, 2)
Limpiar
If PagoJudInterconex = "N" Then
   imgInterconex.Visible = True
Else
   imgInterconex.Visible = False
End If

If PagoJudOtraCMAC = "N" Then
   imgInterconex2.Visible = False
Else
   imgInterconex2.Visible = True
End If

If Mid(gsCodAge, 4, 2) = pAgSede Then
   Set pConexion = dbCmact
Else
   If AbreConeccion(pAgSede, True, False) Then
      Set pConexion = dbCmactN
   Else
      MsgBox "No se Pudo Conectar con la Base de Judicial"
      Exit Sub
   End If
End If
End Sub

Private Sub CargaDatosCredCobJud()
Dim rs As New ADODB.Recordset
Dim tmpSql As String
Dim vTipCom As String
Dim lnDiaUltPago As String

 tmpSql = "SELECT P.cCodPers, P.cNomPers, PC.cCodCta, PC.cRelaCta, C.nMontoApr, C.cRefinan, " & _
    " CJ.nSaldCap, cj.nSaldIntCom, cj.nSaldIntMor, cj.nSaldGast, cj.nTasaInt, cj.cEstado, " & _
    " CJ.nCapPag, cj.nIntComPag, cj.nIntMorPag, cj.nGastoPag, CJ.dFecJud, cj.cDemanda, " & _
    " cj.cMetLiq, cj.cTipCJ, cj.ncodcomi, cj.dFecJud , cj.dFecUltPago, cj.cCondicion " & _
    " FROM " & gcCentralPers & "Persona P JOIN PersCJudi PC ON P.cCodPers = PC.cCodPers " & _
    " LEFT JOIN CJCredito C ON pc.ccodcta = c.ccodcta" & _
    " JOIN CredCJudi CJ ON c.ccodcta = cj.ccodcta" & _
    " WHERE PC.cRelaCta = 'TI' AND PC.cCodCta = '" & Trim(txtCodigo.Text) & "'"

 rs.Open tmpSql, pConexion, adOpenStatic, adLockReadOnly, adCmdText
 lbVacioCred = RSVacio(rs)
 If Not lbVacioCred Then
   With rs
    If !cEstado = "L" Then
        MsgBox "CREDITO SE ENCUENTRA CANCELADO", vbInformation, "Aviso"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    lblCodPers.Caption = Trim(!cCodPers)
    lblCodPers.Caption = Trim(!cCodPers)
    lblNomPers.Caption = PstaNombre(Trim(!cNomPers), False)
    lblFecJudi.Caption = Format(!dFecJud, "dd/mm/yyyy")
   
    lblMetJudi.Caption = Trim(!cMetLiq)
    lblMoneda.Caption = IIf(Mid(lsCodCta, 6, 1) = "1", "SOLES", "DOLARES")
    lblTipCob.Caption = IIf(!cTipcj = "J", "Judicial", "Extra Judicial")
    lblCondicion.Caption = IIf(!cCondicion = "J", "Judicial", "Castigado")
    lblAbogado.Caption = MuestraEstudioAbog(!cCodCta, pConexion)
    lblDemanda.Caption = IIf(!cDemanda = "S", "S", "N")
    If Not IsNull(!nCodComi) Then vTipCom = Comision(!nCodComi, pConexion)
    If Len(vTipCom) > 0 Then
        lblTipCom.Caption = Left(vTipCom, 1) & " - " & Mid(vTipCom, 2)
    End If
    ' Calcula Int Comp Generado
    lnDiaUltPago = DateDiff("d", Format(!dFecUltPago, "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
    If lnDiaUltPago > 0 Then
        vIntComGenerado = CalculaIntComJudi(lnDiaUltPago, !nTasaInt, !nSaldCap)
    Else
        vIntComGenerado = 0
    End If
    lblDeuCap.Caption = Format(!nSaldCap, "#0.00")
    lblDeuCom.Caption = Format(!nSaldIntcom + vIntComGenerado, "#0.00")
    lblDeuMor.Caption = Format(!nSaldIntMor, "#0.00")
    lblDeuGas.Caption = Format(!nsaldgast, "#0.00")
    lblPagCap.Caption = Format(!nCapPag, "#0.00")
    lblPagCom.Caption = Format(!nintcompag, "#0.00")
    lblPagMor.Caption = Format(!nIntMorPag, "#0.00")
    lblPagGas.Caption = Format(!nGastoPag, "#0.00")
    EstJud = !cCondicion
    lsRefinan = IIf(!cRefinan = "R", "R", "N")
   End With
   TxtMonPag.ForeColor = pColPriIngreso
   TxtMonPag.BackColor = IIf(lblMoneda.Caption = "SOLES", pColFonSoles, pColFonDolares)
 End If
 rs.Close
 Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  CierraConexion
End Sub

Private Sub fraCalculos_DblClick()
fraCalculos.Visible = False
End Sub

Private Sub Label7_Click(Index As Integer)
fraCalculos.Visible = True
Me.BorderStyle = 3 ' Medible
End Sub

'Permite mostrar los valores del list en los campos del número de contrato
Private Sub lstContratos_Click()
    txtCodigo.Text = Mid(Trim(lstContratos.Text), 1, 12)

    If Len(Trim(txtCodigo.Text)) = 12 Then
         Limpiar
        lbVacio = True
        'txtCodigo.Text = Mid(Trim(lstContratos.Text), 1, 12)
        txtCodigo.Enabled = True
        txtCodigo.Enfoque (3)
        fraListado.Visible = False
    Else
        MsgBox " Contrato no encontrado ", vbInformation, " Aviso "
    End If
End Sub

Private Sub lstContratos_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
    'txtCodigo.Text = Mid(Trim(lstContratos.Text), 1, 12)
    'If Len(txtCodigo.Text) = 12 Then
        'If VerificaCuenta(dbCmact) Then
        '   Limpiar
        '   lbVacio = True
        '   txtCodigo.Enabled = True
        '   txtCodigo.Enfoque (3)
        '   fraListado.Visible = False
        'Else
        '   MsgBox " Contrato no encontrado ", vbInformation, " Aviso "
        'End If
'End If

End Sub

Private Sub txtCodigo_keypressEnter()
  lsMoneda = Mid(Trim(txtCodigo.Text), 6, 1)
  lsCodCta = Trim(txtCodigo.Text)
  
  If lbEstConexion Then CierraConeccion  ' Cierra conexiones abierta
  If Mid(gsCodAge, 4, 2) = pAgSede Then ' Abre conexion con Agencia
     Set pConexion = dbCmact
  Else
     If AbreConeccion(pAgSede, True, False) Then
        Set pConexion = dbCmactN
     Else
        MsgBox "No se Pudo Conectar con la Base de Judicial"
        Exit Sub
     End If
  End If
  CargaDatosCredCobJud
  
  If Len(Trim(lblTipCom)) > 2 Then
        'cmdExaminar.Enabled = False
        txtCodigo.Enabled = False
        cboForPag.Enabled = True
        TxtMonPag.Enabled = True
        txtNumDoc.Enabled = True
        cboForPag.SetFocus
  Else
        MsgBox "Comisión no reconocida", vbInformation, " Aviso "
  End If
  If lbVacioCred Then
        txtCodigo.Enabled = True
        MsgBox "Código de Crédito no Existe ", vbInformation, "Aviso"
        txtCodigo.Text = Mid(Trim(gsCodAge), 4, 2)
        txtCodigo.SetFocus
  End If
  If lbEstConexion Then CierraConeccion
End Sub

'Valida campo txtMonPag
Private Sub TxtMonPag_GotFocus()
fEnfoque TxtMonPag
End Sub
Private Sub TxtMonPag_KeyPress(KeyAscii As Integer)
KeyAscii = intfNumDec(TxtMonPag, KeyAscii, 13, 2)
If KeyAscii = 13 And Len(TxtMonPag) > 0 Then
    If Right(cboForPag, 2) <> "EF" Then
        If Len(txtNumDoc) > 0 Then
            txtNumDoc.Enabled = False
            cboForPag.Enabled = False
            TxtMonPag.Enabled = False
            DistPago
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        End If
    Else
        cboForPag.Enabled = False
        TxtMonPag.Enabled = False
        DistPago
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End If
End Sub
Private Sub TxtMonPag_LostFocus()
    TxtMonPag = Format(TxtMonPag, "#0.00")
End Sub

'Distribuye el pago del cliente entre los diferentes rubros
Private Sub DistPago()
Dim x As Integer
'Inicializa Variables para la Grabación
vCAP = 0: vCOM = 0: vMOR = 0: vGAS = 0
vComision = 0
If Left(lblTipCom, 1) = "P" Then
    'vDeudaAnt = CCur(lblDeuGas)
    vComision = Round((Val(Mid(lblTipCom, 5)) / 100 * Val(TxtMonPag)) / (1 + Val(Mid(lblTipCom, 5)) / 100), 2)
    lblDeuGas = Round(Val(lblDeuGas) + vComision, 2)
ElseIf Left(lblTipCom, 1) = "M" Then
    MsgBox "Tipo de Pago por Moneda", vbInformation, " Aviso "
Else
    MsgBox "Tipo de Pago no reconocido", vbInformation, " Aviso "
End If
vSobDisPag = Val(TxtMonPag)
For x = 1 To Len(lblMetJudi)
    Select Case Mid(lblMetJudi, x, 1)
        Case "G"
            If Val(vSobDisPag) > Val(lblDeuGas) Then
                lblPagGas = Val(lblPagGas) + Val(lblDeuGas)
                vSobDisPag = Val(vSobDisPag) - Val(lblDeuGas)
                vGAS = Val(lblDeuGas)
                lblDeuGas = "0.00"
            Else
                lblDeuGas = Val(lblDeuGas) - Val(vSobDisPag)
                lblPagGas = Val(lblPagGas) + Val(vSobDisPag)
                vGAS = vSobDisPag
                vSobDisPag = 0
            End If
        Case "C"
            If Val(vSobDisPag) > Val(lblDeuCap) Then
                lblPagCap = Val(lblPagCap) + Val(lblDeuCap)
                vSobDisPag = Val(vSobDisPag) - Val(lblDeuCap)
                vCAP = Val(lblDeuCap)
                lblDeuCap = "0.00"
            Else
                lblDeuCap = Val(lblDeuCap) - Val(vSobDisPag)
                lblPagCap = Val(lblPagCap) + Val(vSobDisPag)
                vCAP = vSobDisPag
                vSobDisPag = 0
            End If
        Case "I"
            If Val(vSobDisPag) > Val(lblDeuCom) Then
                lblPagCom = Val(lblPagCom) + Val(lblDeuCom)
                vSobDisPag = Val(vSobDisPag) - Val(lblDeuCom)
                vCOM = Val(lblDeuCom)
                lblDeuCom = "0.00"
            Else
                lblDeuCom = Val(lblDeuCom) - Val(vSobDisPag)
                lblPagCom = Val(lblPagCom) + Val(vSobDisPag)
                vCOM = vSobDisPag
                vSobDisPag = 0
            End If
        Case "M"
            If Val(vSobDisPag) > Val(lblDeuMor) Then
                lblPagMor = Val(lblPagMor) + Val(lblDeuMor)
                vSobDisPag = Val(vSobDisPag) - Val(lblDeuMor)
                vMOR = Val(lblDeuMor)
                lblDeuMor = "0.00"
            Else
                lblDeuMor = Val(lblDeuMor) - Val(vSobDisPag)
                lblPagMor = Val(lblPagMor) + Val(vSobDisPag)
                vMOR = vSobDisPag
                vSobDisPag = 0
            End If
        Case Else
            MsgBox "Método de Pago no reconocido", vbInformation, " Aviso "
    End Select
    If vSobDisPag <= 0 Then Exit For
Next x
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(txtNumDoc) > 0 Then TxtMonPag.SetFocus
End Sub

Private Sub CargaParametros()
On Error GoTo ControlError
    pAgSede = Right(ReadVarSis("JUD", "cAgeJudicial"), 2)
    'pDias = DateDiff("d", Format(ReadVarSis("JUD", "dFecCierreMesJu"), "dd/mm/yyyy"), Format(gdFecSis, "dd/mm/yyyy"))
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Procedimiento de impresión del recibo de rescate del contrato
Private Sub ImprimirReciboPago()
    Dim vEspacio As Integer
    MousePointer = 11
    vEspacio = 9
    ImpreBegChe False, 22
        Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, Tab(4); "COBRANZA JUDICIAL" & Space(24 + vEspacio) & "COBRANZA JUDICIAL"
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
        Print #ArcSal, Tab(4); ImpreFormat(gsNomAge, 25, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & Space(vEspacio) & ImpreFormat(gsNomAge, 25, 0) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm")
        Print #ArcSal, " "
        Print #ArcSal, Tab(4); "CLIENTE : " & ImpreFormat(ImpreCarEsp(lblNomPers.Caption), 30, 0) & Space(vEspacio) & "CLIENTE : " & ImpreFormat(ImpreCarEsp(lblNomPers.Caption), 30, 0) & Chr(10)
        Print #ArcSal, Tab(4); "CREDITO : " & ImpreFormat(lsCodCta, 15) & IIf(Mid(lsCodCta, 6, 1) = "1", "SOLES   ", "DOLARES ") & Space(vEspacio + 5) & "CREDITO : " & ImpreFormat(lsCodCta, 15) & IIf(Mid(lsCodCta, 6, 1) = "1", "SOLES   ", "DOLARES ") & Chr(10)
        Print #ArcSal, Tab(4); String(40, "-") & Space(vEspacio) & String(40, "-")
        Print #ArcSal, ""
        Print #ArcSal, Chr$(27) & Chr$(69);    'Activa Negrita
        Print #ArcSal, Tab(4); " MONTO DE OPERACION" & ImpreFormat(CCur(TxtMonPag.Text), 18, , True) & Space(vEspacio) & " MONTO DE OPERACION" & ImpreFormat(CCur(TxtMonPag.Text), 18, , True)
        Print #ArcSal, Chr$(27) & Chr$(70);    'Desactiva Negrita
        Print #ArcSal, ""
        Print #ArcSal, Tab(4); String(40, "-") & Space(vEspacio) & String(40, "-") & Chr(10)
        Print #ArcSal, Tab(4); Space(36) & Format(gsCodUser, "@@@@") & Space(vEspacio) & Space(36) & Format(gsCodUser, "@@@@")
        
    ImpreEnd
    MousePointer = 0
End Sub


Private Function MuestraEstudioAbog(pCredito As String, pConexion As ADODB.Connection) As String

Dim rsAb As New ADODB.Recordset
Dim rsPers As New ADODB.Recordset
Dim SSQL As String
 SSQL = "SELECT cCodAbog FROM ExpedJud WHERE cCodCta = '" & _
        pCredito & "'"
 rsAb.Open SSQL, pConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
 If rsAb.BOF And rsAb.EOF Then
   MuestraEstudioAbog = ""
 Else
   SSQL = "Select cNomPers FROM " & gcCentralPers & "Persona WHERE cCodPers = '" & rsAb!cCodAbog & "'"
   rsPers.Open SSQL, pConexion, adOpenForwardOnly, adLockReadOnly, adCmdText
   If rsPers.BOF And rsPers.EOF Then
     MuestraEstudioAbog = ""
   Else
     MuestraEstudioAbog = rsPers!cNomPers
   End If
   rsPers.Close
   Set rsPers = Nothing
 End If
 rsAb.Close
 Set rsAb = Nothing
End Function
