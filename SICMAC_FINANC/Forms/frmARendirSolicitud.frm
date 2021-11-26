VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmARendirSolicitud 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5835
   ClientLeft      =   1230
   ClientTop       =   1575
   ClientWidth     =   9345
   ForeColor       =   &H00000000&
   Icon            =   "frmARendirSolicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   7215
      TabIndex        =   32
      Top             =   0
      Width           =   1980
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   690
         TabIndex        =   1
         Top             =   195
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   75
         TabIndex        =   33
         Top             =   248
         Width           =   540
      End
   End
   Begin Sicmact.Usuario Usu 
      Left            =   135
      Top             =   5310
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6690
      TabIndex        =   6
      Top             =   5355
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7935
      TabIndex        =   7
      Top             =   5355
      Width           =   1245
   End
   Begin VB.Frame frameRecibo 
      Caption         =   "RECIBO DE A RENDIR"
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
      Height          =   4620
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   9105
      Begin VB.Frame frameDestino 
         Caption         =   "Solicitante"
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
         Height          =   1770
         Left            =   300
         TabIndex        =   12
         Top             =   915
         Width           =   8475
         Begin Sicmact.TxtBuscar TxtBuscarPersCod 
            Height          =   360
            Left            =   975
            TabIndex        =   3
            Top             =   195
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   635
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TipoBusqueda    =   7
         End
         Begin VB.Label lblCodAge 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4875
            TabIndex        =   28
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label lblDescAge 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5340
            TabIndex        =   27
            Top             =   1365
            Width           =   2895
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agencia:"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   4140
            TabIndex        =   26
            Top             =   1395
            Width           =   705
         End
         Begin VB.Label lblCodArea 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   975
            TabIndex        =   25
            Top             =   1365
            Width           =   435
         End
         Begin VB.Label txtDirPer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   975
            TabIndex        =   24
            Top             =   960
            Width           =   7305
         End
         Begin VB.Label lblDescArea 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1395
            TabIndex        =   23
            Top             =   1365
            Width           =   2640
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Area :"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   90
            TabIndex        =   22
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label txtLEPer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6870
            TabIndex        =   21
            Top             =   570
            Width           =   1410
         End
         Begin VB.Label txtNomPer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   960
            TabIndex        =   20
            Top             =   585
            Width           =   5370
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre :"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   645
            Width           =   750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D.N.I."
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   6360
            TabIndex        =   15
            Top             =   630
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección :"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   14
            Top             =   1005
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código :"
            BeginProperty Font 
               Name            =   "Bookman Old Style"
               Size            =   9
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   90
            TabIndex        =   13
            Top             =   255
            Width           =   705
         End
      End
      Begin VB.TextBox txtDocNro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7050
         TabIndex        =   2
         Top             =   405
         Width           =   1605
      End
      Begin VB.TextBox txtConcepto 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   300
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2910
         Width           =   8490
      End
      Begin VB.TextBox txtImpCheque 
         Alignment       =   1  'Right Justify
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
         Left            =   7260
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   3705
         Width           =   1515
      End
      Begin VB.Label txtImpTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1290
         TabIndex        =   18
         Top             =   4215
         Width           =   7485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "La Cantidad de :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   255
         TabIndex        =   17
         Top             =   4260
         Width           =   1020
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Por concepto de"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   11
         Top             =   2685
         Width           =   1335
      End
      Begin VB.Label lblImporte 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   315
         Left            =   6360
         TabIndex        =   10
         Top             =   3690
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6165
         TabIndex        =   9
         Top             =   435
         Width           =   825
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   270
         Picture         =   "frmARendirSolicitud.frx":030A
         Stretch         =   -1  'True
         Top             =   180
         Visible         =   0   'False
         Width           =   2160
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
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
      Left            =   150
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   7050
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   60
         TabIndex        =   0
         Top             =   180
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1035
         TabIndex        =   31
         Top             =   180
         Width           =   4560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Proceso :"
         Height          =   195
         Left            =   5685
         TabIndex        =   30
         Top             =   255
         Width           =   675
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6360
         TabIndex        =   29
         Top             =   195
         Width           =   570
      End
   End
   Begin VB.Label lblSaldoActual 
      BackColor       =   &H80000005&
      Caption         =   "Saldo Actual: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   34
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmARendirSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSalir As Boolean
Dim lMN As Boolean, sMoney As String
Dim sSql As String, rs As New ADODB.Recordset
Dim cCodUsu As String, cNomUsu As String
Dim sCodAge As String, sNomAge As String
Dim sDocNat As String, sDocTpo As String, sDocEst As String * 1
Dim lTransActiva As Boolean, lbConfirma As Boolean, lArendir As Boolean
Dim sCtaOrig As String, sCtaOrigDesc As String
Dim sCtaDest As String, sCtaDestDesc As String
Dim sMovNroRef As String
Dim sObjtipo As String
Dim aObj(3, 4) As String

Dim objPista As COMManejador.Pista

'*****************************  NUEVO MODELO
Dim oNContFunc As NContFunciones
Dim oNArendir As NARendir
Dim oOpe As DOperacion
Dim lsDocTpo As String
Dim lsCtaContDebe As String, lsCtaContDDesc As String
Dim lsCtaContHaber As String, lsCtaContHDesc As String
Dim lnTipoArendir As ArendirTipo

Dim fnMovNroSolicitud As Long '***Agregado por ELRO el 20120525, según OYP-RFC016-2012
Dim fsMovNroSolicitud As String '***Agregado por ELRO el 20120525, según OYP-RFC016-2012
Dim fsCtaArendir As String '***Agregado por ELRO el 20120525, según OYP-RFC047-2012
Dim fsCtaFondofijo As String '***Agregado por ELRO el 20120525, según OYP-RFC047-2012

''Dim gvarp As gVarPublicas   '' **** Agregado por ANGC 20200306

Public Sub inicio(ByVal pnTipoArendir As ArendirTipo, Optional plConfirma As Boolean = False)
lbConfirma = plConfirma
lnTipoArendir = pnTipoArendir
lSalir = False
Me.Show 1
End Sub

'***Agregado por ELRO el 20120525, según OYP-RFC016-2012
Public Sub iniciarEdicion(ByVal pnTipoArendir As ArendirTipo, Optional plConfirma As Boolean = False, Optional ByVal psOpeCod As String = "", Optional pnMovNroSolicitud As Long = 0)
lbConfirma = plConfirma
lnTipoArendir = pnTipoArendir
gsOpeCod = psOpeCod
fnMovNroSolicitud = pnMovNroSolicitud
lSalir = False
Me.Show 1
End Sub
'***Fin Agregado por ELRO*******************************

Private Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If ValFecha(txtFecha) = False Then
    ValidaInterfaz = False
    Exit Function
End If
If Len(Trim(TxtBuscarPersCod.Text)) = 0 Then
    MsgBox "Persona no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    '***Modificado por ELRO el 20120623, según OYP-RFC047-2012
    'TxtBuscarPersCod.SetFocus
    If frameDestino.Enabled Then
        TxtBuscarPersCod.SetFocus
    End If
    '***Fin Modificado por ELRO*******************************
    Exit Function
End If

If Len(Trim(lblCodArea)) = 0 Then
    MsgBox "Area no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    TxtBuscarPersCod.SetFocus
    Exit Function
End If
If Len(Trim(Me.lblCodAge)) = 0 Then
    MsgBox "Agencia no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    TxtBuscarPersCod.SetFocus
    Exit Function
End If

If fraCajaChica.Visible Then
    If Len(Trim(txtBuscarAreaCH)) = 0 Then
        MsgBox "Caja Chica Ingresada no válida", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtBuscarAreaCH.SetFocus
        Exit Function
    End If
End If
'***Comentado por ELRO el 20120504, según OYP-RFC016-2012
'If fraArendir.Visible Then
'    If Len(Trim(TxtBuscarArendir)) = 0 Then
'        MsgBox "Area/Agencia a quien solicita el Arendir no Ingresada", vbInformation, "Aviso"
'        ValidaInterfaz = False
'        TxtBuscarArendir.SetFocus
'        Exit Function
'    End If
'End If
If frameRecibo.Enabled Then
    If Len(Trim(txtDocNro)) = 0 Then
        MsgBox "Nro de Documento no válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtDocNro.Enabled Then
            txtDocNro.SetFocus
        End If
        Exit Function
    End If
    If Val(txtImpCheque) = 0 Then
        MsgBox "Importe de a Rendir no válida", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtImpCheque.SetFocus
        Exit Function
    End If
    If Len(Trim(txtConcepto)) = 0 Then
        MsgBox "Concepto no ha se ha Ingresado o no es válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtConcepto.SetFocus
        Exit Function
    End If
    If Len(Trim(TxtBuscarPersCod)) = 0 Then
        MsgBox "Persona no se ha ingresado o código no es válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        TxtBuscarPersCod.SetFocus
        Exit Function
    End If
End If

'*** PEAC 20110107
'If BuscaRendPendiente(Trim(TxtBuscarPersCod)) Then
'    MsgBox "La persona ingresada mantiene una Solicitud pendiente, por favor proceda a realizar la Rendición de esta, para realizar una nueva.", vbInformation, "Aviso"
'    ValidaInterfaz = False
'    TxtBuscarPersCod.SetFocus
'    Exit Function
'End If

End Function

Private Sub cmdAceptar_Click()

''AGREGADO POR ANGC 20200306 -  VALIDA FECHA DEL SISTEMA Y EL SERVER
Dim msj As String
msj = gVarPublicas.ValidarFechaSistServer
If msj <> "" Then
    MsgBox msj, vbInformation, "Aviso"
    Unload frmARendirSolicitud
Else
    If gsOpeCod <> CStr(gCGArendirCtaSolcEditMN) And gsOpeCod <> CStr(gCGArendirCtaSolcEditME) Then
        Dim ctrControl As Control
        Dim nImporte As Currency
        Dim N As Integer
        Dim ldFecha As Date
        Dim sMsg As String
        Dim sCta As String
        Dim oCajaChica As nCajaChica
        Dim lnSaldo As Currency
        Dim lnTope As Currency
        Dim lsMovNro As String
        
        '******************** nuevo modelo *******************************
        Dim lsCodAgeArea As String
        Dim lnPos As String
        Dim lsCadenaPrint    As String
        Dim lnMovNroSol As Long '***Agregado por ELRO el 20120615, según OYP-RFC047-2012
        Dim lnMovNroAte As Long '***Agregado por ELRO el 20120516, según OYP-RFC047-2012
        On Error GoTo ErrSql
        
        Set oCajaChica = New nCajaChica
        If ValidaInterfaz = False Then Exit Sub
        
            If lnTipoArendir = gArendirTipoCajaChica Then
                lnSaldo = oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2))
                If lnSaldo = 0 Then
                    MsgBox "Caja Chica Sin Saldo. Es necesario Solicitar Autorización o Desembolso", vbInformation, "Aviso"
                    Exit Sub
                End If
                If oCajaChica.VerificaTopeCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2)) = True Then
                    '***Modificado por ELRO el 20120620, según OYP-RFC047-2012
                    'If MsgBox("Saldo de esta Caja Chica es menor que el permitido como monto minimo." & oImpresora.gPrnSaltoLinea _
                    '            & "Se recomienda realizar Rendición respectiva. ¿ Desea Confirmar Recibo de Arendir ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                    '    Exit Sub
                    'End If
                    MsgBox "Saldo de esta Caja Chica es menor que el permitido como monto mínimo." & Chr(10) & "Se recomienda realizar la Rendición a Contabilidad respectiva. "
                    '***Fin Modificado por ELRO*******************************
               End If
            End If
            'Volvemos a generar Nro. de Mov. en caso de que el generado ya exista
            If lnTipoArendir <> gArendirTipoCajaChica Then
               sMsg = " ¿ Seguro de Grabar Recibo de A rendir Cuenta ? "
            Else
               sMsg = " ¿ Seguro de Grabar Recibo de A rendir de Caja Chica ? "
            End If
            If MsgBox(sMsg, vbQuestion + vbYesNo, "Aviso de Confirmación") = vbNo Then
               Exit Sub
            End If
            lsMovNro = oNContFunc.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
            nImporte = Val(Format(txtImpCheque, gsFormatoNumeroDato))
            '***Modificado por ELRO el 20120504, según OYP-RFC015-2012
            'If oNArendir.GrabaSolicitudArendir(lnTipoArendir, lsMovNro, gsOpeCod, _
            '                                 Mid(TxtBuscarArendir, 4, 2), Mid(TxtBuscarArendir, 1, 3), Trim(txtConcepto), TxtBuscarPersCod.Text, nImporte, _
            '                                Trim(txtDocNro), lsDocTpo, Format(txtFecha, gsFormatoFecha), Mid(txtBuscarAreaCH.Text, 1, 3), Mid(txtBuscarAreaCH.Text, 4, 2), Val(lblNroProcCH)) = 0 Then
                
            If oNArendir.GrabaSolicitudArendir(lnTipoArendir, lsMovNro, gsOpeCod, _
                                                  "", "", Trim(txtConcepto), TxtBuscarPersCod.Text, _
                                                  nImporte, Trim(txtDocNro), lsDocTpo, _
                                                  Format(txtFecha, gsFormatoFecha), _
                                                  Mid(txtBuscarAreaCH.Text, 1, 3), _
                                                  Mid(txtBuscarAreaCH.Text, 4, 2), _
                                                  Val(lblNroProcCH), lnMovNroSol) = 0 Then
                                                  '***Agregado por ELRO el parametro lnMovNroSol el 20120615
                    '***Agregado por ELRO el 20120615, según OYP-RFC047-2012
                If lnMovNroSol > 0 And (gsOpeCod = CStr(gCHArendirCtaSolMN) Or gsOpeCod = CStr(gCHArendirCtaSolME)) Then
                        Dim lsSubCta As String
                        Dim lsMovNro2 As String
                        Dim oNContImprimir2 As NContImprimir
                        Set oNContImprimir2 = New NContImprimir
                        Dim lsTexto As String
                        
                        lsMovNro2 = oNContFunc.GeneraMovNro(gdFecSis, _
                                                            gsCodAge, _
                                                            gsCodUser)
                                                                 
                        lsSubCta = oNContFunc.GetFiltroObjetos(ObjCMACAgenciaArea, _
                                                                    fsCtaFondofijo, _
                                                                    Trim(txtBuscarAreaCH), _
                                                                    False)
                        '***Modificado por ELRO el 20130221, según SATI INC1301300007
                        'fsCtaFondofijo = fsCtaFondofijo + IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
                        If Trim(lsSubCta) <> "" Then
                            fsCtaFondofijo = fsCtaFondofijo + IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
                        End If
                        '***Modificado por ELRO el 20130221**************************
                        
                        
                        Call oNArendir.GrabaAtencionArendirCH(gArendirTipoCajaChica, _
                                                              gsFormatoFecha, _
                                                              lnMovNroSol, _
                                                              lsMovNro2, _
                                                              IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaAtencMN, gCHArendirCtaAtencME), _
                                                              txtConcepto, _
                                                              txtImpCheque, _
                                                              Mid(txtBuscarAreaCH, 1, 3), _
                                                              Mid(txtBuscarAreaCH, 4, 2), _
                                                              Val(lblNroProcCH), _
                                                              fsCtaArendir, _
                                                              fsCtaFondofijo, _
                                                              txtDocNro, _
                                                              gdFecSis, _
                                                              lnMovNroAte) '***Modificado "gdFecha" a "gdFecSis" por ELRO el 20121026, según SATI INC1210240007
                        '***Fin Agregado por ELRO*******************************
                End If
                '***Fin Modificado por ELRO*******************************
                If lbConfirma Then
                    Unload Me
                    Exit Sub
                Else
                        '***Comentado por ELRO el 20120620, según OYP-RFC047-2012
                        'Dim oNContImprimir As NContImprimir
                        'Set oNContImprimir = New NContImprimir
                        'lsCadenaPrint = oNContImprimir.ImprimeReciboARendir(lsMovNro, gnColPage, gsInstCmac, gsNomCmac, gsNomCmacRUC)
                        'lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina & lsCadenaPrint
                        'Set oNContImprimir = Nothing
                        'EnviaPrevio lsCadenaPrint, Me.Caption, gnLinPage, False
                        '***Comentado por ELRO***********************************
                        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Pago a Proveedores"
                        '*** PEAC - 20110216 - UNA VEZ REGISTRADO LA SOLICITUD YA NO TIENE PORQUE REALIZAR OTRA SI NO HA RENDIDO LA ANTERIOR
                '        If MsgBox(" ¿ Desea continuar registrando Recibos de A rendir ? ", vbQuestion + vbYesNo, "Recibo de Egresos") = vbYes Then
                '            txtImpCheque = "0.00": txtImpTexto = ""
                '            txtDocNro = oNContFunc.GeneraDocNro(Val(lsDocTpo), gsCodUser, Mid(gsOpeCod, 3, 1))
                '            txtConcepto = ""
                '            txtConcepto.SetFocus
                '        Else
                            Unload Me
                '        End If
                        '*** FIN PEAC
                End If
            End If
        Else
        If oNArendir.actualizarSolicitudArendir(fnMovNroSolicitud, txtConcepto, CCur(txtImpCheque)) = True Then
            Dim oNContImprimir As NContImprimir
            Set oNContImprimir = New NContImprimir
            lsCadenaPrint = oNContImprimir.ImprimeReciboARendir(fsMovNroSolicitud, gnColPage, gsInstCmac, gsNomCmac, gsNomCmacRUC)
            lsCadenaPrint = lsCadenaPrint & oImpresora.gPrnSaltoPagina & lsCadenaPrint
            Set oNContImprimir = Nothing
            EnviaPrevio lsCadenaPrint, Me.Caption, gnLinPage, False
            Unload Me
        End If
    End If
End If

Exit Sub
ErrSql:
   MsgBox Err.Number & vbCrLf & TextErr(Err.Description), vbInformation, "Aviso"
End Sub
Private Sub LimpiaControles()
If lbConfirma Then
   txtFecha = ""
Else
   'lblMovNro = oNContFunc.GeneraMovNro(gdFecSis, , gsCodUser)
   txtFecha = Format(gdFecSis, "dd/mm/yyyy")
End If
txtDocNro = "": txtImpCheque = "0.00"
txtImpTexto = "": txtConcepto = ""
txtLEPer = ""
txtNomPer = "": txtDirPer = ""
End Sub
'***Agregado por ELRO el 20120414, según OYP-RFC016-2012
Private Sub LimpiaControles_2()

txtFecha = Format(gdFecSis, "dd/mm/yyyy")
txtDocNro = ""
TxtBuscarPersCod = ""
txtNomPer = ""
txtLEPer = ""
txtDirPer = ""
lblCodArea = ""
lblDescArea = ""
lblCodAge = ""
lblDescAge = ""
txtConcepto = ""
txtImpCheque = "0.00"
txtImpTexto = ""
End Sub
'***Fin Agregado por ELRO*******************************

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Activate()
If lSalir Then
   Unload Me
   Exit Sub
End If
If lbConfirma Then
   txtImpCheque.Enabled = False
   txtConcepto.Enabled = False
Else
   txtDocNro.Enabled = False
End If

'***Agregado por ELRO el 20120414, según OYP-RFC016-2012
If gsOpeCod = CStr(gCGArendirCtaSolcEditMN) Or gsOpeCod = CStr(gCGArendirCtaSolcEditME) Then
    frameDestino.Enabled = False
    txtFecha.Enabled = False
End If
'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120615, según OYP-RFC047-2012
If gsOpeCod = CStr(gCHArendirCtaSolMN) Or gsOpeCod = CStr(gCHArendirCtaSolME) Then

    fsCtaArendir = oOpe.EmiteOpeCta(IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaAtencMN, gCHArendirCtaAtencME), "D")
    fsCtaFondofijo = oOpe.EmiteOpeCta(IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaAtencMN, gCHArendirCtaAtencME), "H")
    
    If Trim(fsCtaArendir) = "" Or Trim(fsCtaFondofijo) = "" Then
        MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
        lSalir = True
        Exit Sub
    End If
    txtBuscarAreaCH_EmiteDatos
End If
'***Fin Agregado por ELRO*******************************
End Sub
Private Sub AsignaValores()
'***Agregado por ELRO el 20120414, según OYP-RFC016-2012
If validarARendirCuentasPendientes = False Then
    LimpiaControles_2
    Exit Sub
End If
'***Fin Agregado por ELRO*******************************

TxtBuscarPersCod.Text = Usu.PersCod
lblCodArea = Usu.AreaCod
lblCodAge = Usu.CodAgeAct
lblDescAge = Usu.DescAgeAct
lblDescArea = Usu.AreaNom
txtDirPer = Usu.DireccionUser
txtLEPer = IIf(Usu.NroDNIUser = "", Usu.NroRucUser, Usu.NroDNIUser)
txtNomPer = PstaNombre(Usu.UserNom)
'If Me.txtConcepto.Visible Then
'    Me.txtConcepto.SetFocus
'End If

'***Agregado por ELRO el 20120625, según OYP-RFC047-2012
mostrarSaldoActual
'***Fin Agregado por ELRO*******************************
End Sub

'***Agregado por ELRO el 20120414, según OYP-RFC016-2012
Private Function validarARendirCuentasPendientes() As Boolean
validarARendirCuentasPendientes = False
If lnTipoArendir = gArendirTipoCajaGeneral Then
    Dim oArendir As NARendir
    Set oArendir = New NARendir
    Dim nSaldoPendienteMN  As Currency
    Dim nSaldoPendienteME  As Currency
    
    '********Agregado por PASI20131118 segun TI-ERS107-2013
    Dim RsRendicion As ADODB.Recordset
    Set RsRendicion = New ADODB.Recordset
    '****FIN PASI
    
    Call oArendir.obtenerSaldoARendirCuentas(Usu.PersCod, nSaldoPendienteMN, nSaldoPendienteME)
    If nSaldoPendienteMN > 0 Then
        MsgBox PstaNombre(Usu.UserNom) & " tiene un Saldo pendiente de " & nSaldoPendienteMN & " Nuevo Soles." & Chr(13) & "Primero sustente y/o rinda." & Chr(13) & "Consultar Reglamento de Entregas a Rendir en la Intranet.", vbInformation, "Aviso"
        Exit Function
    ElseIf nSaldoPendienteME > 0 Then
        MsgBox PstaNombre(Usu.UserNom) & " tiene un Saldo pendiente de " & nSaldoPendienteME & " Dólares." & Chr(13) & "Primero sustente y/o rinda." & Chr(13) & "Consultar Reglamento de Entregas a Rendir en la Intranet.", vbInformation, "Aviso"
        Exit Function
    ElseIf nSaldoPendienteMN = -1 Or nSaldoPendienteME = -1 Then
        MsgBox PstaNombre(Usu.UserNom) & " tiene una Solicitud pendiente por aprobar." & Chr(13) & "Primero que lo eliminen, para registrar una nueva Solicitud.", vbInformation, "Aviso"
        Exit Function
    End If
    
    '********Agregado por PASI20131118 segun TI-ERS107-2013
    Set RsRendicion = oArendir.ObtenerARendirCuentasParaRendirxPersona(Usu.PersCod, "1")
    If RsRendicion.RecordCount > 0 Then
        MsgBox PstaNombre(Usu.UserNom) & "; tiene pendiente una rendición en Moneda Nacional,no se puede realizar una nueva solicitud hasta que se rinda cuenta de la anterior"
        Exit Function
    End If
    Set RsRendicion = Nothing
    Set RsRendicion = oArendir.ObtenerARendirCuentasParaRendirxPersona(Usu.PersCod, "2")
    If RsRendicion.RecordCount > 0 Then
        MsgBox PstaNombre(Usu.UserNom) & "; tiene pendiente una rendición en Moneda Extranjera,no se puede realizar una nueva solicitud hasta que se rinda cuenta de la anterior"
        Exit Function
    End If
    Set RsRendicion = Nothing
    '********Fin Agregado por PASI20131118
End If

validarARendirCuentasPendientes = True
End Function
'***Fin Agregado por ELRO*******************************

Private Sub Form_Load()

If gsOpeCod <> CStr(gCGArendirCtaSolcEditMN) And gsOpeCod <> CStr(gCGArendirCtaSolcEditME) Then

    Dim sTxt As String
    Dim N As Integer
    If gsCodCMAC = "112" Then
        Me.Image1.Visible = True
    End If
    Usu.inicio gsCodUser
    AsignaValores
    Me.Caption = gsOpeDesc
    
    CentraForm Me
    lSalir = False
    If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then    'Identificación de Tipo de Moneda
       txtImpCheque.BackColor = vbGreen
       If gnTipCambio = 0 Then
          If Not GetTipCambio(gdFecSis) Then
             lSalir = True
             Exit Sub
          End If
       End If
    End If
    
    Set oNContFunc = New NContFunciones
    Set oNArendir = New NARendir
    Set oOpe = New DOperacion
    
    Set objPista = New COMManejador.Pista
    
    txtFecha = gdFecSis
    txtFecha.Enabled = False
    lsDocTpo = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioNoDebeExistir, OpeDocMetAutogenerado)
    If lsDocTpo <> "" Then
        txtDocNro = oNContFunc.GeneraDocNro(Val(lsDocTpo), gsCodUser, Mid(gsOpeCod, 3, 1))
    End If
    
    TxtBuscarPersCod.TipoBusqueda = buscaempleado
    If lnTipoArendir = gArendirTipoCajaChica Then
        txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
        txtBuscarAreaCH.rs = oNArendir.EmiteCajasChicas
        frameRecibo.Caption = frameRecibo.Caption & " CAJA CHICA"
        fraCajaChica.Visible = True
        '***Agregado por ELRO el 20120623, según OYP-RFC047-2012
        fraCajaChica.Enabled = False
        verificarEncargadoCH
        If Len(Trim(txtBuscarAreaCH)) = 0 Then
            frameDestino.Enabled = False
            LimpiaControles_2
            Exit Sub
        End If
        mostrarSaldoActual
        '***Fin Agregado por ELRO*******************************
        'fraArendir.Visible = False '***Comentado por ELRO el 20120504, según OYP-RFC016-2012
    Else
       Dim oAreas As DActualizaDatosArea
       Set oAreas = New DActualizaDatosArea
       '***Comentado por ELRO el 20120504, según OYP-RFC016-2012
       'Set rs = oOpe.CargaOpeObj(gsOpeCod, 1)
       'TxtBuscarArendir.psRaiz = "A Rendir de..."
       'If Not rs.EOF Then
       '     TxtBuscarArendir.rs = oAreas.GetAgenciasAreas(rs!cOpeObjFiltro, 1)
       'End If
       Set oAreas = Nothing
    End If
    'lsCtaContDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
    'lsCtaContHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
    'lsCtaContDDesc = oNContFunc.EmiteCtaContDesc(lsCtaContDebe)
    'lsCtaContHDesc = oNContFunc.EmiteCtaContDesc(lsCtaContHaber)
    
    If lbConfirma Then
       txtDocNro.TabIndex = 0
    End If
    frameRecibo.Enabled = True
Else
    Set oNArendir = New NARendir
    Dim rsSolicitud As ADODB.Recordset
    Set rsSolicitud = New ADODB.Recordset
    

    If fnMovNroSolicitud = 0 Then
        Me.Caption = gsOpeDesc
    Else
        Me.Caption = "Editar A Rendir Cuenta"
    End If

    CentraForm Me
    lSalir = False
    If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then    'Identificación de Tipo de Moneda
       txtImpCheque.BackColor = vbGreen
       If gnTipCambio = 0 Then
          If Not GetTipCambio(gdFecSis) Then
             lSalir = True
             Exit Sub
          End If
       End If
    End If
    
    Set rsSolicitud = oNArendir.obtenerSolicitudARendirCta(fnMovNroSolicitud)
    
    If Not rsSolicitud.BOF And Not rsSolicitud.EOF Then
    
        TxtBuscarPersCod = rsSolicitud.Fields(0)
        TxtBuscarPersCod.psCodigoPersona = rsSolicitud.Fields(0)
        txtNomPer = rsSolicitud.Fields(1)
        txtLEPer = rsSolicitud.Fields(3)
        txtDirPer = rsSolicitud.Fields(2)
        lblCodArea = rsSolicitud.Fields(4)
        lblDescArea = rsSolicitud.Fields(5)
        lblCodAge = rsSolicitud.Fields(6)
        lblDescAge = rsSolicitud.Fields(7)
        txtConcepto = rsSolicitud.Fields(8)
        txtDocNro = rsSolicitud.Fields(9)
        txtFecha = rsSolicitud.Fields(10)
        txtImpCheque = Format(rsSolicitud.Fields(11), gsFormatoNumeroView)
        fsMovNroSolicitud = rsSolicitud.Fields(12)
        txtImpTexto = ConvNumLet(Val(Format(txtImpCheque, gsFormatoNumeroDato)), , , Mid(gsOpeCod, 3, 1))


    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub txtBuscarAreaCH_EmiteDatos()
Dim oCajaCH As nCajaChica
Set oCajaCH = New nCajaChica
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
If lblCajaChicaDesc <> "" And txtBuscarAreaCH.Enabled Then
   If txtFecha.Enabled Then
      txtFecha.SetFocus
   Else
     TxtBuscarPersCod.SetFocus
   End If
End If
Set oCajaCH = Nothing
End Sub

'***Comentado por ELRO el 20120504, según OYP-RFC016-212
'Private Sub TxtBuscarArendir_EmiteDatos()
'lblDescArendir = Trim(TxtBuscarArendir.psDescripcion)
'If txtFecha.Visible Then
'    If txtFecha.Enabled Then
'        txtFecha.SetFocus
'    End If
'End If
'End Sub

Private Sub TxtBuscarPersCod_EmiteDatos()
'***Comentado por ELRO el 20120504, según OYP-RFC016-2012
Dim oDPersonas As DPersonas
Set oDPersonas = New DPersonas
Dim rsCodUserEmp As ADODB.Recordset
Set rsCodUserEmp = New ADODB.Recordset
'***Fin Comentado por ELRO*******************************
'If TxtBuscarPersCod.psDescripcion <> "" Then
Usu.DatosPers TxtBuscarPersCod.Text
If Usu.PersCod = "" Then
    MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
Else
    txtConcepto.SetFocus
End If
AsignaValores
'***Agregado por ELRO el 20120504, según OYP-RFC016-2012
If lnTipoArendir = gArendirTipoCajaGeneral Then
    Set rsCodUserEmp = oDPersonas.BuscaCliente(TxtBuscarPersCod.psCodigoPersona, BusquedaEmpleadoCodigo)
    
    If Not rsCodUserEmp.BOF And Not rsCodUserEmp.EOF Then
        lsDocTpo = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioNoDebeExistir, OpeDocMetAutogenerado)
        If lsDocTpo <> "" Then
            txtDocNro = oNContFunc.GeneraDocNro(Val(lsDocTpo), rsCodUserEmp.Fields(14), Mid(gsOpeCod, 3, 1))
        End If
    Else
        lsDocTpo = oOpe.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioNoDebeExistir, OpeDocMetAutogenerado)
        If lsDocTpo <> "" Then
            txtDocNro = oNContFunc.GeneraDocNro(Val(lsDocTpo), gsCodUser, Mid(gsOpeCod, 3, 1))
        End If
    End If
End If


Set rsCodUserEmp = Nothing
Set oDPersonas = Nothing
'***Fin Agregado por ELRO*******************************
'End If
End Sub
Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   txtImpCheque.SetFocus
End If
End Sub

Private Sub txtDocNro_GotFocus()
txtDocNro.SelStart = 0
txtDocNro.SelLength = Len(txtDocNro)
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
Dim nPos As Integer
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   nPos = InStr(1, txtDocNro, "-")
   If nPos > 0 Then
      txtDocNro = Mid(txtDocNro, 1, nPos) & Right(String(8, "0") & Mid(txtDocNro, nPos + 1, 8), 8)
   Else
      txtDocNro = Right(String(8, "0") & txtDocNro, 8)
   End If
   fEnfoque txtDocNro
End If
End Sub
'Private Function ValidaRecibo() As Boolean
'Dim rsVer As New ADODB.Recordset
'ValidaRecibo = False
'SSQL = "SELECT a.cMovNro, a.dDocFecha, b.cMovCabObjOrden, b.cObjetoCod, c.cObjetoDesc, 0 as nMovMonto, '' as cDirPers, '' as cMovDesc, '' as cNuDocI, '' as cMovEstado " _
'     & "FROM   MovDoc a,  MovCabObj b,  Objeto c " _
'     & "WHERE  a.nDocTpo='" & gnDocTpo & "' and a.cDocNro = '" & txtDocNro & "' and b.cMovNro = a.cMovNro AND " _
'     & "       b.cObjetoCod <> '" & sObjtipo & "' and c.cObjetoCod = b.cObjetoCod and " _
'     & "       EXISTS (SELECT cMovNro from  MovCabObj d Where a.cMovNro = d.cMovNro and " _
'     & "                      d.cObjetoCod = '" & sObjtipo & "') " _
'     & "UNION " _
'     & "SELECT a.cMovNro, a.dDocFecha, b.cMovCabObjOrden, b.cObjetoCod, c.cNomPers as cObjetoDesc, e.nMovMonto, c.cDirPers, e.cMovDesc, isnull(c.cNuDocI,'') as cNuDocI, e.cMovEstado " _
'     & "FROM   MovDoc a,  MovCabObj b, Persona c,  Mov e " _
'     & "WHERE  e.cMovFlag <> 'X' and a.nDocTpo='" & gnDocTpo & "' and a.cDocNro = '" & txtDocNro & "' and b.cMovNro = a.cMovNro AND " _
'     & "       b.cObjetoCod <> '" & sObjtipo & "' and substring(b.cObjetoCod,3,10) = c.cCodPers and " _
'     & "       e.cMovNro = a.cMovNro and " _
'     & "       EXISTS (SELECT cMovNro FROM  MovCabObj d Where a.cMovNro = d.cMovNro and " _
'     & "                     d.cObjetoCod = '" & sObjtipo & "') ORDER BY b.cMovCabObjOrden "
'   If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
'   rs.Open SSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'If rs.RecordCount <> 3 Then
'   MsgBox "Recibo no existe. Por favor reintentar", vbInformation, "Aviso"
'   Exit Function
'End If
'If rs!cMovEstado = "1" Then
'   MsgBox " Recibo está ANULADO...! ", vbInformation, "Error"
'   Exit Function
'End If
'If rs!cMovEstado = "2" Then
'   MsgBox " Recibo fue RECHAZADO...! ", vbInformation, "Error"
'   Exit Function
'End If
'
'SSQL = "SELECT cMovNro FROM  MovRef WHERE cMovNroRef = '" & rs!cMovNro & "'"
'Set rsVer = CargaRecord(SSQL)
'If Not RSVacio(rsVer) Then
'   MsgBox "Recibo de Egreso ya fue confirmado...", vbInformation, "Error"
'   Exit Function
'End If
'sMovNroRef = rs!cMovNro
'txtAgeCod = rs!cObjetoCod
'txtAgeDesc = rs!cObjetoDesc
'rs.MoveNext
'txtPerCod = rs!cObjetoCod
'txtNomPer = rs!cObjetoDesc
'txtDirPer = rs!cDirPers
'txtLEPer = rs!cNudoci
'txtImpCheque = Format(rs!nMovMonto, gsFormatoNumeroView)
'txtImpTexto = ConvNumLet(rs!nMovMonto)
'txtConcepto = rs!cMovDesc
'rs.MoveNext
'txtCajaCod = rs!cObjetoCod
'txtCajaDes = rs!cObjetoDesc
'ValidaRecibo = True
'rs.Close
'rsVer.Close
'Set rs = Nothing
'Set rsVer = Nothing
'End Function

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = False Then Exit Sub
    txtConcepto.SetFocus
End If
End Sub

Private Sub txtImpCheque_GotFocus()
fEnfoque txtImpCheque
End Sub

Private Sub txtImpCheque_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImpCheque, KeyAscii, 15, 2)
If KeyAscii = 13 Then
   CmdAceptar.SetFocus
End If
End Sub
Private Sub txtImpCheque_LostFocus()
Dim nImporte As Currency
Dim oCajaCH As nCajaChica
Dim lnImporteTope As Currency
Dim lnImporteArendir As Currency
Dim oConst As NConstSistemas

Set oConst = New NConstSistemas

Set oCajaCH = New nCajaChica



lnImporteArendir = oConst.LeeConstSistema(39)
If Mid(gsOpeCod, 3, 1) = Moneda.gMonedaExtranjera Then
    lnImporteArendir = Format(lnImporteArendir / gnTipCambio, "#.00")
End If

txtImpCheque = Format(txtImpCheque, gsFormatoNumeroView)
   If lnTipoArendir = gArendirTipoCajaGeneral Or lnTipoArendir = gArendirTipoAgencias Then
      If nVal(txtImpCheque) <= lnImporteArendir Then
         MsgBox "El Importe para solicitar A rendir cuenta debe ser mayor a " & Format(lnImporteArendir, gsFormatoNumeroView) & ". " & oImpresora.gPrnSaltoLinea & "En caso contrario solicite A rendir de Caja Chica", vbInformation, "Error"
         txtImpCheque = "0.00"
         txtImpTexto = ""
         Exit Sub
      End If
   Else
      
      lnImporteTope = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), MontoTope)
      If Val(Format(txtImpCheque, gsFormatoNumeroDato)) > lnImporteTope Then
         MsgBox "El Importe para solicitar a Caja Chica no puede ser mayor a " & Format(lnImporteTope, gsFormatoNumeroView) & ". " & oImpresora.gPrnSaltoLinea & "En caso contrario solicite un A rendir Cuenta con Caja General", vbInformation, "Error"
         txtImpCheque = "0.00"
         txtImpTexto = ""
         Exit Sub
      End If
   End If
   txtImpTexto = ConvNumLet(Val(Format(txtImpCheque, gsFormatoNumeroDato)), , , Mid(gsOpeCod, 3, 1))
End Sub

'*** PEAC 20110105
Private Function BuscaRendPendiente(ByVal psPersCod As String) As Boolean
Dim rs As ADODB.Recordset
Set oNArendir = New NARendir
Set rs = oNArendir.BuscaPendienteRendicion(psPersCod)

    If Not (rs.EOF And rs.BOF) Then
        BuscaRendPendiente = True
    Else
        BuscaRendPendiente = False
    End If
    
Set rs = Nothing
'Set oNArendir = Nothing

End Function

'***Agregado por ELRO el 20120623, según OYP-RFC047-2012
Private Sub mostrarSaldoActual()
Dim oCajaChica As nCajaChica
Set oCajaChica = New nCajaChica
If gsOpeCod = CStr(gCHArendirCtaSolMN) Or gsOpeCod = CStr(gCHArendirCtaSolME) Then
    lblSaldoActual.Visible = True
    lblSaldoActual = "Saldo Actual: " & oCajaChica.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), SaldoActual)
Else
    lblSaldoActual.Visible = False
End If
Set oCajaChica = Nothing

End Sub

Private Sub verificarEncargadoCH()
    Dim oNCajaChica As nCajaChica
    Set oNCajaChica = New nCajaChica
    Dim rsEncargado As ADODB.Recordset
    Set rsEncargado = New ADODB.Recordset
    
    Set rsEncargado = oNCajaChica.verificarEncargadoCH(gsCodPersUser)
    
    If Not rsEncargado.BOF And Not rsEncargado.EOF Then
        txtBuscarAreaCH = rsEncargado!cAreaCod & rsEncargado!cAgeCod
    Else
        MsgBox "No carga el código de la Caja Chica por los siguientes motivos:" & Chr(10) & "1. No esta encargado de la Caja Chica." & Chr(10) & "2. Aún no esta Autorizado el nuevo proceso de la Caja Chica." & Chr(10) & "3. Aún no cobra el efectivo habilitado por la Caja Chica.", vbInformation, "Aviso"
    End If
    Set rsEncargado = Nothing
    Set oNCajaChica = Nothing
End Sub
'***Fin Agregado por ELRO*******************************


