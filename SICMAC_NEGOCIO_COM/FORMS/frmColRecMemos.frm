VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColRecMemos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4695
   ClientLeft      =   2085
   ClientTop       =   2325
   ClientWidth     =   8595
   Icon            =   "frmColRecMemos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7110
      TabIndex        =   3
      Top             =   4140
      Width           =   1275
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "A&brir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   4140
      Width           =   1125
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   4140
      Width           =   1125
   End
   Begin RichTextLib.RichTextBox rtfMemo 
      Height          =   3585
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   6324
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   5
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmColRecMemos.frx":030A
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   540
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir/Guradar Como"
      Filter          =   "*.doc;*.txt"
   End
   Begin VB.Label lblCodCta 
      Alignment       =   2  'Center
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   60
      Width           =   2715
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNomPers 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Top             =   60
      Width           =   4515
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      Height          =   195
      Left            =   3180
      TabIndex        =   4
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmColRecMemos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Cuenta As String
Public Tipo As String
Dim clsMemo As UColRecMemos

Public Function Inicio(ByVal psTipo As String, ByVal psTexto As String, ByVal psCuenta As String, ByVal psNomPers As String) As UColRecMemos
    
    'Me.lblNombrePersona.Caption = psPersNombre
    Select Case psTipo
        Case "P"
            Me.Caption = " Recuperaciones -  Petitorio "
        Case "H"
            Me.Caption = " Recuperaciones -  Hechos    "
        Case "J"
            Me.Caption = " Recuperaciones -  Jur  "
        Case "P"
            Me.Caption = " Recuperaciones -  Pro "
        Case "C"
            Me.Caption = " Recuperaciones -  Complementos "
    End Select
    
    Me.rtfMemo.Text = psTexto
    Me.lblCodCta.Caption = psCuenta
    Me.lblNomPers.Caption = psNomPers
    
    Me.Show 1
    Set Inicio = clsMemo
    Set clsMemo = Nothing
End Function

Private Sub cmdAbrir_Click()
   CDialog.CancelError = True
    On Error GoTo ErrHandler    ' Establecer los indicadores
    'CDialog.Flags = cdlOFNHideReadOnly   ' Establecer los filtros
    CDialog.Filter = "Archivos RTF (*.RTF)|*.RTF|Archivos de texto" & _
    "(*.txt)|*.txt"
    ' Especificar el filtro predeterminado
    CDialog.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
    CDialog.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    rtfMemo.Filename = CDialog.Filename
    Exit Sub
ErrHandler:    ' El usuario ha hecho clic en el botón Cancelar
  Exit Sub

End Sub
Private Sub cmdCancelar_Click()
 rtfMemo.Text = ""
End Sub

Private Sub CmdAceptar_Click()
    clsMemo.CargaDatos Tipo, Me.rtfMemo.Text
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim loPrevio As previo.clsprevio
Dim loRec As COMNColocRec.NCOMColRecImpre
Dim lsCadImp As String
    Set loRec = New COMNColocRec.NCOMColRecImpre
        lsCadImp = loRec.ImprimeRecMemos(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, Me.lblCodCta.Caption, Me.lblNomPers.Caption, Me.rtfMemo.Text)
    Set loRec = Nothing
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
        'loPrevio.Show lsCadImp, "Memo Expediente en Recuperaciones", True
        loPrevio.Show lsCadImp, "Memo Expediente en Recuperaciones", True, , gImpresora
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
End Sub


Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.Caption = "Recuperaciones Memo "
    Set clsMemo = New UColRecMemos
End Sub
