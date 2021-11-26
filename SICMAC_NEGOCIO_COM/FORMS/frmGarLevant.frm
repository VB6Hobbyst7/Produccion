VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGarLevant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Levantamiento de Garantias Reales"
   ClientHeight    =   4770
   ClientLeft      =   2730
   ClientTop       =   1125
   ClientWidth     =   7920
   Icon            =   "frmGarLevant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameCreditos 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1200
      TabIndex        =   15
      Top             =   5160
      Width           =   6495
      Begin VB.CommandButton CmdSalir3 
         Caption         =   "&Salir"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2270
         Width           =   6255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHCreditos 
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   4320
      Width           =   1365
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame FraResultado 
      Height          =   2775
      Left            =   8400
      TabIndex        =   10
      Top             =   1080
      Width           =   6855
      Begin VB.CommandButton CmdSalir2 
         Caption         =   "&Salir"
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   6375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHGARA 
         Height          =   2055
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   3625
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   1365
   End
   Begin VB.TextBox txtGarantia 
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
      Left            =   1290
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1125
   End
   Begin VB.Frame fraComentario 
      Caption         =   "Comentario"
      Enabled         =   0   'False
      Height          =   1860
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   7725
      Begin VB.TextBox txtComentario 
         Height          =   1455
         Left            =   180
         MaxLength       =   254
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   7380
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   4320
      Width           =   1365
   End
   Begin VB.CommandButton CmdLevantar 
      Caption         =   "&Levantar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1365
   End
   Begin VB.Frame fraPersonas 
      Caption         =   "Personas Relacionadas"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7725
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.Label lblCodigoAnterior 
      Height          =   285
      Left            =   1170
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Garantía :"
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
      Left            =   210
      TabIndex        =   6
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "frmGarLevant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* Garantias Reales
Option Explicit
Dim Garantia As String
Dim Moneda As String
Dim Realizacion As Double
Dim PersCod As String
Dim nTipoGar As Integer
Private Sub CmdAceptar_Click()
Dim OptBt As Integer
Dim OptBt2 As Integer
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim lsCadImp As String   'cadena q forma
Dim loPrevio As Previo.clsPrevio
Dim Garant As COMDCredito.DCOMGarantia 'DGarantia
Dim REP As COMNCredito.NCOMCredDoc 'NCredDoc
Dim nFicSal As Integer
Set Garant = New COMDCredito.DCOMGarantia
Set REP = New COMNCredito.NCOMCredDoc

Dim MovGar As Long
  
OptBt = MsgBox("Desea hacer el Levantamiento de la Garantia " & IIf(nTipoGar = 1, "Real", ""), vbYesNo, "CMACT")

If vbYes = OptBt Then

    MovGar = Garant.InsertaLiberacion(txtGarantia, txtComentario, gsGarantLevanta, _
                                  gdFecSis, gsCodAge, gsCodUser)
    If nTipoGar = 1 Then
    lsDestino = "P"
    'lsCadImp = REP.ImprimeBoletaGarantiaR(gsNomAge, gdFecSis, Time, txtGarantia, gsCodUser, MovGar, gsNomCmac, gdFecSis)
          If Len(Trim(lsCadImp)) > 0 Then
            Set loPrevio = New Previo.clsPrevio
            frmImpresora.Show 1
            Do
             OptBt2 = MsgBox("Desea Imprimir la Boleta", vbInformation + vbYesNo, "CMACT")
             If vbYes = OptBt2 Then
             nFicSal = FreeFile
             Open sLpt For Output As nFicSal
                Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
                    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
                    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
                    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
                    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
                    Print #nFicSal, lsCadImp
                    Print #nFicSal, ""
                    Close #nFicSal
             End If
            Loop Until OptBt2 = vbNo
            Set loPrevio = Nothing
         End If
    End If
    MsgBox "Garantia Nro: " & txtGarantia & " Liberada...!", vbInformation, "AVISO"
        '---------
       Call cmdCancelar_Click
End If

cmdCancelar.Enabled = True
cmdSalir.Enabled = True
CmdLevantar.Visible = False
CmdAceptar.Visible = False

Set Garant = Nothing
Set REP = Nothing
End Sub

Private Sub cmdBuscar_Click()
Dim Pers As COMDPersona.UCOMPersona
Set Pers = New COMDPersona.UCOMPersona
Set Pers = frmBuscaPersona.Inicio


'txtComentario = ""
Garantia = ""

MSH.ClearStructure
Marco

If Pers Is Nothing Then
txtGarantia.SetFocus

Else
   'cmdCancelar_Click
   MSHGARA.Clear
   PersCod = Pers.sPersCod
   CargaBuscar (Pers.sPersCod)
   fraComentario.Enabled = False
   fraPersonas.Enabled = False
   FraResultado.Left = 480
   FraResultado.Top = 840
   FraResultado.Visible = True
   
   txtGarantia.Enabled = False
   cmdSalir.Enabled = False
   CmdAceptar.Enabled = False
   cmdBuscar.Enabled = False
   cmdCancelar.Enabled = False

End If
Set Pers = Nothing
End Sub

Sub Marco()
With MSH
    .TextMatrix(0, 0) = "Cod Cliente"
    .TextMatrix(0, 1) = "Nombres"
    .TextMatrix(0, 2) = "Relacion"
    .ColWidth(0) = 1400
    .ColWidth(1) = 4000
    .ColWidth(2) = 1500
End With

End Sub

Private Sub cmdCancelar_Click()
txtGarantia = ""
txtComentario = ""
Garantia = ""
Moneda = ""
PersCod = ""
Realizacion = 0
MSH.Rows = 2
MSH.ClearStructure
Marco
cmdBuscar.Enabled = True
CmdLevantar.Enabled = False
CmdAceptar.Visible = False
FrameCreditos.Visible = False
txtGarantia.Enabled = True
txtGarantia.SetFocus
End Sub

Private Sub CmdLevantar_Click()

Dim GaranCred As COMDCredito.DCOMGarantia 'DGarantia
Dim rs As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Set GaranCred = New COMDCredito.DCOMGarantia
Set rs = New ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Set Rs3 = New ADODB.Recordset

'Set rs = GaranCred.RecuperaNroGaran_X_Credito(txtGarantia)
CmdLevantar.Enabled = False
'If rs.EOF And rs.BOF Then
'    MsgBox "La Garantia no tiene Creditos Asociaodos"
    
    Set Rs3 = GaranCred.RecuperaTitularGarantReal(txtGarantia)
    PersCod = Rs3!CodPers
    nTipoGar = Rs3!nGarantReal
    Set Rs2 = GaranCred.RecuperaPesona_X_Credito(PersCod)
    
    If Rs2.EOF And Rs2.BOF Then
    Else
        FrameCreditos.Left = 480
        FrameCreditos.Top = 840
        FrameCreditos.Visible = True
        MSHCreditos.ColWidth(0) = 2000
        MSHCreditos.ColWidth(1) = 4000
        FrameCreditos.Caption = "Creditos Asociados al Cliente."
        Set MSHCreditos.DataSource = Rs2
        CmdLevantar.Enabled = False
    End If
    CmdAceptar.Visible = True
    CmdAceptar.Enabled = True
    cmdBuscar.Enabled = False
    'CmdSalir3.SetFocus
    
'Else
'    MsgBox "La Garantia tiene Creditos Vigentes", vbInformation, "AVISO"
'    FrameCreditos.Left = 480
'    FrameCreditos.Top = 840
'    FrameCreditos.Visible = True
'    MSHCreditos.ColWidth(0) = 2000
'    MSHCreditos.ColWidth(1) = 4000
'    FrameCreditos.Visible = True
'    FrameCreditos.Caption = "Creditos Asociados a la Garantia"
'    Set MSHCreditos.DataSource = rs
'    CmdAceptar.Visible = False
'    CmdAceptar.Enabled = False
'    cmdBuscar.Enabled = False
'End If

Set rs = Nothing
Set Rs2 = Nothing
Set Rs3 = Nothing
Set GaranCred = Nothing

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()
   fraComentario.Enabled = True
   fraPersonas.Enabled = True
   FraResultado.Visible = False
   
   txtGarantia.Enabled = True
   cmdSalir.Enabled = True
   CmdAceptar.Enabled = True
   cmdBuscar.Enabled = True
   cmdCancelar.Enabled = True
   CmdLevantar.Enabled = False
End Sub


Private Sub CmdSalir3_Click()
FrameCreditos.Visible = False
CmdAceptar.Enabled = True
CmdLevantar.Enabled = False
End Sub




Private Sub Form_Load()
CmdAceptar.Visible = False
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub



Private Sub MSHGARA_Click()
    
'If MSHGARA.Col = 0 Then
   If MSHGARA.Row <> 0 And MSHGARA.TextMatrix(MSHGARA.Row, MSHGARA.Col) <> "" Then
   txtGarantia = MSHGARA.TextMatrix(MSHGARA.Row, 0)
   fraComentario.Enabled = True
   fraPersonas.Enabled = True
   FraResultado.Visible = False
   'MsgBox MSHGARA.TextMatrix(MSHGARA.Row, 0)

    CargaRelacion (txtGarantia)
   Marco
'   msh.SetFocus
    Me.FrameCreditos.Visible = False
   End If
'End If
End Sub
Private Sub MSHGARA_KeyPress(KeyAscii As Integer)
MSHGARA.SetFocus
End Sub

Private Sub CargaRelacion(ByVal Nrog As String)
Dim DG As COMDCredito.DCOMGarantia
Dim rs As ADODB.Recordset
Dim RsDatosG As ADODB.Recordset
Set DG = New COMDCredito.DCOMGarantia
Set rs = New ADODB.Recordset
Set RsDatosG = New ADODB.Recordset

Set rs = DG.RecuperaGaratRealPersoRelac(Nrog)
CmdLevantar.Enabled = False
MSH.Clear
If rs.EOF And rs.BOF Then
 MsgBox "No existe esa Garantia", vbInformation, "AVISO"
 txtGarantia.SetFocus
Else
 Set MSH.DataSource = rs
 CmdLevantar.Visible = True
 CmdLevantar.Enabled = True
 cmdCancelar.Enabled = True
 cmdSalir.Enabled = True
End If

txtComentario = ""
Marco
Set DG = Nothing
Set rs = Nothing
Set RsDatosG = Nothing

End Sub
Private Sub CargaBuscar(ByVal CodPers As String)
Dim DG As COMDCredito.DCOMGarantia
Dim rs As ADODB.Recordset

Set DG = New COMDCredito.DCOMGarantia
Set rs = New ADODB.Recordset
Set rs = DG.RecuperaGarantiasRealPersona(CodPers)


Set MSHGARA.DataSource = rs
MSHGARA.TextMatrix(0, 0) = "NroGarantia"
MSHGARA.TextMatrix(0, 1) = "Cliente"
MSHGARA.ColWidth(0) = 1000
MSHGARA.ColWidth(1) = 5000

If rs.EOF And rs.BOF Then
    'CmdSalir2.SetFocus
Else
   cmdBuscar.Enabled = True
   cmdCancelar.Enabled = True
   CmdLevantar.Enabled = False
   fraComentario.Enabled = True
End If

Set DG = Nothing
Set rs = Nothing
End Sub

Private Sub txtGarantia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(Trim(txtGarantia)) Then
    txtGarantia = Format(txtGarantia, "00000000")
    MSH.Rows = 2
    Marco
    CargaRelacion (txtGarantia)
End If
End Sub


