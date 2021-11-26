VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCFDuplicado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Duplicado de Carta Fianza"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Lista de Documentos de Cartas Fianzas"
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   6375
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   4680
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin MSComctlLib.TreeView tvwLista 
         Height          =   2475
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4366
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   4080
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCFDuplicado.frx":0000
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCFDuplicado.frx":031A
               Key             =   "Hijo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin SICMACT.ActXCodCta CodCta 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "C F. Nro:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   6120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lblDocJuridico 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lblDocNat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         Width           =   4905
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Juridico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   9
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado CF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   8
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Nat. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Datos del Titular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1425
      End
   End
End
Attribute VB_Name = "FrmCFDuplicado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnreposelec As Integer
Dim ban As Integer
Private Function Carta(ByRef CTA As String) As Boolean
Dim lrG As NCartaFianzaValida
Dim Rs As ADODB.Recordset
Carta = False
Set lrG = New NCartaFianzaValida
If Len(CTA) = 18 Then
    Set Rs = CargaCta(CTA, lrG)
    If Rs.EOF And Rs.BOF Then
     ClearForm
     MsgBox "NO EXISTE CARTA FIANZA" & CTA, vbInformation, "AVISO"
    Else
    Carta = True
    lblCodigo = Rs!Codigo
    lblNombre = Rs!Nombre
    lblDocNat = IIf(IsNull(Rs!DocDni), "", Rs!DocDni)
    lblDocJuridico = IIf(IsNull(Rs!DocRuc), "", Rs!DocRuc)
    lblEstado = IIf(IsNull(Rs!Estado), "", Rs!Estado)
    ban = 1
    End If
Else
    ClearCta
    ClearForm
    MsgBox "Falta Nros en la Cuenta", vbInformation, "AVISO"
    ban = 0
    CodCta.SetFocusAge
End If

End Function

Private Function CargaCta(ByRef CTA As String, ByRef lg As NCartaFianzaValida) As ADODB.Recordset
Set CargaCta = lg.RecuperaDatosT(CTA)
End Function

Private Sub CmdBuscar_Click()
Dim Pers As UPersona
Set Pers = New UPersona
Set Pers = frmBuscaPersona.Inicio
MsgBox Pers.sPersCod

If IsNull(Pers.sPersCod) Then
    
Else
    FrmCFcarga.Inicio (Pers.sPersCod)
    FrmCFcarga.Show 1

End If

End Sub

Private Sub CmdCancelar_Click()
ClearForm
ClearCta
ban = 0
End Sub
Private Sub ClearForm()
lblCodigo = ""
lblEstado = ""
lblNombre = ""
lblDocNat = ""
lblDocJuridico = ""
ban = 0
End Sub
Private Sub ClearCta()
CodCta.Age = ""
CodCta.CMAC = ""
CodCta.Cuenta = ""
CodCta.Prod = ""
End Sub
Private Sub CmdImprimir_Click()
Dim loRep As NCartaFianzaReporte
Dim lsCadImp As String
Dim loPrevio As Previo.clsPrevio
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel

'ClearForm
If ban = 1 And (Len(CodCta.NroCuenta) = 18) Then
 If (fnreposelec = 2 Or fnreposelec = 3) Then
 Set loRep = New NCartaFianzaReporte
 loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
 lsCadImp = loRep.nRepoDuplicado(CodCta.NroCuenta, fnreposelec)
 lsDestino = "P"
 If lsDestino = "P" Then
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New Previo.clsPrevio
            loPrevio.Show lsCadImp, "Reporte Duplicados de Carta Fianza", True
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte ", vbInformation, "Aviso"
    End If
 End If
 Else
   MsgBox "Elija un de la opcines", vbInformation, "Aviso"
 End If
End If

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub
Private Sub Carga()

Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim nodOpe As Node
Dim nodOpe2 As Node
Dim nodOpe3 As Node


        sOperacion = "REPORTES DE CARTAS FIANZAS"
        sOpePadre = "P1"
         Set nodOpe = tvwLista.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
         nodOpe.Tag = "1"
         sOperacion = "SOLICITUD"
         sOpeHijo = "H1"
         Set nodOpe2 = tvwLista.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
         nodOpe2.Tag = "2"
         sOperacion = "SUGERENCIA"
         sOpeHijo = "H2"
         Set nodOpe3 = tvwLista.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
         nodOpe3.Tag = "3"

End Sub

Private Sub CodCta_KeyPress(KeyAscii As Integer)
ClearForm
'ban = 0
If KeyAscii = 13 And Len(CodCta.NroCuenta) = 18 Then
   If Carta(CodCta.NroCuenta) Then
   Else
    'MsgBox "No Existe Carta Fianza", vbInformation, "AVISO"
    ClearCta
   End If
Else
   MsgBox "Nro de Cta Imcompleto", vbInformation, "AVISO"
   ClearCta
End If
End Sub

Private Sub Form_Load()
ClearForm
Carga
'CodCta.SetFocusCuenta
End Sub


Private Sub tvwLista_NodeClick(ByVal Node As MSComctlLib.Node)
Dim NodRep  As Node
Dim lsDesc As String

Set NodRep = tvwLista.SelectedItem

If NodRep Is Nothing Then
   Exit Sub
End If
'ban = 0
lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnreposelec = CLng(NodRep.Tag)

Select Case fnreposelec
    Case 2
        'ban = 1
    Case 3
        'ban = 1
Case Else
End Select

Set NodRep = Nothing
End Sub
