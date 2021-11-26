VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmPersReporte 
   Caption         =   "Reportes de Personas"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Height          =   375
      Left            =   8070
      TabIndex        =   1
      Top             =   6465
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   5568
      TabIndex        =   0
      Top             =   75
      Width           =   4584
      Begin MSMask.MaskEdBox txtfecFin 
         Height          =   276
         Left            =   3144
         TabIndex        =   4
         Top             =   312
         Visible         =   0   'False
         Width           =   972
         _ExtentX        =   1693
         _ExtentY        =   476
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecini 
         Height          =   252
         Left            =   912
         TabIndex        =   5
         Top             =   312
         Visible         =   0   'False
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Fecha Fin:"
         Height          =   252
         Left            =   2208
         TabIndex        =   3
         Top             =   336
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label lblFecINi 
         Caption         =   "Fecha Ini:"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   336
         Visible         =   0   'False
         Width           =   732
      End
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   984
      Top             =   5376
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersReporte.frx":0000
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersReporte.frx":0352
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersReporte.frx":06A4
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersReporte.frx":09F6
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TVRep 
      Height          =   6552
      Left            =   240
      TabIndex        =   6
      Top             =   216
      Width           =   4716
      _ExtentX        =   8308
      _ExtentY        =   11562
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imglstFiguras"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPersReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCad As NCredReporte
Dim sCad1 As String
 
'  Clase para ver vista preliminar
Dim P As previo.clsPrevio

'Dim RepValores As New n

Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String

Private Sub cmdImprimir_Click()

Dim strs As String
Dim oRep As NpersReporte
Set oRep = New NpersReporte
   
   Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
    Case gColCredRepIngxPagoCred
   
    Case 808101
    
      If IsDate(txtFecini.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            txtFecini.SetFocus
            Exit Sub
        End If

       txtFecini.Visible = True
       lblFecINi.Visible = True
       
      oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
      strs = oRep.ReporteDeClientes(txtFecini)
      
        Case 808102
    
      If IsDate(txtFecini.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            txtFecini.SetFocus
            Exit Sub
        End If
      
       txtFecini.Visible = True
       lblFecINi.Visible = True
       
      oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
      strs = oRep.ReporteDeClientesMod(txtFecini)
      
    End Select
   
    If strs <> "" Then
    Set P = New previo.clsPrevio
         P.Show strs, "Reportes de Creditos", True, 66, gEPSON
    Set P = Nothing
    End If

    txtFecini.Visible = False
    txtfecFin.Visible = False
    lblFecINi.Visible = False
    lblFecFin.Visible = False
    
End Sub


Private Sub Form_Load()
    Dim I As Integer
    Dim lsCadena As String
    Dim lsCad As String
    Dim oPersona As UPersona
    Dim oPrevio As previo.clsPrevio
    Set oPrevio = New previo.clsPrevio
    Dim oRep As NpersReporte
    Set oRep = New NpersReporte
 '   Dim oRepr As NPigRemate
     
Dim lnNumRep As Integer
Dim lsCadenaBuscar As String
Dim lsRep() As String

LlenaArbol
   
End Sub

Private Sub LlenaArbol()
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim lsTipREP As String

'Para filtrar el tipo de reporte de la tabla OpeTipo
    lsTipREP = "808"
    
    Set clsGen = New DGeneral
    
    'ARCV 20-07-2006
    'Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
    Set rsUsu = clsGen.GetOperacionesUsuario_NEW(lsTipREP, , gRsOpeRepo)

    
    Set clsGen = Nothing
      
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("cOpeCod")
        sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
        Select Case rsUsu("nOpeNiv")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = TVRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
End Sub

Private Sub TVRep_Click()

' Se inserto una constante gColPR-epBovedaLotes en el proyecto
     txtFecini = gdFecSis
    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
       Case "808101"
       txtFecini.Visible = True
       lblFecINi.Visible = True
       CmdImprimir.Enabled = True
       txtFecini.SetFocus
    Case "808102"
       txtFecini.Visible = True
       lblFecINi.Visible = True
       CmdImprimir.Enabled = True
       txtFecini.SetFocus
     End Select
     
End Sub


Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFecini.Visible = True Then
        If IsDate(txtfecFin.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            txtfecFin.SetFocus
            Exit Sub
        End If
    CmdImprimir.SetFocus
Else
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            txtfecFin.SetFocus
            Exit Sub

    End If
End If
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFecini.Visible = True Then
        If IsDate(txtFecini.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            txtFecini.SetFocus
            Exit Sub
        End If
    CmdImprimir.SetFocus
Else
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            txtFecini.SetFocus
            Exit Sub

    End If
End If

End Sub
