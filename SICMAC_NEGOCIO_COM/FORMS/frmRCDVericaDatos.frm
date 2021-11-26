VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRCDVericaDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RCD Verifica Datos"
   ClientHeight    =   4365
   ClientLeft      =   4065
   ClientTop       =   3105
   ClientWidth     =   6345
   Icon            =   "frmRCDVericaDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Verificación de Datos"
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6105
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   300
         Left            =   4800
         TabIndex        =   5
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdComprobar 
         Caption         =   "&Comprobar "
         Height          =   300
         Left            =   4800
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
      Begin MSComctlLib.ImageList Imagenes 
         Left            =   5280
         Top             =   2160
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
               Picture         =   "frmRCDVericaDatos.frx":000C
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRCDVericaDatos.frx":035E
               Key             =   "Hijo"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwReporte 
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5953
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "Imagenes"
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label lblGlosa 
         AutoSize        =   -1  'True
         Height          =   1080
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   1200
         WordWrap        =   -1  'True
      End
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   330
      Left            =   855
      TabIndex        =   3
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "frmRCDVericaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnRepoSelec As Long

Private Sub cmdComprobar_Click()
Dim lsCad As String
Dim lsServerCons As String
Dim loRep As COMNCredito.NCOMRCD 'nRcdReportes
Dim rs As New ADODB.Recordset

Set loRep = New COMNCredito.NCOMRCD 'nRcdReportes
lsServerCons = loRep.GetServerConsol
If Len(ValidaFecha(txtFecha.Text)) >= 1 Then
    MsgBox ValidaFecha(txtFecha.Text), vbInformation, "AVISO"
    Exit Sub
End If

If loRep.nExisteTabla("RCDvc" & Format(gdFecDataFM, "yyyymm") & "01", lsServerCons) = False Then
   MsgBox "No existe la Tabla " & "RCDvc" & Format(gdFecDataFM, "yyyymm") & "01"
   Exit Sub
End If
If "01/01/2000" <= Format(txtFecha.Text, "DD/MM/YYYY") And Format(txtFecha.Text, "DD/MM/YYYY") <= gdFecDataFM Then
    Select Case fnRepoSelec
        Case 179201: Set rs = loRep.nRepo179201_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179202: Set rs = loRep.nRepo179202_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179203: Set rs = loRep.nRepo179203_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179204: Set rs = loRep.nRepo179204_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179205: Set rs = loRep.nRepo179205_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179206: Set rs = loRep.nRepo179206_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179207: Set rs = loRep.nRepo179207_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179208: Set rs = loRep.nRepo179208_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179209: Set rs = loRep.nRepo179209_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179210: Set rs = loRep.nRepo179210_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179211: Set rs = loRep.nRepo179211_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179212: Set rs = loRep.nRepo179212_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179213: Set rs = loRep.nRepo179213_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179214: Set rs = loRep.nRepo179214_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179215: Set rs = loRep.nRepo179215_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179216: Set rs = loRep.nRepo179216_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case 179217: Set rs = loRep.nRepo179217_(lsServerCons, Format(txtFecha, "YYYYMM"))
        Case Else: Exit Sub
    End Select
    Call FrmMostrarDatos.Inicio(rs)
Else
    MsgBox "Ingrese una Fecha Correcta", vbInformation, "AVISO"
End If
Set loRep = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargaMenu
    txtFecha = gdFecData
End Sub

Private Sub CargaMenu()
Dim clsGen As DGeneral   'COMDConstSistema.DCOMGeneral ARCV 25-10-2006
Dim rsUsu As Recordset
Dim sOperacion As String
Dim sOpeCod As String
Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String
Dim nodOpe As Node
Dim lsTipREP As String
lsTipREP = "1792"
Set clsGen = New DGeneral 'COMDConstSistema.DCOMGeneral
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
         Set nodOpe = tvwReporte.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwReporte.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
    End Select
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub

Private Sub tvwReporte_Click()
Dim NodRep  As Node
Dim lsDesc As String
Set NodRep = tvwReporte.SelectedItem
If NodRep Is Nothing Then
   Exit Sub
End If
lsDesc = Mid(NodRep.Text, 8, Len(NodRep.Text) - 7)
fnRepoSelec = CLng(NodRep.Tag)
End Sub
