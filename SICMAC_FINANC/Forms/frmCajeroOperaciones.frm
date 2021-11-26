VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajeroOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7140
   ClientLeft      =   3255
   ClientTop       =   960
   ClientWidth     =   7635
   Icon            =   "frmCajeroOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmMoneda 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15
      TabIndex        =   4
      Top             =   150
      Width           =   1275
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &N."
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "M. &E."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5070
      TabIndex        =   3
      Top             =   6630
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3735
      TabIndex        =   2
      Top             =   6630
      Width           =   1335
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Seleccione Operación"
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
      Height          =   6435
      Left            =   1410
      TabIndex        =   0
      Top             =   45
      Width           =   6435
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   1230
         Top             =   5685
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
               Picture         =   "frmCajeroOperaciones.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":065C
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":09AE
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":0D00
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwOperacion 
         Height          =   6075
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10716
         _Version        =   393217
         LabelEdit       =   1
         Style           =   5
         SingleSel       =   -1  'True
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmCajeroOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EjecutaOperacion(ByVal nOperacion As CaptacOperacion)
Select Case nOperacion
    Case gAhoApeEfec
        'frmCapAperturas.Inicia gCapAhorros, nOperacion
    Case gAhoApeChq
        'frmCapAperturas.Inicia gCapAhorros, nOperacion
    Case gAhoApeTransf
        'frmCapAperturas.Inicia gCapAhorros, nOperacion
End Select
End Sub

Private Sub Form_Load()
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String
Dim nodOpe As Node

Me.Caption = "Cajero - Operaciones"
Set clsGen = New DGeneral
Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, "4", MatOperac, NroRegOpe)
Set clsGen = Nothing
Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvwOperacion.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwOperacion.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    rsUsu.MoveNext
Loop
rsUsu.Close
Set rsUsu = Nothing
End Sub

Private Sub tvwOperacion_DblClick()
Dim nodOpe As Node
Set nodOpe = tvwOperacion.SelectedItem
EjecutaOperacion CLng(nodOpe.Tag)
Set nodOpe = Nothing
End Sub

Private Sub tvwOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim nodOpe As Node
    Set nodOpe = tvwOperacion.SelectedItem
    EjecutaOperacion CLng(nodOpe.Tag)
    Set nodOpe = Nothing
End If
End Sub
