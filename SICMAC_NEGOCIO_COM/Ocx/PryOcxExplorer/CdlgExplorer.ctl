VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl OcxCdlgExplorer 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "CdlgExplorer.ctx":0000
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   555
      Index           =   0
      Left            =   1545
      TabIndex        =   3
      Top             =   570
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   979
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"CdlgExplorer.ctx":0312
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   15
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   510
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox PictPresent 
      Height          =   405
      Left            =   30
      Picture         =   "CdlgExplorer.ctx":03E7
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   30
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      Height          =   390
      Index           =   0
      Left            =   1620
      ScaleHeight     =   330
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "OcxCdlgExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum TVentana
    Normal = 0
    Con_Imagen = 1
    Con_Texto = 2
End Enum
Private sTipoVentana As TVentana
Private sFiltro As String
Public nHwd As Long

Private nHeigth As Long
Public Function Ruta() As String
    Ruta = CDlg.FileName
End Function
Public Sub Show()
  Set CDlg = New CdlgExplorer

  'CDlg.Filter = "Archivos de Texto|*.txt;*.rtf|Graphic Files|*.bmp;*.gif;*.jpg;*.ico;*.wmf|All files|*.*"
  'CDlg.Filter = "Archivos Graficos|*.jpg|Todos los Archivos|*.*"
  CDlg.Filter = sFiltro
  CDlg.hOwner = nHwd
  CDlg.Height = nHeigth
  Select Case sTipoVentana
    Case Con_Imagen
        Load Picture1(1)
        CDlg.Left = 100
        CDlg.Top = 150
        Set Pict = Picture1(1)
        CDlg.Width = 600
        CDlg.TipoVentana = Con_Imagen
    Case Con_Texto
        Load RichTextBox1(1)
        CDlg.Left = 200
        CDlg.Top = 150
        Set rtb = RichTextBox1(1)
        CDlg.Width = 450
        CDlg.TipoVentana = Con_Texto
    Case Normal
        CDlg.Left = 200
        CDlg.Top = 150
        CDlg.Width = 430
        CDlg.TipoVentana = Normal
  End Select
  CDlg.ShowOpen
  If RichTextBox1.Count > 1 Then Unload RichTextBox1(1)
  If Picture1.Count > 1 Then Unload Picture1(1)
  'PictPresent.SetFocus
  'If CDlg.FileName <> "" Then Me.Caption = CDlg.FileName
End Sub
Public Property Get TipoVentana() As TVentana
    TipoVentana = sTipoVentana
End Property
Public Property Let TipoVentana(ByVal vNewValue As TVentana)
    sTipoVentana = vNewValue
End Property

Public Property Get Filtro() As String
    Filtro = sFiltro
End Property

Public Property Let Filtro(ByVal vNewValue As String)
    sFiltro = vNewValue
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Filtro = PropBag.ReadProperty("Filtro", "Todos los Archivos|*.*")
 TipoVentana = PropBag.ReadProperty("TipoVentana", 0)
 Altura = PropBag.ReadProperty("Altura", 330)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Filtro", Filtro, "Todos los Archivos|*.*"
    PropBag.WriteProperty "TipoVentana", TipoVentana, 0
    PropBag.WriteProperty "Altura", Altura, 330
End Sub

Public Property Get Altura() As Long
    Altura = nHeigth
End Property

Public Property Let Altura(ByVal pnAlto As Long)
    nHeigth = pnAlto
End Property

