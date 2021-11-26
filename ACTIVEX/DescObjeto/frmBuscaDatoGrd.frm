VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBuscaDatoGrd 
   Caption         =   "Busqueda de Cuentas Contables"
   ClientHeight    =   1095
   ClientLeft      =   2385
   ClientTop       =   3795
   ClientWidth     =   7200
   Icon            =   "frmBuscaDatoGrd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCtaContCod 
      BackColor       =   &H00F0FFFF&
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
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame frabusca 
      Caption         =   "Buscar por ...."
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
      Height          =   1050
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2010
      Begin VB.OptionButton optOpcion 
         Caption         =   "&Cuenta Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   1800
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "&Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1605
      End
   End
   Begin VB.TextBox txtCtaContDesc 
      BackColor       =   &H00F0FFFF&
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
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc AdoBuscaCta 
      Height          =   330
      Left            =   780
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBuscaDatoGrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOrden As Integer
Dim lOkey As Boolean
Dim lsValBusqueda As String
Dim lsTitulo As String
Dim AdoSeek As ADODB.Recordset
Dim IniSeek As Long

Dim nColCod As Integer
Dim nColDes As Integer
Public Sub Inicia(prs As ADODB.Recordset, nOrden As Integer, sTitulo As String, Optional lnColCod As Integer = 0, Optional lnColDes As Integer = 1)
' nOrden    => 0 Codigo
'           => 1 Descripcion
nColCod = lnColCod
nColDes = lnColDes

lnOrden = nOrden
lsTitulo = sTitulo
Set AdoSeek = prs
Me.Show 1
End Sub

Private Sub Form_Activate()
optOpcion(lnOrden) = True
End Sub

Private Sub Form_Load()
    CentraForm Me
    IniSeek = 1
    If lsTitulo <> "" Then
        Me.Caption = "Búsqueda de " & lsTitulo
        optOpcion(0).Caption = "&" & lsTitulo
    End If
End Sub

Private Sub txtCtaContCod_KeyPress(KeyAscii As Integer)
Dim Criterio As String
KeyAscii = NumerosEnteros(KeyAscii)
lOkey = False
If KeyAscii = 27 Then
   Unload Me
End If
If KeyAscii = 13 Then
   If Len(Trim(txtCtaContCod)) = 0 Then
      Unload Me
      Exit Sub
   End If
   If BuscaDato(DefCriterio(nColCod, nColDes), AdoSeek, IIf(nOrden = 0, 1, AdoSeek.AbsolutePosition + 1), True) Then
      lOkey = True
      Unload Me
   End If
End If
End Sub
Private Sub txtCtaContDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii, True)
lOkey = False
If KeyAscii = 27 Then
   Unload Me
End If
If KeyAscii = 13 Then
   If Len(Trim(txtCtaContDesc)) = 0 Then
      Unload Me
      Exit Sub
   End If
   If BuscaDato(DefCriterio(), AdoSeek, IIf(nOrden = 0, 1, AdoSeek.AbsolutePosition + 1), True) Then
      lOkey = True
      Unload Me
   End If
End If
End Sub

Private Sub optOpcion_Click(Index As Integer)
lnOrden = Index
Select Case Index
 Case 0
        txtCtaContCod.Visible = True
        txtCtaContDesc.Visible = False
        txtCtaContCod.SetFocus
        txtCtaContCod.Text = ""
 Case 1
        txtCtaContCod.Visible = False
        txtCtaContDesc.Visible = True
        txtCtaContDesc.SetFocus
        txtCtaContCod.Text = ""
End Select
End Sub

Private Function DefCriterio(Optional nColCod As Integer = 0, Optional nColDes As Integer = 1) As String
Dim ss As String
   DefCriterio = IIf(optOpcion(0), AdoSeek(nColCod).Name, AdoSeek(nColDes).Name) & " LIKE '"
   DefCriterio = DefCriterio & IIf(optOpcion(0).Value, txtCtaContCod.Text, txtCtaContDesc.Text) & "*'"
End Function

Public Property Get nOrden() As Integer
nOrden = lnOrden
End Property

Public Property Let nOrden(ByVal vNewValue As Integer)
lnOrden = vNewValue
End Property

Public Property Get lOk() As Boolean
lOk = lOkey
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
lOkey = vNewValue
End Property

