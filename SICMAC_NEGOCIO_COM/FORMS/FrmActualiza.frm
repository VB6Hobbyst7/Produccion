VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PARA RETIRO DE PLANILLAS"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Retiro"
      Height          =   450
      Left            =   3045
      TabIndex        =   1
      Top             =   5280
      Width           =   1125
   End
   Begin SICMACT.FlexEdit flxPlanilla 
      Height          =   4470
      Left            =   945
      TabIndex        =   0
      Top             =   330
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7885
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   -1
      RowHeight0      =   240
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim ssql As String, rstemp As Recordset, ocon As DConecta
  Set rstemp = New Recordset
  Set ocon = New DConecta
  
  ocon.AbreConexion
  
  ssql = "select  CCTACOD=mc.cctacod,NMONTO=sum(mc.nmonto),mc.copecod from mov m " _
         & " join movcap mc on mc.nmovnro=m.nmovnro " _
         & " where m.cmovnro like '20041207%EELA' and mc.copecod='200204' " _
         & " group by mc.cctacod,mc.copecod " _
         & "   order by mc.cctacod  "
  
  rstemp.CursorLocation = adUseClient
  rstemp.Open ssql, ocon, adOpenStatic, adLockReadOnly, adCmdText
  If Not rstemp.EOF Then
           Set flxPlanilla.Recordset = rstemp
            
  End If
  ocon.CierraConexion
  Set ocon = Nothing
End Sub
