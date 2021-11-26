VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmArchivos 
   Caption         =   "Archivos"
   ClientHeight    =   5805
   ClientLeft      =   1515
   ClientTop       =   1905
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   7770
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5175
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6300
      TabIndex        =   3
      Top             =   5175
      Width           =   1335
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid grdFile 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   675
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCon As New Connection
Dim rs   As ADODB.Recordset

Dim sSql As String
Dim vFile As String

Private Sub cmdEliminar_Click()
Dim nPos As Variant
nPos = AdoFile.Recordset.Bookmark
AdoFile.Recordset.Delete
AdoFile.Refresh
AdoFile.Recordset.Bookmark = nPos
End Sub

Private Sub cmdSalir_Click()
'AdoFile.Recordset.Close
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Dim oConect As New DConecta
oConect.AbreConexion gsConnection
oConect.CierraConexion
Set oConect = Nothing
Set rs = New ADODB.Recordset
    oCon.Open gsConnection
    oCon.Execute "Set DateFormat mdy"
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rs
oCon.Close: Set oCon = Nothing
End Sub

Private Sub txtFile_Keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
   vFile = Trim(txtFile.Text)
   CargaArchivo
End If
End Sub

Private Sub CargaArchivo()
On Error GoTo errorCarga
If InStr(UCase(sSql), "SELECT") = 0 Then
   sSql = "select * from " & vFile
Else
   sSql = vFile
End If

    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    rs.Open sSql, oCon, adOpenStatic, adLockOptimistic, adCmdText
    Set grdFile.DataSource = rs
Exit Sub
errorCarga:
   MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
