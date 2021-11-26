VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBusqDiversa 
   Caption         =   "Operaciones : Búsqueda de "
   ClientHeight    =   3915
   ClientLeft      =   945
   ClientTop       =   2325
   ClientWidth     =   8340
   Icon            =   "frmBusqDiversa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grdObjs 
      Height          =   2835
      Left            =   120
      TabIndex        =   8
      Top             =   300
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   5001
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1725.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5760
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1260
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Asignar Objeto a Cuenta"
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Frame Frame2 
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
      Height          =   675
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3405
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2850
         Picture         =   "frmBusqDiversa.frx":030A
         TabIndex        =   1
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox Text1 
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
         Left            =   2880
         TabIndex        =   7
         Top             =   210
         Width           =   285
      End
      Begin VB.TextBox txtObj 
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
         Left            =   810
         TabIndex        =   0
         Top             =   210
         Width           =   2115
      End
      Begin VB.Label Label6 
         Caption         =   "Código "
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label lblObj 
      Caption         =   "Datos"
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   2355
   End
End
Attribute VB_Name = "frmBusqDiversa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCod As String, sDesc As String
Dim rs As ADODB.Recordset
Dim grdActivo As Boolean, lPressMouse As Boolean

Dim Ok As Boolean
Dim lUltNivel As Boolean
Dim sTitulo As String
Dim nOrden As Integer

Public Sub inicio(prs As ADODB.Recordset, Optional psTitulo As String = "Búsqueda de Datos")
Set rs = prs
sTitulo = psTitulo
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
If Not rs.EOF And Not rs.BOF Then
    sCod = rs(0)
    sDesc = rs(1)
    If lUltNivel Then
       rs.MoveNext
       If Not rs.EOF Then
         If Mid(rs(0), 1, Len(sCod)) = sCod Then
            rs.MovePrevious
            MsgBox "Dato no es última Instancia", vbInformation, "¡Aviso!"
            Exit Sub
         End If
       End If
       rs.MovePrevious
    End If
    Ok = True
Else
    Ok = False
End If
Unload Me
End Sub

Private Sub CmdBuscar_Click()
If rs.RecordCount > 0 Then
   grdObjs.MarqueeStyle = dbgNoMarquee
   sTitulo = rs.Fields(0).Name
   frmBuscaDatoGrd.Inicia rs, nOrden, sTitulo
   nOrden = frmBuscaDatoGrd.nOrden
   txtObj.Text = rs(0)
   grdObjs.MarqueeStyle = dbgHighlightRow
   txtObj.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrBuscar
Ok = False
lblObj.Caption = sTitulo
Me.Caption = Me.Caption & lblObj.Caption
CentraForm Me
grdObjs.ScrollBars = dbgBoth
Set grdObjs.DataSource = rs
If grdObjs.Columns.Count = 2 Then
   grdObjs.Columns(0).Width = 1800
   grdObjs.Columns(1).Width = 5700
End If
If rs.EOF And rs.BOF Then
   txtObj.Enabled = False
   cmdBuscar.Enabled = False
   cmdAceptar.Enabled = False
Else
End If
Exit Sub
ErrBuscar:
   Err.Raise Err.Number, "Busqueda de Datos", Err.Description
End Sub

Private Sub grdObjs_DblClick()
If grdActivo Then
   cmdAceptar_Click
End If
End Sub

Private Sub grdObjs_GotFocus()
grdObjs.MarqueeStyle = dbgHighlightRow
grdActivo = True
End Sub

Private Sub grdObjs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 34 And Shift = 2 Then
   rs.MoveLast
End If
If KeyCode = 33 And Shift = 2 Then
   rs.MoveFirst
End If

End Sub

Private Sub grdObjs_KeyPress(KeyAscii As Integer)
If grdActivo Then
   If KeyAscii = 13 Then
      cmdAceptar_Click
      KeyAscii = 0
   End If
End If
End Sub
Private Sub grdObjs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
   If grdActivo Then
      txtObj.Text = rs(0)
   End If
End If
End Sub
Private Sub grdObjs_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lPressMouse = True
End Sub

Private Sub grdObjs_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lPressMouse = True
End Sub

Private Sub grdObjs_LostFocus()
grdActivo = False
End Sub

Private Sub grdObjs_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdActivo Or lPressMouse Then
   If Not rs.EOF Then
      txtObj.Text = rs(0)
   End If
   lPressMouse = False
End If
End Sub

Private Sub txtObj_Change()
Dim Criterio As String
Dim nPos As Long
If rs.EOF Then rs.MoveFirst

nPos = rs.Bookmark
If Len(Trim(txtObj.Text)) > 0 Then
   Criterio = rs(0).Name & " LIKE '" & txtObj.Text & "*'"
   rs.Find Criterio, , , 1
   If rs.EOF And rs.BOF Then
      rs.Bookmark = nPos
   End If
End If
End Sub

Private Sub txtObj_GotFocus()
txtObj.SelLength = Len(txtObj)
End Sub

Private Sub txtObj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not rs.EOF And Not rs.BOF Then
        txtObj.Text = rs(0)
    End If
   cmdAceptar.SetFocus
End If
End Sub


Public Property Get pCod() As String
pCod = sCod
End Property

Public Property Let pCod(ByVal vNewValue As String)
sCod = vNewValue
End Property

Public Property Get pDesc() As String
pDesc = sDesc
End Property

Public Property Let pDesc(ByVal vNewValue As String)
sDesc = vNewValue
End Property

Public Property Get lOk() As Boolean
lOk = Ok
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
Ok = vNewValue
End Property


Public Property Get lvUltNivel() As Boolean
lvUltNivel = lUltNivel
End Property

Public Property Let lvUltNivel(ByVal vNewValue As Boolean)
lUltNivel = vNewValue
End Property
