VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmProtesto 
   Caption         =   "Mantenimiento de Protesto"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   3300
      Width           =   4725
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   2610
         TabIndex        =   4
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1260
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   5636
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
End
Attribute VB_Name = "FrmProtesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsProtesto As ADODB.Recordset

Private Sub CmdAceptar_Click()
    If Not rsProtesto.EOF And Not rsProtesto.BOF Then
        ActualizarProtesto
    End If
End Sub

Private Sub CmdCancelar_Click()
    Form_Load
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Sub ConfigurarGD()
    Set rsProtesto = New ADODB.Recordset
        
     With rsProtesto.Fields
        .Append "CodAgen", adChar, 2
        .Append "DesAgen", adVarChar, 25
        .Append "Monto", adCurrency
     End With
     
    rsProtesto.Open
    Set DG.DataSource = rsProtesto
    
    DG.Columns(0).Caption = "Cod.Agen"
    DG.Columns(1).Caption = "Agencia"
    DG.Columns(2).Caption = "Monto"
    DG.Columns(0).Width = 1000
    DG.Columns(2).NumberFormat = "#0.00"
    
    DG.AllowUpdate = True
    DG.AllowDelete = False
    DG.AllowAddNew = False
End Sub

Sub CargarProtesto()
    Dim rs As ADODB.Recordset
    Dim objProtesto As DProtesto
    On Error GoTo ErrHandler
        Set objProtesto = New DProtesto
        Set rs = objProtesto.ObtAgenProtesto
        Set pbjprotesto = Nothing
            
        Do Until rs.EOF
            rsProtesto.AddNew
            rsProtesto(0) = rs!cAgeCod
            rsProtesto(1) = rs!cAgeDescripcion
            rsProtesto(2) = Format(rs!Monto, "#0.00")
            rsProtesto.Update
            rs.MoveNext
        Loop
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not objProtesto Is Nothing Then Set objProtesto = Nothing
    MsgBox "Error al cargar las agencias", vbInformation, "AVISO"
End Sub

Private Sub DG_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = 2 Then
        If KeyAscii = 13 Then
             If IsNumeric(DG.Columns(ColIndex).value) Then
                 
             Else
                MsgBox "El valor debe ser numerico", vbInformation, "AVISO"
             End If
        Else
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub Form_Load()
    ConfigurarGD
    CargarProtesto
End Sub

Sub ActualizarProtesto()
    Dim nValor As Double
    Dim cCodAgen As String
    Dim objProtesto As DProtesto
    Dim bValor As Boolean
    Dim rsx As ADODB.Recordset
    On Error GoTo ErrHandler
        Set rsx = rsProtesto.Clone
        Do Until rsx.EOF
            Set objProtesto = New DProtesto
            bValor = objProtesto.ActualizaProtesto(rsx!CodAgen, rsx!Monto)
            Set objProtesto = Nothing
            If bValor = False Then
                MsgBox "Se ha producido un error", vbInformation, "AVISO"
                Exit Do
            End If
            rsx.MoveNext
        Loop
        Set rsx = Nothing
        MsgBox "Se actualizo correctamente", vbInformation, "AVISO"
        Form_Load
    Exit Sub
ErrHandler:
    If Not objProtesto Is Nothing Then Set objProtesto = Nothing
    MsgBox "Error el actualizarProtesto", vbInformation, "AVISO"
End Sub
