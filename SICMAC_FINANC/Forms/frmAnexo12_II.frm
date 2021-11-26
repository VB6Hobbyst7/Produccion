VERSION 5.00
Begin VB.Form frmAnexo12_II 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANEXO 12-II"
   ClientHeight    =   1260
   ClientLeft      =   5610
   ClientTop       =   5580
   ClientWidth     =   4380
   Icon            =   "frmAnexo12_II.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4380
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   1155
   End
   Begin VB.Frame fraMes 
      Caption         =   "Periodo"
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
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   630
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmAnexo12_II.frx":030A
         Left            =   2280
         List            =   "frmAnexo12_II.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   1710
         TabIndex        =   4
         Top             =   390
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmAnexo12_II"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdGenerar_Click()
    If Not ValidaAnio Then
        MsgBox "Debe Ingresar una Fecha Correcto", vbInformation, "AVISO!"
        Me.txtAnio.SetFocus
        Exit Sub
    End If
    Dim dFecha As Date
    Dim rs As New Recordset
    Dim oDatosAnexo As New DbalanceCont
    Dim oImpAnexo As New NContImprimir
    
    dFecha = DateAdd("d", -1, DateAdd("m", 1, CDate(("01/" + Format(Me.CboMes.ListIndex + 1, "00") + "/" + Me.txtAnio))))
    
    If MsgBox("Se va ha Generar el Anexo 12-II", vbYesNo, "AVISO!") = vbYes Then
        
        Set rs = oDatosAnexo.ObtenerDatosAnexo12_II(Me.txtAnio.Text, Format(Me.CboMes.ListIndex + 1, "00"))
        
        If Not (rs.BOF And rs.EOF) Then
            oImpAnexo.GeneraExcelAnexo12_II rs, dFecha, App.path
        Else
            MsgBox "No se encontraron Datos con esta Fecha", vbInformation, "AVISO!"
        End If
    
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.txtAnio = Year(gdFecSis)
    Me.CboMes.ListIndex = Month(DateAdd("m", -1, gdFecSis)) - 1
End Sub
Private Function ValidaAnio() As Boolean
    ValidaAnio = True
    
    If Val(Me.txtAnio) > Year(gdFecSis) Or Val(Me.txtAnio) < 1990 Then
        ValidaAnio = False
        Exit Function
    End If
End Function
