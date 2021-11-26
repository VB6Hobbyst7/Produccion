VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmCredRelGarantias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relaciones de la Garantia con sus Creditos"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   Icon            =   "FrmCredRelGarantias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   0
      TabIndex        =   8
      Top             =   3630
      Width           =   7065
      Begin ComctlLib.ProgressBar PB 
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   -30
      TabIndex        =   7
      Top             =   1320
      Width           =   7095
      Begin MSComctlLib.ListView lstGarantias 
         Height          =   2175
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Num.Garantía"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estado Actual Cta."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estado en Garantía"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descripcion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Titular"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Disponible"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Titular de la Garantia"
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7065
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   5550
         TabIndex        =   6
         Top             =   630
         Width           =   1335
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   315
         Left            =   5550
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin SICMACT.TxtBuscar TxtBuscar1 
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   4320
         Picture         =   "FrmCredRelGarantias.frx":030A
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape Shape1 
         Height          =   1125
         Left            =   5460
         Top             =   90
         Width           =   1605
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   900
         TabIndex        =   4
         Top             =   600
         Width           =   4035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmCredRelGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub InicializarControles()
    TxtBuscar1.Text = ""
    lblNombre.Caption = ""
    lstgarantias.ListItems.Clear
    PB.value = 0
End Sub

Private Sub cmdCancelar_Click()
    Call Form_Load
End Sub

Private Sub cmdMostrar_Click()
    If TxtBuscar1.Text <> "" And lblNombre.Caption <> "" Then
        Call CargarDatos
     End If
End Sub

Private Sub Form_Load()
    InicializarControles
    
End Sub

Private Sub TxtBuscar1_EmiteDatos()
    Dim oMantGarantia As COMDCredito.DCOMGarantia 'DMantGarantia
    Dim sNombre As String
    
    On Error GoTo ErrHandler
    
    If TxtBuscar1.Text <> "" Then
        Set oMantGarantia = New COMDCredito.DCOMGarantia 'DMantGarantia
        sNombre = oMantGarantia.CargarNombrePersona(TxtBuscar1.Text)
        Set oMantGarantia = Nothing
    Else
        MsgBox "No existe  codigo de persona ", vbInformation, "AVISO"
        Exit Sub
    End If
    
    lblNombre.Caption = sNombre
    cmdMostrar.SetFocus
    Exit Sub
ErrHandler:
    If Not oMantGarantia Is Nothing Then Set oMantGarantia = Nothing
    MsgBox "Error al Cargar los datos de la persona " & vbCrLf & _
           "Consulte con el Area de TI", vbInformation, "AVISO"
End Sub


Sub CargarDatos()
    Dim cPersCod As String
    Dim oMantGarantia As COMDCredito.DCOMGarantia 'DMantGarantia
    Dim RS As ADODB.Recordset
    Dim iLisItem As ListItem
    On Error GoTo ErrHandler
    
    ' se cargar los datos
    cPersCod = TxtBuscar1.Text
    lstgarantias.ListItems.Clear
    Set oMantGarantia = New COMDCredito.DCOMGarantia 'DMantGarantia
    
    Set RS = oMantGarantia.ListaCreditosGarantias(cPersCod)
    Set oMantGarantia = Nothing
    
    Do Until RS.EOF
        Set iLisItem = lstgarantias.ListItems.Add(, , RS!cCtaCod)
         iLisItem.SubItems(1) = IIf(IsNull(RS!cNumGarant), "", RS!cNumGarant)
         iLisItem.SubItems(2) = IIf(IsNull(RS!Est_Act_Cred), "", RS!Est_Act_Cred) 'peac 20080111
         iLisItem.SubItems(3) = IIf(IsNull(RS!Est_En_Grntia), "", RS!Est_En_Grntia) 'peac 20080111
         iLisItem.SubItems(4) = IIf(IsNull(RS!cGarantias), "", RS!cGarantias)
         iLisItem.SubItems(5) = IIf(IsNull(RS!cPersNombre), "", RS!cPersNombre)
         iLisItem.SubItems(6) = Format(IIf(IsNull(RS!nMonto), 0, RS!nMonto), "#0.00")
         iLisItem.SubItems(7) = Format(IIf(IsNull(RS!nDisponible), "", RS!nDisponible), "#0.00")
         
        RS.MoveNext
    Loop

    Set RS = Nothing

    Exit Sub
ErrHandler:
    MsgBox "Error al cargar los datos", vbInformation, "AVISO"
    If Not oMantGarantia Is Nothing Then Set oMantGarantia = Nothing
    If Not RS Is Nothing Then Set RS = Nothing
End Sub
