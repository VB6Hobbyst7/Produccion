VERSION 5.00
Begin VB.Form frmCapMantenimientoCliente 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmCapMantenimientoCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3150
      TabIndex        =   5
      Top             =   3465
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1785
      TabIndex        =   4
      Top             =   3465
      Width           =   1170
   End
   Begin VB.Frame fraPersona 
      Height          =   3375
      Left            =   105
      TabIndex        =   6
      Top             =   50
      Width           =   5895
      Begin VB.Frame fraRelacion 
         Caption         =   "Relación"
         Height          =   750
         Left            =   105
         TabIndex        =   13
         Top             =   2520
         Width           =   5685
         Begin VB.ComboBox cboRelacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   300
            Width           =   3480
         End
      End
      Begin VB.Frame fraDatosPers 
         Height          =   1905
         Left            =   105
         TabIndex        =   8
         Top             =   580
         Width           =   5685
         Begin VB.ListBox lstID 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   3255
            TabIndex        =   2
            Top             =   945
            Width           =   2220
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Doc ID:"
            Height          =   195
            Left            =   2625
            TabIndex        =   16
            Top             =   945
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Diercción :"
            Height          =   195
            Left            =   105
            TabIndex        =   15
            Top             =   525
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nombre :"
            Height          =   195
            Left            =   105
            TabIndex        =   14
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Personería :"
            Height          =   195
            Left            =   105
            TabIndex        =   12
            Top             =   945
            Width           =   870
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1050
            TabIndex        =   11
            Top             =   525
            Width           =   4425
         End
         Begin VB.Label lblPersoneria 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   750
            Left            =   1050
            TabIndex        =   10
            Top             =   945
            Width           =   1485
         End
         Begin VB.Label lblNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1050
            TabIndex        =   9
            Top             =   210
            Width           =   4425
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   330
         Left            =   2520
         TabIndex        =   1
         Top             =   210
         Width           =   435
      End
      Begin VB.TextBox txtPersCod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   0
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   210
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmCapMantenimientoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bAgregar As Boolean
Dim sPersoneria As String
Dim sTelefono As String
Dim sZona As String
Dim rsID As Recordset

Private Function ExistePersona(ByVal sPers As String) As Boolean
Dim i As Integer
Dim Existe As Boolean
Existe = False
For i = 1 To frmCapMantenimiento.grdCliente.Rows - 1
    If frmCapMantenimiento.grdCliente.TextMatrix(i, 9) = sPers Then
        Existe = True
        Exit For
    End If
Next i
ExistePersona = Existe
End Function

Private Sub cboRelacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdAceptar.SetFocus
End If
End Sub

Private Sub CmdAceptar_Click()
Dim nItem As Integer
Dim i As Integer
If bAgregar Then
    rsID.MoveFirst
    Do While Not rsID.EOF
        With frmCapMantenimiento.grdCliente
            .AdicionaFila ""
            nItem = .Rows - 1
            .TextMatrix(nItem, 1) = lblNombre
            .TextMatrix(nItem, 2) = Left(cboRelacion.Text, 2)
            .TextMatrix(nItem, 3) = LblDireccion
            .TextMatrix(nItem, 4) = sZona
            .TextMatrix(nItem, 5) = sTelefono
            .TextMatrix(nItem, 6) = rsID("Tipo")
            .TextMatrix(nItem, 7) = rsID("cPersIDNro")
            .TextMatrix(nItem, 8) = Trim(Right(cboRelacion, 2))
            .TextMatrix(nItem, 9) = txtPersCod
            .TextMatrix(nItem, 10) = sPersoneria
            .TextMatrix(nItem, 11) = rsID("cPersIDTpo")
        End With
        rsID.MoveNext
    Loop
    'frmCapMantenimiento.AgregaClienteMatriz txtPersCod, Trim(Right(cboRelacion, 2))
Else
    With frmCapMantenimiento.grdCliente
        For i = 1 To .Rows - 1
            If txtPersCod = .TextMatrix(i, 7) Then
                .TextMatrix(i, 2) = Left(cboRelacion.Text, 2)
                .TextMatrix(i, 6) = Trim(Right(cboRelacion, 2))
            End If
        Next i
    End With
    'frmCapMantenimiento.ModificaClienteMatriz txtPersCod, Trim(Right(cboRelacion, 2))
End If
Unload Me
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As UPersona
Set clsPers = New UPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    txtPersCod = clsPers.sPersCod
    If Not ExistePersona(txtPersCod) Then
        lblPersoneria = clsPers.sPersPersoneriaDesc
        lblNombre = clsPers.sPersNombre
        LblDireccion = clsPers.sPersDireccDomicilio
        sPersoneria = clsPers.sPersPersoneria
        sTelefono = clsPers.sPersTelefono
        'sZona = clsPers.sPersZona
        Set rsID = clsPers.DocsPers
        lstID.Clear
        Do While Not clsPers.DocsPers.EOF
            lstID.AddItem clsPers.DocsPers("Tipo") & Space(2) & clsPers.DocsPers("cPersIDNro") & Space(50) & clsPers.DocsPers("cPersIDTpo")
            clsPers.DocsPers.MoveNext
        Loop
        Set clsPers = Nothing
        cboRelacion.SetFocus
    Else
        MsgBox "Persona ya existe en relación con cuenta.", vbInformation, "Aviso"
        txtPersCod = ""
        cmdBuscar.SetFocus
    End If
Else
    cmdBuscar.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = ""
Dim clsRel As NCapMantenimiento
Dim rsRel As Recordset

Set clsRel = New NCapMantenimiento
Set rsRel = New Recordset
rsRel.CursorLocation = adUseClient
'Set rsRel = clsRel.GetRelProdPersona()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Do While Not rsRel.EOF
    cboRelacion.AddItem Trim(UCase(rsRel("cConsDescripcion"))) & Space(100) & rsRel("cConsValor")
    rsRel.MoveNext
Loop
Set clsRel = Nothing
Set rsRel = Nothing
cboRelacion.ListIndex = 0
txtPersCod.Enabled = False
Set rsID = New Recordset
rsID.CursorLocation = adUseClient
sPersoneria = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsID = Nothing
End Sub

Private Sub txtPersCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtPersCod = "" Then
        cmdBuscar.SetFocus
        
    End If
End If
End Sub
