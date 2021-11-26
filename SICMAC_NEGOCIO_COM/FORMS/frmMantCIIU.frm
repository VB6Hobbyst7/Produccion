VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantCIIU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento CIIU"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTAB 
      Height          =   6975
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "CIIU"
      TabPicture(0)   =   "frmMantCIIU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraCIIU"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Giro"
      TabPicture(1)   =   "frmMantCIIU.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "FraDesc"
      Tab(1).Control(2)=   "fraGiro"
      Tab(1).Control(3)=   "LblSector"
      Tab(1).Control(4)=   "LblDescripcion"
      Tab(1).Control(5)=   "LblCIIU"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame5 
         Height          =   705
         Left            =   -74880
         TabIndex        =   22
         Top             =   6120
         Width           =   8025
         Begin VB.CommandButton CmdSalir2 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   5400
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdAceptar2 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   2520
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdNuevo2 
            Caption         =   "&Nuevo"
            Height          =   375
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdModificar2 
            Caption         =   "&Modificar"
            Height          =   375
            Left            =   3960
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FraDesc 
         Caption         =   "Descripcion"
         Height          =   615
         Left            =   -74880
         TabIndex        =   17
         Top             =   5520
         Visible         =   0   'False
         Width           =   7935
         Begin VB.TextBox TxtDescripcionGiro 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.Frame fraGiro 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   7935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FleGiro 
            Height          =   3735
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   6588
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
      End
      Begin VB.Frame Frame4 
         Height          =   705
         Left            =   120
         TabIndex        =   10
         Top             =   6120
         Width           =   8025
         Begin VB.CommandButton CmdModificar 
            Caption         =   "&Modificar"
            Height          =   375
            Left            =   3960
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   375
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   2520
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   5400
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   5160
         Width           =   8055
         Begin VB.ComboBox CboSector 
            Height          =   315
            Left            =   5880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox TxtCodCIIU 
            Height          =   285
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo CIIU"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   8
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Sector"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   7
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.Frame FraCIIU 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8055
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexCIIU 
            Height          =   4335
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   7646
            _Version        =   393216
            Cols            =   4
            TextStyleFixed  =   1
            GridLinesUnpopulated=   1
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
      End
      Begin VB.Label LblSector 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -68880
         TabIndex        =   21
         Top             =   795
         Width           =   720
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73560
         TabIndex        =   20
         Top             =   795
         Width           =   720
      End
      Begin VB.Label LblCIIU 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74520
         TabIndex        =   19
         Top             =   795
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMantCIIU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjDB As COMDPersona.DCOMPersGeneral
Dim opt As String
Dim ban As Integer
Dim nSector As Integer
Dim Opt2 As String

Private Sub CmdAceptar_Click()
Dim Op As Integer
Dim nCod As Integer

If opt = "" Then Exit Sub
If Me.TxtCodCIIU = "" Then
    MsgBox "Ingrese el Codigo CIIU", vbCritical, "AVISO"
    Exit Sub
End If

If Me.TxtDescripcion = "" Then
    MsgBox "Ingrese Descripcion del CIIU", vbCritical, "AVISO"
    Exit Sub
End If
If CboSector.ListIndex = -1 Then
    MsgBox "Escoja un Sector", vbCritical, "AVISO"
    Exit Sub
End If

Select Case opt
Case "M":
          Op = MsgBox("Esta seguro de Modificar?", vbInformation + vbYesNo, "AVISO")
          If Op = vbYes Then
            With Me.FlexCIIU
                Call ObjDB.dUpdateCIIU(.TextMatrix(.Row, 1), TxtDescripcion.Text, Trim(Right(CboSector.Text, 10)), gdFecSis, gsCodUser, gsCodAge)
            End With
          End If
          
Case "N":
          Op = MsgBox("Esta seguro de Grabar?", vbInformation + vbYesNo, "AVISO")
          If Op = vbYes Then
            If ObjDB.ExisteCIIU(UCase(Me.TxtCodCIIU)) Then
                MsgBox "Ya Existe Codigo CIIU. Ingrese uno Diferente", vbInformation, "AVISO"
                Exit Sub
            End If
            Call ObjDB.dInsertaCIIU(Me.TxtCodCIIU, Me.TxtDescripcion, Trim(Right(CboSector.Text, 10)), gdFecSis, gsCodUser, gsCodAge)
          End If

End Select
opt = ""
CmdNuevo.Enabled = True
cmdModificar.Enabled = True
cmdsalir.Caption = "&Salir"
FraCIIU.Height = 5535
FlexCIIU.Height = 5295
Carga
TxtCodCIIU.Enabled = True
End Sub

Private Sub CmdAceptar2_Click()
Dim Op As Integer
Select Case Opt2
Case "M":
          Op = MsgBox("Esta seguro de Modificar?", vbInformation + vbYesNo, "AVISO")
          If Op = vbYes Then
            If Trim(Me.TxtDescripcionGiro) = "" Then
                MsgBox "No se puede Actualizar este Campo", vbInformation, "AVISO"
            Else
                Call ObjDB.dUpdateCIIUGiro(LblCIIU, Trim(Me.FleGiro.TextMatrix(FleGiro.Row, 1)), Trim(Me.TxtDescripcionGiro))
            End If
          End If
          
Case "N":
          Op = MsgBox("Esta Seguro de Grabar?", vbInformation + vbYesNo, "AVISO")
          If Op = vbYes Then
            If ObjDB.ExisteCIIUGiro(Trim(Me.TxtDescripcionGiro), LblCIIU) Then
                 MsgBox "Ya Existe Descripcion Giro. Ingrese uno Diferente", vbInformation, "AVISO"
                 Exit Sub
            End If
            Call ObjDB.dInsertaCIIUGiro(LblCIIU, Trim(TxtDescripcionGiro))
          End If
End Select
Opt2 = ""
Me.CargaFlexGiro
Me.carga2
TxtDescripcionGiro = ""
CmdNuevo2.Enabled = True
CmdModificar2.Enabled = True
CmdSalir2.Caption = "&Salir"
Me.fraGiro.Height = 4815
Me.FleGiro.Height = 4500
FraDesc.Visible = False

End Sub

Private Sub CmdModificar_Click()
opt = "M"
TxtCodCIIU.Enabled = False
cmdsalir.Caption = "&Cancelar"
CmdNuevo.Enabled = False
cmdModificar.Enabled = False
FraCIIU.Height = 4695
FlexCIIU.Height = 4335
Frame2.Visible = True

'Limpia
'Carga los valores anteriores
TxtCodCIIU = FlexCIIU.TextMatrix(FlexCIIU.Row, 1)
TxtDescripcion = FlexCIIU.TextMatrix(FlexCIIU.Row, 2)
End Sub
Sub Limpia()
Me.TxtCodCIIU = ""
Me.TxtDescripcion = ""
End Sub

Private Sub CmdModificar2_Click()
CmdSalir2.Caption = "&Cancelar"
CmdNuevo2.Enabled = False
CmdModificar2.Enabled = False
Opt2 = "M"
Me.fraGiro.Height = 4215
Me.FleGiro.Height = 3735
FraDesc.Visible = True
End Sub

Private Sub cmdNuevo_Click()
cmdModificar.Enabled = False
cmdsalir.Caption = "&Cancelar"
CmdNuevo.Enabled = False
opt = "N"
FraCIIU.Height = 4695
FlexCIIU.Height = 4335
Frame2.Visible = True
Me.TxtCodCIIU.Enabled = True
Limpia

End Sub

Private Sub CmdNuevo2_Click()
CmdModificar2.Enabled = False
CmdSalir2.Caption = "&Cancelar"
CmdNuevo2.Enabled = False
TxtDescripcionGiro = ""
Opt2 = "N"
    
Me.fraGiro.Height = 4215
Me.FleGiro.Height = 3735
FraDesc.Visible = True
End Sub

Private Sub cmdsalir_Click()
If cmdsalir.Caption = "&Cancelar" Then
    cmdsalir.Caption = "&Salir"
    FraCIIU.Height = 5535
    FlexCIIU.Height = 5295
    Frame2.Visible = False
    CmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    opt = ""
    Limpia
Else
    Unload Me
End If

End Sub

Sub CargaFlexGiro()
Dim rs As New ADODB.Recordset
    Set rs = ObjDB.GetCIIUGiro(Trim(LblCIIU))
    Me.FleGiro.Clear
    Me.FleGiro.Rows = 2
    While Not rs.EOF
    If Me.FleGiro.Rows >= 2 And Me.FleGiro.TextMatrix(Me.FleGiro.Row, 0) = "" Then
       Me.FleGiro.TextMatrix(Me.FleGiro.Rows - 1, 1) = IIf(IsNull(rs(1)), "", rs(1))
       Me.FleGiro.TextMatrix(Me.FleGiro.Rows - 1, 2) = IIf(IsNull(rs(2)), "", rs(2))
       Me.FleGiro.Rows = Me.FleGiro.Rows + 1
    End If
        rs.MoveNext
    Wend
    If FleGiro.Rows > 2 Then FleGiro.Rows = FleGiro.Rows - 1
    carga2
End Sub

Private Sub CmdSalir2_Click()
If CmdSalir2.Caption = "&Cancelar" Then
    CmdSalir2.Caption = "&Salir"
    CmdNuevo2.Enabled = True
    CmdModificar2.Enabled = True
    opt = ""
    Me.fraGiro.Height = 4815
    Me.FleGiro.Height = 3735
    FraDesc.Visible = False
Else
    Unload Me
End If
End Sub

Private Sub FleGiro_Click()
If Opt2 = "M" Then
    TxtDescripcionGiro = Me.FleGiro.TextMatrix(FleGiro.Row, 1)
End If
End Sub

Private Sub FlexCIIU_Click()
Dim rs As New ADODB.Recordset
Dim id As String
    

LblCIIU = FlexCIIU.TextMatrix(FlexCIIU.Row, 1)
lblDescripcion = FlexCIIU.TextMatrix(FlexCIIU.Row, 2)
LblSector = FlexCIIU.TextMatrix(FlexCIIU.Row, 3)

ban = 0
If FlexCIIU.Row > 0 Then
    CargaFlexGiro
    Me.fraGiro.Height = 4815
    Me.FleGiro.Height = 4500
    ban = 1
End If

If opt = "M" Then
    TxtCodCIIU = FlexCIIU.TextMatrix(FlexCIIU.Row, 1)
    TxtDescripcion = FlexCIIU.TextMatrix(FlexCIIU.Row, 2)
    id = Trim(Right(FlexCIIU.TextMatrix(FlexCIIU.Row, 3), 4))
   
    If id = "" Then id = -1
    CboSector.ListIndex = Val(id)
End If

Set rs = Nothing
End Sub


Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim oCons As COMDConstantes.DCOMConstantes

    Set ObjDB = New COMDPersona.DCOMPersGeneral
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    ban = 0
        
    Set oCons = New COMDConstantes.DCOMConstantes
    Set rs = oCons.RecuperaConstantes(1013)
    Set oCons = Nothing
    Call Llenar_Combo_con_Recordset(rs, CboSector)
    'Call CargaComboConstante(1013, CboSector)
    Carga
 
    Opt2 = ""
    opt = ""
End Sub
Sub carga2()
With Me.FleGiro
 .TextMatrix(0, 1) = "Descripcion"
 .TextMatrix(0, 2) = "CIIU"
 .ColWidth(0) = 100
 .ColWidth(1) = 4500
 .ColWidth(2) = 1500
 .ColAlignment(1) = flexAlignLeftTop
End With
End Sub

Sub Carga()
Dim rs As New ADODB.Recordset
Set rs = ObjDB.GetCIIU
FlexCIIU.Clear
FlexCIIU.Rows = 2
With FlexCIIU
 .TextMatrix(0, 1) = "Cod CIIU"
 .TextMatrix(0, 2) = "Descripcion"
 .TextMatrix(0, 3) = "Sector"
 .ColWidth(0) = 100
 .ColWidth(1) = 1300
 .ColWidth(2) = 4500
 .ColWidth(3) = 1500
 .ColAlignment(1) = flexAlignLeftTop
 While Not rs.EOF
    If .Rows >= 2 And .TextMatrix(.Row, 0) = "" Then
       .TextMatrix(.Rows - 1, 1) = rs!cCIIUcod
       .TextMatrix(.Rows - 1, 2) = rs!cCIIUdescripcion
       .TextMatrix(.Rows - 1, 3) = IIf(IsNull(rs!Sector), "", rs!Sector)
       .Rows = .Rows + 1
    End If
    rs.MoveNext
 Wend
End With
If FlexCIIU.Rows > 2 Then FlexCIIU.Rows = FlexCIIU.Rows - 1
FraCIIU.Height = 5535
FlexCIIU.Height = 5295
Frame2.Visible = False
Limpia
End Sub

Private Sub Form_Terminate()
Set ObjDB = Nothing
End Sub

Private Sub SSTAB_Click(PreviousTab As Integer)
If ban = 0 Then
    Me.SSTAB.Tab = 0
    MsgBox "Escoja una un CIIU", vbInformation, "AVISO"
End If
If Trim(LblSector) = "" Then
    Me.SSTAB.Tab = 0
    MsgBox "Escoja un sector para el CIIU", vbInformation, "AVISO"
End If
End Sub

Private Sub TxtDescripcionGiro_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
End Sub
