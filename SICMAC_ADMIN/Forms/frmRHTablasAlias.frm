VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRHTablasAlias 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmRHTablasAlias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTotal 
      Height          =   5475
      Left            =   30
      TabIndex        =   1
      Top             =   -15
      Width           =   6555
      Begin VB.Frame fraTablas 
         Caption         =   "Tablas"
         Height          =   2610
         Left            =   105
         TabIndex        =   2
         Top             =   150
         Width           =   6345
         Begin VB.CommandButton cmdPasarDerTabBD 
            Caption         =   "Tablas BD >>"
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
            Left            =   2385
            TabIndex        =   6
            Top             =   1020
            Width           =   1575
         End
         Begin VB.CommandButton cmdPasarIzqTabBD 
            Caption         =   "<< Tablas BD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2385
            TabIndex        =   5
            Top             =   1410
            Width           =   1575
         End
         Begin VB.TextBox txtNomAliasTab 
            Height          =   285
            Left            =   105
            TabIndex        =   4
            Top             =   2250
            Width           =   2175
         End
         Begin VB.ListBox lstTabBD 
            Height          =   2010
            Left            =   120
            TabIndex        =   3
            Top             =   210
            Width           =   2175
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexAliasTab 
            Height          =   2085
            Left            =   4065
            TabIndex        =   7
            Top             =   450
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3678
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   1
         End
         Begin VB.Label lblAliasTab 
            Caption         =   "Alias Tablas de la BD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4065
            TabIndex        =   8
            Top             =   195
            Width           =   2175
         End
      End
      Begin VB.Frame fraCampos 
         Caption         =   "Campos"
         Height          =   2610
         Left            =   105
         TabIndex        =   9
         Top             =   2775
         Width           =   6345
         Begin VB.ListBox lstCamBD 
            Height          =   2010
            Left            =   120
            TabIndex        =   13
            Top             =   225
            Width           =   2175
         End
         Begin VB.CommandButton cmdPasarDerCamBD 
            Caption         =   "Campos BD >>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2400
            TabIndex        =   12
            Top             =   990
            Width           =   1575
         End
         Begin VB.CommandButton cmdPasarIzqCamBD 
            Caption         =   "<< Campos BD"
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
            Left            =   2400
            TabIndex        =   11
            Top             =   1395
            Width           =   1575
         End
         Begin VB.TextBox txtNomAliasCam 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   2250
            Width           =   2175
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexAliasCam 
            Height          =   2085
            Left            =   4020
            TabIndex        =   14
            Top             =   450
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   3678
            _Version        =   393216
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   1
         End
         Begin VB.Label lblAliasCam 
            Caption         =   "Alias Campo de la BD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   15
            Top             =   195
            Width           =   2175
         End
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      TabIndex        =   0
      Top             =   5520
      Width           =   1065
   End
End
Attribute VB_Name = "frmRHTablasAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String

Public Sub Ini(psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub

Private Sub cmdPasarDerCamBD_Click()
    Dim sqlPD As String
    Dim lsCod As String
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    Dim oConN As NRHConcepto
    Set oConN = New NRHConcepto
    
    If lstTabBD.ListIndex = -1 Then
        MsgBox "Debe Ingresar un Alias para la Tabla.", vbInformation, "Aviso"
        Me.txtNomAliasTab.SetFocus
    ElseIf Trim(txtNomAliasCam) = "" Then
        MsgBox "Debe Ingresar un Alias para el Campo.", vbInformation, "Aviso"
        Me.txtNomAliasCam.SetFocus
        Set oCon = Nothing
        Exit Sub
    End If
    
    If Not oCon.ExisteTablaAlias(lstTabBD.List(lstTabBD.ListIndex)) Then
       MsgBox "Para poder agregar alias a los campos de una tabla esta tiene que tener un alias.", vbInformation, "Aviso"
       Set oCon = Nothing
       Exit Sub
    End If
    
    If oCon.ExisteTablaAlias(lstCamBD.List(lstCamBD.ListIndex)) Then
       MsgBox "El Campo ya tiene un alias no puede agregarlo nuevamente.", vbInformation, "Aviso"
       Set oCon = Nothing
       Exit Sub
    End If
    
    lsCod = Me.FlexAliasTab.TextMatrix(FlexAliasTab.Row, 0)
    oConN.AgregaTablaAlias lsCod, Me.lstCamBD.List(lstCamBD.ListIndex), "C_" & Trim(Me.txtNomAliasCam.Text), GetMovNro(gsCodUser, gsCodAge)
    
    FlexAliasCam.Rows = FlexAliasCam.Rows + 1
    FlexAliasCam.TextMatrix(FlexAliasCam.Rows - 1, 0) = lsCod
    FlexAliasCam.TextMatrix(FlexAliasCam.Rows - 1, 1) = lstCamBD.List(lstCamBD.ListIndex)
    FlexAliasCam.TextMatrix(FlexAliasCam.Rows - 1, 2) = "C_" & Trim(txtNomAliasCam)
    
    txtNomAliasCam = ""
    lstCamBD.RemoveItem lstCamBD.ListIndex
    Set oCon = Nothing
End Sub

Private Sub cmdPasarDerTabBD_Click()
    
    Dim lsCod As String
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    Dim oConN As NRHConcepto
    Set oConN = New NRHConcepto
    
    If Trim(txtNomAliasTab) = "" Then
        MsgBox "Debe Ingresar un Alias para la Tabla.", vbInformation, "Aviso"
        Me.txtNomAliasTab.SetFocus
        Exit Sub
    ElseIf lstTabBD.ListIndex = -1 Then
        MsgBox "Debe una tabla a migrar.", vbInformation, "Aviso"
        Me.lstTabBD.SetFocus
        Exit Sub
    End If
    
    If oCon.ExisteTablaAlias(lstTabBD.List(lstTabBD.ListIndex)) Then
       MsgBox "La Tabla ya tiene un alias no puede agregarla", vbInformation, "Aviso"
       Exit Sub
    End If
    
    
    lsCod = ""
    oConN.AgregaTablaAlias lsCod, lstTabBD.List(lstTabBD.ListIndex), "T_" & Trim(txtNomAliasTab), GetMovNro(gsCodUser, gsCodAge)
    FlexAliasTab.Rows = FlexAliasTab.Rows + 1
    If FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 0) <> "" Then FlexAliasTab.Rows = FlexAliasTab.Rows + 1
    FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 0) = lsCod
    FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 1) = lstTabBD.List(lstTabBD.ListIndex)
    FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 2) = "T_" & Trim(txtNomAliasTab)
    
    txtNomAliasTab = ""
End Sub

Private Sub cmdPasarIzqCamBD_Click()
    Dim sqlDet As String
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    If MsgBox("Se eliminara el Alias Vinculados con el campo seleccionado." & FlexAliasTab.TextMatrix(FlexAliasTab.Row, 1), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    oCon.EliminaTablaAlias FlexAliasCam.TextMatrix(FlexAliasCam.Row, 0)
    
    If FlexAliasCam.Rows = 1 Then
        FlexAliasCam.TextMatrix(0, 0) = ""
        FlexAliasCam.TextMatrix(0, 1) = ""
        FlexAliasCam.TextMatrix(0, 2) = ""
    Else
        FlexAliasCam.RemoveItem FlexAliasCam.Row
    End If
    
    lstTabBD_Click
End Sub

Private Sub cmdPasarIzqTabBD_Click()
    Dim sqlDet As String
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    If MsgBox("Se eliminara todos Alias Vinculados con los campos y el nombre de la tabla " & FlexAliasTab.TextMatrix(FlexAliasTab.Row, 1), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    oCon.EliminaTablaAlias Mid(FlexAliasTab.TextMatrix(FlexAliasTab.Row, 0), 1, 3)
    
    If FlexAliasTab.Rows < 2 Then
        FlexAliasTab.TextMatrix(0, 0) = ""
        FlexAliasTab.TextMatrix(0, 1) = ""
        FlexAliasTab.TextMatrix(0, 2) = ""
    Else
        FlexAliasTab.RemoveItem FlexAliasTab.Row
    End If
    
    
    FlexAliasTab_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FlexAliasTab_Click()
    Dim sqlI As String
    Dim rsI As New ADODB.Recordset
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    If FlexAliasTab.TextMatrix(FlexAliasTab.Row, 0) = "" Then Exit Sub
    Set rsI = oCon.GetTablasAlias(Mid(FlexAliasTab.TextMatrix(FlexAliasTab.Row, 0), 1, 3))

    FlexAliasCam.Cols = 3
    FlexAliasCam.Rows = 0
    
    FlexAliasCam.ColWidth(0) = 1
    FlexAliasCam.ColWidth(1) = 1
    FlexAliasCam.ColWidth(2) = 2100
    
    If Not RSVacio(rsI) Then
        FlexAliasCam.Clear
        While Not rsI.EOF
            FlexAliasCam.Rows = FlexAliasCam.Rows + 1
            FlexAliasCam.TextMatrix(FlexAliasCam.Rows - 1, 0) = rsI!Codigo
            FlexAliasCam.TextMatrix(FlexAliasCam.Rows - 1, 1) = rsI!Nombre
            FlexAliasCam.TextMatrix(FlexAliasCam.Rows - 1, 2) = rsI!Alias
            rsI.MoveNext
        Wend
    End If
    
    rsI.Close
    Set rsI = Nothing
End Sub

Private Sub FlexAliasTab_EnterCell()
    FlexAliasTab_Click
End Sub

Private Sub Form_Load()
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    Dim rsT As ADODB.Recordset
    Set rsT = New ADODB.Recordset
    
    Caption = lsCaption
    
    IniFlex
    
    Set rsT = oCon.GetTablasBase

    lstTabBD.Clear
    If Not RSVacio(rsT) Then
        While Not rsT.EOF
            lstTabBD.AddItem rsT!Name
            rsT.MoveNext
        Wend
    End If
    rsT.Close
    Set rsT = Nothing
    Set oCon = Nothing
End Sub

Private Sub lstTabBD_Click()
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    If lstTabBD.ListIndex = -1 Then Exit Sub
    
    Set rsC = oCon.GetCamposTabla(lstTabBD.List(lstTabBD.ListIndex))
    lstCamBD.Clear
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            If Not oCon.ExisteTablaAlias(rsC!COLUMN_NAME) Then lstCamBD.AddItem rsC!COLUMN_NAME
            rsC.MoveNext
        Wend
    End If
    
    If FlexAliasTab.Rows <> 0 Then
        FlexAliasTab.Row = GetRowAlias(lstTabBD.List(lstTabBD.ListIndex))
        FlexAliasTab_Click
    End If
    rsC.Close
    Set rsC = Nothing
    Set oCon = Nothing
End Sub

Private Sub txtNomAliasCam_GotFocus()
    txtNomAliasCam.SelStart = 0
    txtNomAliasCam.SelLength = 50
End Sub

Private Sub txtNomAliasTab_GotFocus()
    txtNomAliasTab.SelStart = 0
    txtNomAliasTab.SelLength = 30
End Sub

Private Sub IniFlex()
    Dim rsI As ADODB.Recordset
    Set rsI = New ADODB.Recordset
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    Set rsI = oCon.GetTablasAlias("2")
    
    FlexAliasTab.Cols = 3
    FlexAliasTab.Rows = 0
    
    FlexAliasTab.ColWidth(0) = 1
    FlexAliasTab.ColWidth(1) = 1
    FlexAliasTab.ColWidth(2) = 2100
    
    FlexAliasCam.Cols = 3
    FlexAliasCam.Rows = 0
    
    FlexAliasCam.ColWidth(0) = 1
    FlexAliasCam.ColWidth(1) = 1
    FlexAliasCam.ColWidth(2) = 2100
    
    If Not RSVacio(rsI) Then
        While Not rsI.EOF
            FlexAliasTab.Rows = FlexAliasTab.Rows + 1
            FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 0) = rsI!Codigo
            FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 1) = rsI!Nombre
            FlexAliasTab.TextMatrix(FlexAliasTab.Rows - 1, 2) = rsI!Alias
            rsI.MoveNext
        Wend
    End If
    
    rsI.Close
    Set rsI = Nothing
    Set oCon = Nothing
End Sub

Private Function GetRowAlias(psNomTabla As String) As Integer
    Dim I As Integer
    Dim lnRes As Integer
    Dim lnPos As Integer
    
    lnPos = 0
    For I = 0 To FlexAliasTab.Rows - 1
        If Me.FlexAliasTab.TextMatrix(I, 1) = psNomTabla Then
            lnPos = I
            I = FlexAliasTab.Rows - 1
        End If
    Next I
        
    GetRowAlias = lnPos
End Function


