VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContabManVar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Contables: Mantenimiento (Basado en Variables)"
   ClientHeight    =   6030
   ClientLeft      =   2340
   ClientTop       =   1905
   ClientWidth     =   7155
   Icon            =   "frmContabManVar.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7155
   Begin MSComctlLib.StatusBar SBEstado 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   30
      Top             =   5730
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Cuentas Contables"
      TabPicture(0)   =   "frmContabManVar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Valores de Var."
      TabPicture(1)   =   "frmContabManVar.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Variables"
      TabPicture(2)   =   "frmContabManVar.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Restricciones"
      TabPicture(3)   =   "frmContabManVar.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2"
      Tab(3).Control(1)=   "Label3"
      Tab(3).Control(2)=   "Label4"
      Tab(3).Control(3)=   "LstVar"
      Tab(3).Control(4)=   "LstValVar"
      Tab(3).Control(5)=   "LstRest"
      Tab(3).Control(6)=   "CmdAddRest"
      Tab(3).Control(7)=   "CmdRemRest"
      Tab(3).ControlCount=   8
      Begin VB.CommandButton CmdRemRest 
         Caption         =   "&Remover"
         Height          =   345
         Left            =   -69705
         TabIndex        =   35
         Top             =   2475
         Width           =   1095
      End
      Begin VB.CommandButton CmdAddRest 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   -71040
         TabIndex        =   34
         Top             =   2475
         Width           =   1065
      End
      Begin MSComctlLib.ListView LstRest 
         Height          =   1650
         Left            =   -71640
         TabIndex        =   33
         Top             =   3120
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   2910
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cod. Valor"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Descrip. Rest."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Valor"
            Object.Width           =   794
         EndProperty
      End
      Begin MSComctlLib.ListView LstValVar 
         Height          =   1665
         Left            =   -71670
         TabIndex        =   32
         Top             =   750
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2937
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cod. Valor"
            Object.Width           =   1323
         EndProperty
      End
      Begin MSComctlLib.ListView LstVar 
         Height          =   4050
         Left            =   -74865
         TabIndex        =   31
         Top             =   735
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cod. Valor"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   3975
         Left            =   -74895
         TabIndex        =   27
         Top             =   795
         Width           =   6645
         Begin VB.Frame Frame5 
            Height          =   960
            Left            =   135
            TabIndex        =   29
            Top             =   2895
            Width           =   6405
            Begin VB.CommandButton CmdAceptar 
               Caption         =   "&Aceptar"
               Enabled         =   0   'False
               Height          =   345
               Left            =   1545
               TabIndex        =   21
               Top             =   570
               Width           =   1425
            End
            Begin VB.CommandButton CmdCancelar 
               Caption         =   "&Cancelar"
               Enabled         =   0   'False
               Height          =   345
               Left            =   3405
               TabIndex        =   22
               Top             =   570
               Width           =   1485
            End
            Begin VB.TextBox TxtValor 
               Enabled         =   0   'False
               Height          =   300
               Left            =   5070
               TabIndex        =   20
               Top             =   195
               Width           =   840
            End
            Begin VB.TextBox txtDescrip 
               Enabled         =   0   'False
               Height          =   300
               Left            =   435
               TabIndex        =   19
               Top             =   195
               Width           =   4635
            End
         End
         Begin VB.CommandButton CmdElimVar 
            Caption         =   "&Eliminar"
            Height          =   360
            Left            =   5340
            TabIndex        =   18
            Top             =   1530
            Width           =   1230
         End
         Begin VB.CommandButton CmdModifCta 
            Caption         =   "&Modificar"
            Height          =   360
            Left            =   5340
            TabIndex        =   17
            Top             =   1080
            Width           =   1230
         End
         Begin VB.CommandButton CmdNuevaVar 
            Caption         =   "&Nuevo"
            Height          =   360
            Left            =   5325
            TabIndex        =   16
            Top             =   630
            Width           =   1230
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFVar 
            Height          =   2385
            Left            =   180
            TabIndex        =   15
            Top             =   480
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin VB.Label Label1 
            Caption         =   "Variables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   165
            TabIndex        =   28
            Top             =   210
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Variables"
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
         Height          =   4035
         Left            =   -74790
         TabIndex        =   24
         Top             =   750
         Width           =   6435
         Begin VB.CommandButton CmdAceptarVal 
            Caption         =   "&Aceptar"
            Height          =   345
            Left            =   1815
            TabIndex        =   13
            Top             =   3645
            Width           =   1215
         End
         Begin VB.CommandButton CmdCancelarVal 
            Caption         =   "&Cancelar"
            Height          =   345
            Left            =   3330
            TabIndex        =   14
            Top             =   3645
            Width           =   1215
         End
         Begin VB.CommandButton CmdNewVal 
            Caption         =   "&Nuevo"
            Height          =   345
            Left            =   5010
            TabIndex        =   6
            Top             =   345
            Width           =   1215
         End
         Begin VB.CommandButton CmdGenVal 
            Caption         =   "&Generar"
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
            Left            =   5040
            TabIndex        =   9
            Top             =   2535
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Height          =   615
            Left            =   30
            TabIndex        =   26
            Top             =   2955
            Width           =   6375
            Begin VB.TextBox txtVarValor 
               Enabled         =   0   'False
               Height          =   315
               Left            =   5490
               TabIndex        =   12
               Top             =   210
               Width           =   765
            End
            Begin VB.TextBox txtVarDesc 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1020
               TabIndex        =   11
               Top             =   210
               Width           =   4455
            End
            Begin VB.TextBox txtVarCod 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   10
               Top             =   210
               Width           =   885
            End
         End
         Begin VB.ComboBox CboVar 
            Height          =   315
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   4665
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFValVar 
            Height          =   2385
            Left            =   150
            TabIndex        =   5
            Top             =   570
            Width           =   4665
            _ExtentX        =   8229
            _ExtentY        =   4207
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   4
         End
         Begin VB.CommandButton CmdEdiVal 
            Caption         =   "&Editar"
            Height          =   345
            Left            =   5010
            TabIndex        =   7
            Top             =   705
            Width           =   1215
         End
         Begin VB.CommandButton CmdEliVal 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   5010
            TabIndex        =   8
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estructura de Cuentas"
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
         Height          =   3975
         Left            =   210
         TabIndex        =   25
         Top             =   750
         Width           =   6435
         Begin VB.ListBox lstCtas 
            Height          =   3570
            ItemData        =   "frmContabManVar.frx":037A
            Left            =   180
            List            =   "frmContabManVar.frx":037C
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   270
            Width           =   4665
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Variables"
            Height          =   375
            Left            =   4995
            TabIndex        =   2
            Top             =   270
            Width           =   1215
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   4995
            TabIndex        =   3
            Top             =   705
            Width           =   1215
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Restricciones:"
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
         Left            =   -71610
         TabIndex        =   38
         Top             =   2880
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "Variables A Restringir :"
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
         Left            =   -71640
         TabIndex        =   37
         Top             =   495
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "Variables :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74790
         TabIndex        =   36
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5580
      TabIndex        =   23
      Top             =   5250
      Width           =   1215
   End
End
Attribute VB_Name = "frmContabManVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nAccion As Integer
Dim nCod As Integer
Dim nPosVal As Integer
Dim oCon As DConecta

Private Sub CargaDatosLstVar(lstLista As Control)
Dim sSql As String
Dim R As New ADODB.Recordset
Dim L As ListSubItems
    lstLista.ListItems.Clear
    sSql = "Select VV.*,VC.cAbrev From ValVarCtas VV inner join VarCtasCont VC  On VV.nCodigo = VC.nCodigo " _
           & " Order By VV.nCodigo"
    Set R = oCon.CargaRecordSet(sSql)
        Do While Not R.EOF
            Set L = lstLista.ListItems.Add(, , Left(R!cAbrev & Space(3), 4) & Trim(R!cDescrip))
            Call L.Add(, , Trim(R!cValor))
            Call L.Add(, , R!nCodigo)
            Call L.Add(, , Trim(R!cCodValor))
            R.MoveNext
        Loop
    R.Close
End Sub
Private Sub DescripRest(ByVal nCodRes As Integer, ByVal cCodValRes As String, DRest As String, VRest As String)
Dim sSql As String
Dim R As New ADODB.Recordset
    sSql = "Select cValor, cDescrip From ValVarCtas Where nCodigo = " & Trim(Str(nCodRes)) & " And cCodValor = '" & Trim(cCodValRes) & "'"
    Set R = oCon.CargaRecordSet(sSql)
        If Not R.BOF And Not R.EOF Then
            DRest = Trim(R!cDescrip)
            VRest = Trim(R!cValor)
        End If
    R.Close
    Set R = Nothing
End Sub

Private Sub CargaDatosRestricciones(ByVal nCodigo As Integer, ByVal cCodValor As String)
Dim sSql As String
Dim R As New ADODB.Recordset
Dim L As ListSubItems
Dim DRest As String
Dim VRest As String

    LstRest.ListItems.Clear
    sSql = "Select R.nCodRes, R.cCodValres, VV.cDescrip, VV.cValor from RestVarCta R inner join ValVarCtas VV On R.nCodigo = VV.nCodigo and R.cCodValor = VV.cCodValor " _
          & " Where R.nCodigo = " & Trim(Str(nCodigo)) & " And R.cCodValor = '" & Trim(cCodValor) & "'"
    Set R = oCon.CargaRecordSet(sSql)
        Do While Not R.EOF
            Set L = LstRest.ListItems.Add(, , Trim(R!cDescrip))
            Call L.Add(, , Trim(R!cValor))
            Call L.Add(, , Trim(R!nCodRes))
            Call L.Add(, , Trim(R!cCodValRes))
            Call DescripRest(R!nCodRes, R!cCodValRes, DRest, VRest)
            Call L.Add(, , Trim(DRest))
            Call L.Add(, , Trim(VRest))
            R.MoveNext
        Loop
    R.Close
End Sub
Private Sub SelecFila(MSFlex As Control, ByVal nFila As Integer, ByVal ColIni As Integer)
    MSFlex.Row = nFila
    MSFlex.Col = ColIni
    'MSFlex.CellBackColor = &H8000000D
    MSFlex.RowSel = nFila
    MSFlex.ColSel = MSFlex.Cols - 1
End Sub
Private Function EliminaVarCuentas(ByVal nCodigo As Long, ByVal sValor As String) As Boolean
Dim I, j As Integer
Dim sSql As String
Dim R As New ADODB.Recordset
Dim R2 As New ADODB.Recordset
Dim nPos As Integer
Dim nLen As Integer
Dim CadMensage As String
Dim CadSql() As String
Dim ContCad As Integer
    ContCad = 0
    ReDim CadSql(ContCad)
    
    SBEstado.Panels(1).Text = "Generando Cuentas.."
    On Error GoTo ErrorCta
    
    nPos = 0
    sSql = "Select CG.cClase From CtasGen CG  Where CG.nCodigo = " & Trim(Str(nCodigo))
    Set R = oCon.CargaRecordSet(sSql)
    Do While Not R.EOF
        nPos = Len(R!cClase) + 1
        sSql = "Select CG.nCodigo,VC.cAbrev from CtasGen CG inner join VarCtasCont VC on CG.nCodigo = VC.nCodigo Where cClase = '" & R!cClase & "' Order By nOrden"
        Set R2 = oCon.CargaRecordSet(sSql)
            Do While Trim(R2!nCodigo) <> Trim(Str(nCodigo))
                nPos = nPos + Len(RTrim(R2!cAbrev))
                R2.MoveNext
            Loop
            nLen = Len(R2!cAbrev)
        R2.Close
        sSql = "DELETE CtaCont Where SubString(cCtaContCod,1,2) = '" & R!cClase & "' And SubString(cCtaContCod," & Trim(Str(nPos)) & "," & Trim(Str(nLen)) & ") = '" & Trim(sValor) & "'"
        ContCad = ContCad + 1
        ReDim Preserve CadSql(ContCad)
        CadSql(ContCad - 1) = sSql
        'ocon.Ejecutar sSql
        
        R.MoveNext
    Loop
    R.Close
    oCon.BeginTrans
        For I = 0 To ContCad - 1
            oCon.Ejecutar CadSql(I)
        Next I
    oCon.CommitTrans
    EliminaVarCuentas = True
    SBEstado.Panels(1).Text = "Proceso ha Finalizado.."
    Exit Function
ErrorCta:
    oCon.RollbackTrans
    CadMensage = "Error Nº [" & Err.Number & "]" & Err.Description & Chr(13) & "Consulte al Area de Sistemas"
    
    MsgBox CadMensage, vbInformation, "Aviso"
    EliminaVarCuentas = False
    SBEstado.Panels(1).Text = "Proceso ha Finalizado.."
    On Error GoTo 0
End Function


Private Sub GeneraVarCuentasModificadas(ByVal nCodigo As Long, ByVal sValor As String)
Dim I, j As Integer
Dim sCadSel As String
Dim sCadFrom As String
Dim sCadWhere As String
Dim sCadRest As String
Dim sCadTotal As String
Dim ContVar As Integer
Dim sSql As String
Dim R As New ADODB.Recordset
Dim R2 As New ADODB.Recordset
Dim R3 As New ADODB.Recordset
Dim R4 As New ADODB.Recordset
Dim sCtaArmada As String

    SBEstado.Panels(1).Text = "Generando Cuentas.."
    oCon.Ejecutar "Delete TempCtasGen"
    
    sSql = "Select CG.cClase From CtasGen CG Where CG.nCodigo = " & Trim(Str(nCodigo))
    Set R = oCon.CargaRecordSet(sSql)
    Do While Not R.EOF
        sSql = "Select * from CtasGen Where cClase = '" & R!cClase & "' Order By nOrden"
        Set R2 = oCon.CargaRecordSet(sSql)
            R2.Find "nCodigo = " & Trim(Str(nCodigo)), , adSearchForward, 1
            ContVar = R2.Bookmark
            Do While Not R2.EOF
                sCadSel = "INSERT INTO TempCtasGen Select Cta.cCtaContCod, Cta.cDescrip, newid() FROM (Select '" & R!cClase & "'"
                sCadFrom = " FROM "
                sCadWhere = " WHERE "
                sCadRest = "And ("
                sCtaArmada = "'" & Trim(R!cClase) & "'"
                If ContVar <= R2.RecordCount Then
                    sSql = "Select * from CtasGen Where cClase = '" & R!cClase & "' Order By nOrden"
                    Set R3 = oCon.CargaRecordSet(sSql)
                        For I = 0 To ContVar - 1
                                sCadSel = sCadSel & " + RTRIM(VV" & Trim(Str(R3!nCodigo)) & ".cValor) as cCtaContCod "
                                sCtaArmada = sCtaArmada & " + RTRIM(VV" & Trim(Str(R3!nCodigo)) & ".cValor)"
                                sCadFrom = sCadFrom & "ValvarCtas " & "VV" & Trim(Str(R3!nCodigo)) & ", "
                                sCadWhere = sCadWhere & "VV" & Trim(Str(R3!nCodigo)) & ".nCodigo = " & Trim(Str(R3!nCodigo)) & " and "
                                
                                sSql = "Select * from RestVarCta Where nCodigo = " & Trim(Str(R3!nCodigo)) & " And cCodValor = '" & R3!cClase & "'"
                                Set R4 = oCon.CargaRecordSet(sSql)
                                    Do While Not R4.EOF
                                        sCadRest = sCadRest & " (VV" & Trim(Str(R4!nCodigo)) & ".cCodValor = '" & Trim(R4!cCodValor) & "' And VV" & Trim(Str(R4!nCodRes)) & ".cCodValor = '" & Trim(R4!cCodValRes) & "') OR "
                                        R4.MoveNext
                                    Loop
                                R4.Close
                                
                                R3.MoveNext
                        Next I
                    R3.MovePrevious
                    sCadSel = sCadSel & ",VV" & Trim(Str(R3!nCodigo)) & ".cDescrip "
                    R3.Close
                    
                    sCadFrom = Mid(sCadFrom, 1, Len(sCadFrom) - 2)
                    sCadWhere = Mid(sCadWhere, 1, Len(sCadWhere) - 4)
                    sCadTotal = sCadSel + sCadFrom + sCadWhere
                    'Añade las Restricciones
                    If Len(sCadRest) > 5 Then
                        sCadRest = "Select " & sCtaArmada & Space(1) & sCadFrom & sCadWhere & Mid(sCadRest, 1, Len(sCadRest) - 3) + ")"
                        sCadTotal = sCadTotal & " ) Cta And LEFT JOIN (" & sCadRest & ")"
                    End If
                    
                    SBEstado.Panels(1).Text = "Generando Cuentas.."
                    oCon.Ejecutar sCadTotal
                    ContVar = ContVar + 1
                End If
                R2.MoveNext
            Loop
        R2.Close
        
        R.MoveNext
    Loop
    R.Close
    
    sSql = "INSERT INTO CtaCont  Select * from TempCtasGen Where cCta not in (Select cCtaContCod from CtaCont)"
    oCon.Ejecutar sSql
    
    SBEstado.Panels(1).Text = "Proceso ha Finalizado.."
End Sub


Private Sub CargaDatosCtasVar()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim R2 As New ADODB.Recordset
Dim sCta() As String
Dim NumCtas As Integer
Dim I As Integer
Dim Cad As String
    lstCtas.Clear
    NumCtas = 0
    ReDim sCta(NumCtas)
    
    sSql = "Select cClase from CtasGen Group By cClase"
    Set R = oCon.CargaRecordSet(sSql)
        Do While Not R.EOF
            NumCtas = NumCtas + 1
            ReDim Preserve sCta(NumCtas)
            sCta(NumCtas - 1) = R!cClase
            R.MoveNext
        Loop
    R.Close
    For I = 0 To NumCtas - 1
        sSql = "Select RTrim(V.cAbrev) as Abrev from CtasGen CG inner join VarCtasCont V " _
              & " On V.ncodigo = CG.nCodigo Where Rtrim(CG.cClase)='" & sCta(I) & "' ORDER BY nOrden"
        Set R = oCon.CargaRecordSet(sSql)
            Cad = sCta(I)
            Do While Not R.EOF
                Cad = Cad + R!Abrev
                R.MoveNext
            Loop
        R.Close
        Call lstCtas.AddItem(Cad)
    Next I
End Sub
Private Sub HabDesBotonVarVal(ByVal bVarCod As Boolean, ByVal bVarDesc As Boolean, ByVal bVarValor As Boolean, ByVal bAcepval As Boolean, ByVal bCanVal As Boolean)
    txtVarCod.Enabled = bVarCod
    txtVarDesc.Enabled = bVarDesc
    txtVarValor.Enabled = bVarValor
    CmdAceptarVal.Enabled = bAcepval
    CmdCancelarVal.Enabled = bCanVal
End Sub
Private Sub MuestraFilaVal()
Dim R As New ADODB.Recordset
Dim sSql As String
Dim nTotal As Integer
    sSql = "Select count(*) as Total from ValVarCtas Where nCodigo = " & Right(CboVar.Text, 2)
    Set R = oCon.CargaRecordSet(sSql)
        nTotal = IIf(IsNull(R!Total), 0, R!Total)
    R.Close
    If nTotal > 0 Then
        txtVarCod.Text = MSFValVar.TextMatrix(MSFValVar.Row, 1)
        txtVarDesc.Text = MSFValVar.TextMatrix(MSFValVar.Row, 2)
        txtVarValor.Text = MSFValVar.TextMatrix(MSFValVar.Row, 3)
    End If
End Sub
Private Sub CargaDatosVal(ByVal pnCodigo As Integer)
Dim sSql As String
Dim R As New ADODB.Recordset
Dim nPos As Integer
    sSql = "Select C.nCodigo, VC.cCodValor, VC.cDescrip, VC.cValor from ValvarCtas VC inner join  VarCtasCont C " _
           & " ON VC.nCodigo = C.nCodigo  Where C.nCodigo = " & Trim(Str(pnCodigo)) & " Order By VC.cCodValor "
    Set R = oCon.CargaRecordSet(sSql)
        MSFValVar.Rows = 0
        MSFValVar.Rows = IIf(R.RecordCount = 0, 1, R.RecordCount) + 1
        Call CargaCabeceraVar
        nPos = 1
        MSFValVar.FixedRows = 1
        Do While Not R.EOF
            MSFValVar.TextMatrix(nPos, 0) = Trim(Str(R!nCodigo))
            MSFValVar.TextMatrix(nPos, 1) = Trim(R!cCodValor)
            MSFValVar.TextMatrix(nPos, 2) = Trim(IIf(IsNull(R!cDescrip), "", R!cDescrip))
            MSFValVar.TextMatrix(nPos, 3) = Trim(R!cValor)
            nPos = nPos + 1
            R.MoveNext
        Loop
    R.Close
End Sub

Private Sub CargaComboVal()
Dim sSql As String
Dim R As New ADODB.Recordset
    CboVar.Clear
    sSql = "Select * from VarCtasCont"
    Set R = oCon.CargaRecordSet(sSql)
    Do While Not R.EOF
        Call CboVar.AddItem(R!cDescrip + Space(150) + Trim(Str(R!nCodigo)))
        R.MoveNext
    Loop
    R.Close
    If CboVar.ListCount > 0 Then
        CboVar.ListIndex = 0
    End If
End Sub

Private Sub MuestraFilaSelIns(ByVal nPosFil As Integer)
    txtDescrip.Text = MSFVar.TextMatrix(nPosFil, 1)
    TxtValor.Text = MSFVar.TextMatrix(nPosFil, 2)
End Sub

Private Sub CargaCabeceraIns()
Dim I As Integer
Dim sColumnas(3) As String
Dim sTamCol(3) As Integer
    
    sColumnas(0) = "Codigo"
    sColumnas(1) = "Descripcion"
    sColumnas(2) = "Valor"
    sTamCol(0) = 750
    sTamCol(1) = 3450
    sTamCol(2) = 750
        
    MSFVar.RowSel = 0
    For I = 0 To MSFVar.Cols - 1
        MSFVar.TextMatrix(0, I) = sColumnas(I)
        MSFVar.Col = I
        MSFVar.ColWidth(I) = sTamCol(I)
    Next I
End Sub
Private Sub CargaCabeceraVar()
Dim I As Integer
Dim sColumnas(4) As String
Dim sTamCol(4) As Integer
    
    sColumnas(0) = "Cod 01"
    sColumnas(1) = "Cod 02 "
    sColumnas(2) = "Descripcion"
    sColumnas(3) = "Valor"
    sTamCol(0) = 400
    sTamCol(1) = 400
    sTamCol(2) = 3000
    sTamCol(3) = 750
    MSFVar.Row = 0
    For I = 0 To MSFValVar.Cols - 1
        MSFValVar.TextMatrix(0, I) = sColumnas(I)
        MSFValVar.ColSel = I
        MSFValVar.ColWidth(I) = sTamCol(I)
    Next I
    MSFValVar.FixedRows = 1
End Sub
Private Sub CargaDatosIns()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim nPos As Integer
        
    sSql = "Select * from VarCtasCont"
    Set R = oCon.CargaRecordSet(sSql)
        MSFVar.Rows = R.RecordCount + 1
        Call CargaCabeceraIns
        nPos = 1
        Do While Not R.EOF
            MSFVar.TextMatrix(nPos, 0) = Trim(Str(R!nCodigo))
            MSFVar.TextMatrix(nPos, 1) = Trim(R!cDescrip)
            MSFVar.TextMatrix(nPos, 2) = Trim(R!cAbrev)
            nPos = nPos + 1
            R.MoveNext
        Loop
    R.Close
    Call CargaComboVal
End Sub
Private Function Valida_InsVariable() As Boolean
    If Len(Trim(Me.TxtValor.Text)) = 0 Or Len(Trim(txtDescrip.Text)) = 0 Then
        Valida_InsVariable = False
    Else
        Valida_InsVariable = True
    End If
End Function
Private Sub HabDesBotonesVarIns(ByVal bTxtdes As Boolean, ByVal bTxtvalor As Boolean, ByVal bAceptar As Boolean, ByVal bCancelar As Boolean)
    txtDescrip.Enabled = bTxtdes
    'txtDescrip.Text = ""
    TxtValor.Enabled = bTxtvalor
    'TxtValor.Text = ""
    CmdAceptar.Enabled = bAceptar
    CmdCancelar.Enabled = bCancelar
End Sub

Private Sub CboVar_Click()
    If CboVar.ListIndex >= 0 Then
        Call CargaDatosVal(CInt(Trim(Right(CboVar.Text, 2))))
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim sSql As String
    If Valida_InsVariable Then
        If nAccion = 1 Then
                sSql = "Insert Into VarCtasCont(cDescrip,cAbrev) " _
                       & " VALUES('" & txtDescrip.Text & "','" & TxtValor.Text & "')"
                oCon.Ejecutar sSql
        Else
                sSql = "UPDATE VarCtasCont Set cDescrip = '" & txtDescrip.Text & "', cAbrev = '" & TxtValor.Text & "' Where nCodigo = " & Trim(Str(nCod))
                oCon.Ejecutar sSql
        End If
    Else
        MsgBox "Datos No validos", vbInformation, "Aviso"
    End If
    Call HabDesBotonesVarIns(False, False, False, False)
    Call CargaDatosIns
    txtDescrip.Text = ""
    TxtValor.Text = ""
End Sub

Private Sub CmdAceptarVal_Click()
Dim sSql As String
Dim P As Integer
    If nAccion = 1 Then
        sSql = "Insert Into ValVarCtas(nCodigo,cCodValor,cValor,cDescrip) " _
             & " Values(" & Trim(Right(CboVar.Text, 2)) & ",'" & Trim(txtVarCod.Text) & "','" & Trim(txtVarValor.Text) & "','" & txtVarDesc.Text & "')"
        oCon.Ejecutar sSql
    Else
        If EliminaVarCuentas(CLng(Trim(Right(CboVar.Text, 5))), MSFValVar.TextMatrix(MSFValVar.Row, 3)) Then
            sSql = "Update ValVarCtas Set cValor = '" & Trim(txtVarValor.Text) & "', cDescrip = '" & txtVarDesc.Text & "' " _
                   & " Where nCodigo = " & Trim(Right(CboVar.Text, 2)) & " And cCodValor = '" & Trim(MSFValVar.TextMatrix(nPosVal, 1)) & "'"
            oCon.Ejecutar sSql
            Call GeneraVarCuentasModificadas(CLng(Trim(Right(CboVar.Text, 5))), Trim(txtVarValor.Text))
        End If
    End If
    Call HabDesBotonVarVal(False, False, False, False, False)
    P = CboVar.ListIndex
    CboVar.ListIndex = -1
    CboVar.ListIndex = P
    txtVarValor.Text = ""
    txtVarCod.Text = ""
    txtVarDesc.Text = ""
End Sub

Private Sub CmdAddRest_Click()
Dim sSql As String
    sSql = "INSERT INTO RestVarCta(nCodigo,cCodValor,nCodRes,cCodValres) Values (" _
         & Trim(LstVar.SelectedItem.SubItems(2)) & ",'" & Trim(LstVar.SelectedItem.SubItems(3)) & "'," _
         & Trim(LstValVar.SelectedItem.SubItems(2)) & ",'" & Trim(LstValVar.SelectedItem.SubItems(3)) & "')"
    oCon.Ejecutar sSql
    Call CargaDatosRestricciones(CLng(LstVar.SelectedItem.SubItems(2)), LstVar.SelectedItem.SubItems(3))
End Sub

Private Sub cmdCancelar_Click()
    Call HabDesBotonesVarIns(False, False, False, False)
    txtDescrip.Text = ""
    TxtValor.Text = ""
End Sub

Private Sub CmdCancelarVal_Click()
    Call HabDesBotonVarVal(False, False, False, False, False)
    txtVarValor.Text = ""
    txtVarCod.Text = ""
    txtVarDesc.Text = ""
End Sub

Private Sub CmdEdiVal_Click()
    nPosVal = MSFValVar.Row
    Call HabDesBotonVarVal(False, True, True, True, True)
    nAccion = 2
End Sub


Private Sub cmdEliminar_Click()
Dim sSql As String
If lstCtas.ListCount = 0 Then
    MsgBox "No existen datos para eliminar", vbInformation, "¡Aviso!"
    Exit Sub
End If

    If MsgBox("Esta Seguro Que Desea Eliminar la Clase y Todas sus Divisionarias?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        SBEstado.Panels(1).Text = "Generando Cuentas.."
        Screen.MousePointer = 11
        sSql = "DELETE CtasGen Where SubString(cClase,1,2)='" & Mid(Trim(lstCtas.Text), 1, 2) & "'"
        oCon.Ejecutar sSql
        Call CargaDatosCtasVar
        Screen.MousePointer = 0
        SBEstado.Panels(1).Text = "Proceso Finalizado..."
    End If
End Sub

Private Sub CmdElimVar_Click()
Dim sSql As String
If MSFVar.TextMatrix(1, 1) = "" Then
    MsgBox "No existen datos para eliminar", vbInformation, "¡Aviso!"
    Exit Sub
End If

    If MsgBox("Desea Eliminar el Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        sSql = "Delete VarCtasCont Where nCodigo = " & Trim(Str(MSFVar.TextMatrix(MSFVar.RowSel, 0)))
        oCon.Ejecutar sSql
        Call CargaDatosIns
    End If
End Sub

Private Sub CmdEliVal_Click()
Dim sSql As String
Dim nPos As Integer
If MSFValVar.TextMatrix(1, 1) = "" Then
    MsgBox "No existen datos para eliminar", vbInformation, "¡Aviso!"
    Exit Sub
End If
    If MsgBox("Desea Eliminar el registro ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        'If EliminaVarCuentas(CLng(Trim(Right(CboVar.Text, 5))), MSFValVar.TextMatrix(MSFValVar.Row, 3)) Then
            sSql = "DELETE ValVarCtas Where nCodigo = " & Trim(Right(CboVar.Text, 2)) & " And cCodValor = '" & Trim(MSFValVar.TextMatrix(MSFValVar.Row, 1)) & "'"
            oCon.Ejecutar sSql
        'End If
        
        nPos = CboVar.ListIndex
        CboVar.ListIndex = -1
        CboVar.ListIndex = nPos
    End If
End Sub

Private Sub CmdGenVal_Click()
    Call GeneraVarCuentasModificadas(CLng(Trim(Right(CboVar.Text, 5))), MSFValVar.TextMatrix(MSFValVar.Row, 3))
End Sub

Private Sub CmdModifCta_Click()
    nAccion = 2
    nCod = CInt(MSFVar.TextMatrix(MSFVar.RowSel, 0))
    Call HabDesBotonesVarIns(True, True, True, True)
End Sub


Private Sub CmdNewVal_Click()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim NumMax As Integer
Dim sCod As String
    sSql = "Select Max(convert(Int,cCodValor)) as nCodValor from ValVarCtas where nCodigo = '" & Trim(Right(CboVar.Text, 5)) & "' "
    Set R = oCon.CargaRecordSet(sSql)
        NumMax = IIf(IsNull(R!nCodValor), 0, R!nCodValor)
    R.Close
    NumMax = NumMax + 1
    sCod = Right("0" & Trim(Str(NumMax)), 2)
    nAccion = 1
    nPosVal = MSFVar.RowSel
    Call HabDesBotonVarVal(True, True, True, True, True)
    'txtVarCod.Text = Right(CboVar.Text, 2)
    txtVarDesc.Text = Trim(Mid(CboVar.Text, 1, Len(CboVar.Text) - 2))
    txtVarCod.Text = sCod
    txtVarDesc.Text = ""
    Call fEnfoque(txtVarDesc)
    txtVarDesc.SetFocus
End Sub

Private Sub CmdNuevaVar_Click()
    nAccion = 1
    Call HabDesBotonesVarIns(True, True, True, True)
    txtDescrip.Text = ""
    TxtValor.Text = ""
    txtDescrip.SetFocus
End Sub

Private Sub cmdNuevo_Click()
Dim R As ADODB.Recordset
Dim sSql As String
Dim lsCta As String
    sSql = "Select cClase FROM CtasGen WHERE '" & lstCtas.Text & "%' LIKE cClase + '%'"
    Set R = oCon.CargaRecordSet(sSql)
    If Not R.EOF Then
        lsCta = R!cClase
    Else
        lsCta = ""
    End If
    frmContabManVarNew.Inicio lsCta
    Call CargaDatosCtasVar
End Sub

Private Sub CmdRemRest_Click()
Dim sSql As String
    If LstRest.ListItems.Count > 0 Then
        sSql = "DELETE RestVarCta Where nCodigo = " & LstVar.SelectedItem.SubItems(2) _
               & " And cCodValor = '" & LstVar.SelectedItem.SubItems(3) _
               & "' And nCodRes = " & LstRest.SelectedItem.SubItems(2) _
               & " and cCodValRes = '" & LstRest.SelectedItem.SubItems(3) & "'"
        oCon.Ejecutar sSql
    End If
    Call CargaDatosRestricciones(CLng(LstVar.SelectedItem.SubItems(2)), LstVar.SelectedItem.SubItems(3))
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Set oCon = New DConecta
    oCon.AbreConexion
    CargaCabeceraVar
    CargaCabeceraIns
    CargaDatosIns
    CargaDatosCtasVar
    Call CargaDatosLstVar(LstVar)
    Call CargaDatosLstVar(LstValVar)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oCon.CierraConexion
    Set oCon = Nothing
End Sub

Private Sub LstVar_Click()
    Call CargaDatosRestricciones(CLng(LstVar.SelectedItem.SubItems(2)), LstVar.SelectedItem.SubItems(3))
End Sub

Private Sub LstVar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        Call CargaDatosRestricciones(CLng(LstVar.SelectedItem.SubItems(2)), LstVar.SelectedItem.SubItems(3))
    End If
End Sub

Private Sub MSFValVar_Click()
    Call MuestraFilaVal
    Call SelecFila(MSFValVar, MSFValVar.Row, 0)
End Sub

Private Sub MSFVar_Click()
    Call MuestraFilaSelIns(MSFVar.RowSel)
    Call SelecFila(MSFVar, MSFVar.Row, 0)
End Sub
Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        fEnfoque TxtValor
        TxtValor.SetFocus
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub txtVarCod_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtVarDesc.SetFocus
    End If
End Sub

Private Sub txtVarDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        txtVarValor.SetFocus
    End If
End Sub

Private Sub txtVarValor_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptarVal.SetFocus
    End If
End Sub
