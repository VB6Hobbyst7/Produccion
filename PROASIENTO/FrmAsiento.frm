VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmAsiento 
   Caption         =   "Consulta de Asientos"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FrmAsiento.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmAsiento.frx":030A
   ScaleHeight     =   6825
   ScaleWidth      =   11880
   Begin TabDlg.SSTab tabCuentas 
      Height          =   4935
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      TabCaption(0)   =   "Asiento Contable"
      TabPicture(0)   =   "FrmAsiento.frx":0694
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LstAsiento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Asiento Contable 2"
      TabPicture(1)   =   "FrmAsiento.frx":06B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lswAsiento2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   -72960
         TabIndex        =   22
         Top             =   4200
         Width           =   7215
         Begin VB.CommandButton cmdEportar2 
            Caption         =   "&Exportar Excel"
            Height          =   375
            Left            =   5880
            TabIndex        =   23
            Top             =   160
            Width           =   1215
         End
         Begin VB.Label lblHaber2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
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
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4080
            TabIndex        =   27
            Top             =   160
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total  Haber "
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
            Left            =   2880
            TabIndex        =   26
            Top             =   220
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total  Debe"
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
            Left            =   120
            TabIndex        =   25
            Top             =   220
            Width           =   1020
         End
         Begin VB.Label lblDebe2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
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
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1200
            TabIndex        =   24
            Top             =   160
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   2040
         TabIndex        =   17
         Top             =   4200
         Width           =   7215
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Exportar Excel"
            Height          =   375
            Left            =   5880
            TabIndex        =   10
            Top             =   160
            Width           =   1215
         End
         Begin VB.Label LblDebe 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
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
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            Top             =   160
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total  Debe"
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
            Left            =   120
            TabIndex        =   19
            Top             =   220
            Width           =   1020
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total  Haber "
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
            Left            =   2880
            TabIndex        =   18
            Top             =   220
            Width           =   1140
         End
         Begin VB.Label LblHaber 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
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
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4080
            TabIndex        =   9
            Top             =   160
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView LstAsiento 
         Height          =   3615
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cta Contable"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Haber"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuenta Cliente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Codigo Ope"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Desc Operacion"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Descrip Movimiento"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Usuario"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fecha Reg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Doc Custodia"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswAsiento2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Cta Contable"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debe"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Haber"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuenta Cliente"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Codigo Ope"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Desc Operacion"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Descrip Movimiento"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Usuario"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Fecha Reg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Doc Custodia"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CheckBox chkSinAgencia 
      Caption         =   "Sin Agencia"
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox TxtCtaCon 
         Height          =   285
         Left            =   8520
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker TxtFechaF 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72024065
         CurrentDate     =   38986
      End
      Begin MSComCtl2.DTPicker TxtFechaI 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72024065
         CurrentDate     =   38986
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Al: "
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CtaContable"
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   14
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencia"
         Height          =   195
         Index           =   0
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del: "
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   330
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "DETALLE DE ASIENTO CONTABLE"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   11670
   End
End
Attribute VB_Name = "FrmAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCon As COMConecta.DCOMConecta

Private Sub chkSinAgencia_Click()
 If chkSinAgencia.Value = 1 Then
    CboAgencia.Enabled = False
 Else
    CboAgencia.Enabled = True
 End If
End Sub

Private Sub CmdBuscar_Click()
    Dim sAge As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lst As ListItem
    Dim nSumD As Double
    Dim nSumTH As Double
    Dim nSumTD As Double
    Dim nSumH As Double
    
    Dim sql1 As String
    Dim rs1 As New ADODB.Recordset
    Dim lst1 As ListItem
    Dim nSumD1 As Double
    Dim nSumTH1 As Double
    Dim nSumTD1 As Double
    Dim nSumH1 As Double
    Dim sAnio As String
    Dim dFecha As Date
    nSumD = 0
    nSumH = 0
    nSumD1 = 0
    nSumH1 = 0
    
    If CDate(TxtFechaI.Value) > CDate(TxtFechaF.Value) Then
        MsgBox "La Fecha Final no puede ser mayor", vbInformation, "Aviso"
        Exit Sub
    End If
    'RIRO20160123 INC1601210006
    If Year(TxtFechaI.Value) > Year(Now) Then
        MsgBox "El año seleccionado debe ser igual o menor al actual", vbInformation, "Aviso"
        Exit Sub
    End If
    'END RIRO
    
    If CDate(TxtFechaI.Value) Then
        sAge = Trim(Right(CboAgencia.Text, 2))
        LstAsiento.ListItems.Clear
        
        'RIRO20160123 INC1601210006 , comentado
        'Add by Gitu 16-12-2009
        'If Month(TxtFechaI.Value) <> Month(Now) And Year(TxtFechaI.Value) <> Year(Now) Then
        '    sAnio = Year(TxtFechaI.Value)
        'End If
        'End Gitu
        
        'RIRO20160123 INC1601210006, add
        If Year(TxtFechaI.Value) <> Year(Now) Then
            sAnio = Year(TxtFechaI.Value)
        End If
        'END RIRO
        
        sql = " select dFecha,cCtaCnt,nDebe,nHaber,A.cCtaCod,O.cOpeCod,cOpeDesc Ope, " & _
               " cMovDesc,right(cMovNro,4)Usu,  " & _
               " (Select cDocNro from MovDoc MD where MD.nMovNro = M.nMovNro) NroDoc,  " & _
               " C.dRegistro, C.cNroDoc NroCustodia " & _
               " from AsientoDN" & sAnio & " A " & _
               " inner join  OpeTpo O ON A.copecod=O.copecod " & _
               " inner join Mov M ON M.nMovNro= A.nMovNro " & _
               " LEFT join MovOpevarias MO on MO.nMovNro=A.nMovNro and MO.cOpeCod='300408' " & _
               " LEFT join ColocacConvenioRegDev C ON C.NMOVNROREG=MO.nMovNro " & _
               " where datediff(day, '" & Format(TxtFechaI.Value, "MM/DD/YYYY ") & "', dfecha)>=0 and datediff(day, '" & Format(TxtFechaF.Value, "MM/DD/YYYY ") & "', dfecha)<=0 " & _
               " and cCtaCnt like '" & Trim(TxtCtaCon) & "%' and nMovFlaG =0"
        If Mid(TxtCtaCon.Text, 3, 1) = 2 Then
              sql = sql & " and ctipo='3'"
        End If
               
        If chkSinAgencia.Value = 0 Then
            If Trim(Right(CboAgencia, 2)) <> 0 Then
              sql = sql & " and A.cCodAge='" & sAge & "'"
            End If
        End If
        sql = sql & " order by dfecha"
        Set rs = oCon.CargaRecordSet(sql)
        dFecha = TxtFechaI.Value

' JEOM  Segundo Tab  Cuentas sin MovNRo  ------------------------------------------------------------
        sql1 = " select dFecha,cCtaCnt,nDebe,nHaber,A.cCtaCod,O.cOpeCod,cOpeDesc Ope, " & _
               " cMovDesc,right(cMovNro,4)Usu,  " & _
               " (Select cDocNro from MovDoc MD where MD.nMovNro = M.nMovNro) NroDoc,  " & _
               " C.dRegistro, C.cNroDoc NroCustodia " & _
               " from AsientoDN" & sAnio & " A " & _
               " left join  OpeTpo O ON A.copecod=O.copecod " & _
               " left join Mov M ON M.nMovNro= A.nMovNro " & _
               " LEFT join MovOpevarias MO on MO.nMovNro=A.nMovNro and MO.cOpeCod='300408' " & _
               " LEFT join ColocacConvenioRegDev C ON C.NMOVNROREG=MO.nMovNro " & _
               " where datediff(day, '" & Format(TxtFechaI.Value, "MM/DD/YYYY ") & "', dfecha)>=0 and datediff(day, '" & Format(TxtFechaF.Value, "MM/DD/YYYY ") & "', dfecha)<=0 " & _
               " and cCtaCnt like '" & Trim(TxtCtaCon) & "%' and  m.nmovnro is null"
        If Mid(TxtCtaCon.Text, 3, 1) = 2 Then
              sql1 = sql1 & " and ctipo='3'"
        End If
               
        If chkSinAgencia.Value = 0 Then
            If Trim(Right(CboAgencia, 2)) <> 0 Then
              sql1 = sql1 & " and A.cCodAge='" & sAge & "'"
            End If
        End If
        sql1 = sql1 & " order by dfecha"
        Set rs1 = oCon.CargaRecordSet(sql1)

'Fin------------------------------------------------------------------------------------------------

        If Not rs.EOF And Not rs.BOF Then
        lswAsiento2.ListItems.Clear
'SACR: 06/12/06 -------------------------------------------
            Do While Not rs.EOF
                If Not rs.BOF Then
                    If rs.Bookmark = rs.RecordCount Then
                        Set lst = LstAsiento.ListItems.Add(, , rs(0))
                        lst.SubItems(1) = rs(1)
                        lst.SubItems(2) = Format(CDbl(rs(2)), "0.00")
                        lst.SubItems(3) = Format(CDbl(rs(3)), "0.00")
                        lst.SubItems(4) = IIf(IsNull(rs(4)), "", rs(4))
                        lst.SubItems(5) = rs(5)
                        lst.SubItems(6) = rs(6)
                        lst.SubItems(7) = rs(7)
                        lst.SubItems(8) = rs(8)
                        lst.SubItems(9) = IIf(IsNull(rs(9)), "", rs(9))   ' NroDoc
                        lst.SubItems(10) = IIf(IsNull(rs(10)), "", rs(10))
                        lst.SubItems(11) = IIf(IsNull(rs(11)), "", rs(11))
                        nSumD = nSumD + lst.SubItems(2)
                        nSumH = nSumH + lst.SubItems(3)
                    End If
                End If
            If dFecha <> Format(rs(0), "dd/mm/yyyy") Or rs.Bookmark = rs.RecordCount Then
                        Set lst = LstAsiento.ListItems.Add(, , "")
                        lst.SubItems(1) = "SUBTOTAL"
                        lst.SubItems(2) = Format(CDbl(nSumD), "0.00")
                        lst.SubItems(3) = Format(CDbl(nSumH), "0.00")
                        lst.SubItems(4) = ""
                        lst.SubItems(5) = ""
                        lst.SubItems(6) = ""
                        lst.SubItems(7) = ""
                        lst.SubItems(8) = ""
                        lst.SubItems(9) = ""
                        lst.SubItems(10) = ""
                        lst.SubItems(11) = ""
                        nSumTD = nSumTD + nSumD
                        nSumTH = nSumTH + nSumH
                        nSumD = 0
                        nSumH = 0
                        dFecha = Format(rs(0), "dd/mm/yyyy")
                        If rs.Bookmark < rs.RecordCount Then
                            rs.MovePrevious
                        End If
'                    End If
                Else
                    Set lst = LstAsiento.ListItems.Add(, , rs(0))
                    lst.SubItems(1) = rs(1)
                    lst.SubItems(2) = Format(CDbl(rs(2)), "0.00")
                    lst.SubItems(3) = Format(CDbl(rs(3)), "0.00")
                    lst.SubItems(4) = IIf(IsNull(rs(4)), "", rs(4))
                    lst.SubItems(5) = rs(5)
                    lst.SubItems(6) = rs(6)
                    lst.SubItems(7) = rs(7)
                    lst.SubItems(8) = rs(8)
                    lst.SubItems(9) = IIf(IsNull(rs(9)), "", rs(9))   ' NroDoc
                    lst.SubItems(10) = IIf(IsNull(rs(10)), "", rs(10))
                    lst.SubItems(11) = IIf(IsNull(rs(11)), "", rs(11))
                    nSumD = nSumD + lst.SubItems(2)
                    nSumH = nSumH + lst.SubItems(3)
                End If
                rs.MoveNext
            Loop
        End If

        LblDebe.Caption = Format(CDbl(nSumTD), "0.00")
        LblHaber.Caption = Format(CDbl(nSumTH), "#0.00")
  
  '  JEOM ---------------- segundo  ListView ----
     If Not rs1.EOF And Not rs1.BOF Then
            Do While Not rs1.EOF
                If Not rs1.BOF Then
                    If rs1.Bookmark = rs1.RecordCount Then
                        Set lst1 = lswAsiento2.ListItems.Add(, , rs1(0))
                        lst1.SubItems(1) = rs1(1)
                        lst1.SubItems(2) = Format(CDbl(rs1(2)), "0.00")
                        lst1.SubItems(3) = Format(CDbl(rs1(3)), "0.00")
                        lst1.SubItems(4) = IIf(IsNull(rs1(4)), "", rs1(4))
                        lst1.SubItems(5) = IIf(IsNull(rs1(5)), "", rs1(5))
                        lst1.SubItems(6) = IIf(IsNull(rs1(6)), "", rs1(6))
                        lst1.SubItems(7) = IIf(IsNull(rs1(7)), "", rs1(7))
                        lst1.SubItems(8) = IIf(IsNull(rs1(8)), "", rs1(8))
                        lst1.SubItems(9) = IIf(IsNull(rs1(9)), "", rs1(9))   ' NroDoc
                        lst1.SubItems(10) = IIf(IsNull(rs1(10)), "", rs1(10))
                        lst1.SubItems(11) = IIf(IsNull(rs1(11)), "", rs1(11))
                        nSumD1 = nSumD1 + lst1.SubItems(2)
                        nSumH1 = nSumH1 + lst1.SubItems(3)
                    End If
                End If
            If dFecha <> Format(rs1(0), "dd/mm/yyyy") Or rs1.Bookmark = rs1.RecordCount Then
                        Set lst1 = lswAsiento2.ListItems.Add(, , "")
                        lst1.SubItems(1) = "SUBTOTAL"
                        lst1.SubItems(2) = Format(CDbl(nSumD1), "0.00")
                        lst1.SubItems(3) = Format(CDbl(nSumH1), "0.00")
                        lst1.SubItems(4) = ""
                        lst1.SubItems(5) = ""
                        lst1.SubItems(6) = ""
                        lst1.SubItems(7) = ""
                        lst1.SubItems(8) = ""
                        lst1.SubItems(9) = ""
                        lst1.SubItems(10) = ""
                        lst1.SubItems(11) = ""
                        nSumTD1 = nSumTD1 + nSumD1
                        nSumTH1 = nSumTH1 + nSumH1
                        nSumD1 = 0
                        nSumH1 = 0
                        dFecha = Format(rs1(0), "dd/mm/yyyy")
                        If rs1.Bookmark < rs1.RecordCount Then
                            rs1.MovePrevious
                        End If
                Else
                    Set lst1 = lswAsiento2.ListItems.Add(, , rs1(0))
                    lst1.SubItems(1) = rs1(1)
                    lst1.SubItems(2) = Format(CDbl(rs1(2)), "0.00")
                    lst1.SubItems(3) = Format(CDbl(rs1(3)), "0.00")
                    lst1.SubItems(4) = IIf(IsNull(rs1(4)), "", rs1(4))
                    lst1.SubItems(5) = IIf(IsNull(rs1(5)), "", rs1(5))
                    lst1.SubItems(6) = IIf(IsNull(rs1(6)), "", rs1(6))
                    lst1.SubItems(7) = IIf(IsNull(rs1(7)), "", rs1(7))
                    lst1.SubItems(8) = IIf(IsNull(rs1(8)), "", rs1(8))
                    lst1.SubItems(9) = IIf(IsNull(rs1(9)), "", rs1(9))   ' NroDoc
                    lst1.SubItems(10) = IIf(IsNull(rs1(10)), "", rs1(10))
                    lst1.SubItems(11) = IIf(IsNull(rs1(11)), "", rs1(11))
                    nSumD1 = nSumD1 + lst1.SubItems(2)
                    nSumH1 = nSumH1 + lst1.SubItems(3)
                End If
                rs1.MoveNext
            Loop
        End If
        
        lblDebe2.Caption = Format(CDbl(nSumTD1), "0.00")
        lblHaber2.Caption = Format(CDbl(nSumTH1), "#0.00")
        
       tabCuentas.TabVisible(1) = True
  
  '----------- Fin---- JEOM
        
    Else
        MsgBox "Error"
    End If
End Sub

Private Sub cmdEportar2_Click()
    Dim fs As Scripting.FileSystemObject
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim nLineaInicio As Integer
    Dim nLineas As Integer
    Dim nLineasTemp As Integer
    
    Dim i As Integer
    Dim nTotal As Double
    
    Dim glsarchivo As String
    
    
    glsarchivo = "Rep_ProAsiento2_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add


    xlAplicacion.Range("A1").ColumnWidth = 15
    xlAplicacion.Range("B1").ColumnWidth = 13
    xlAplicacion.Range("C1").ColumnWidth = 10
    xlAplicacion.Range("D1").ColumnWidth = 11
    xlAplicacion.Range("E1").ColumnWidth = 19
    xlAplicacion.Range("F1").ColumnWidth = 10
    xlAplicacion.Range("G1").ColumnWidth = 35
    xlAplicacion.Range("H1").ColumnWidth = 16
    xlAplicacion.Range("I1").ColumnWidth = 7.5
    xlAplicacion.Range("J1").ColumnWidth = 11
    
    xlAplicacion.Range("A5").Value = "Fecha"
    xlAplicacion.Range("B5").Value = "Cta Contable"
    xlAplicacion.Range("C5").Value = "Debe"
    xlAplicacion.Range("D5").Value = "Haber"
    xlAplicacion.Range("E5").Value = "Cuenta Cliente"
    xlAplicacion.Range("F5").Value = "Codigo Ope"
    xlAplicacion.Range("G5").Value = "Desc Operación"
    xlAplicacion.Range("H5").Value = "Descrip Movimiento"
    xlAplicacion.Range("I5").Value = "Usuario"
    xlAplicacion.Range("J5").Value = "Documento"
    xlAplicacion.Range("A5:K5").Font.Bold = True
    
    nLineas = 1
    xlHoja1.Cells(nLineas, 5) = "REPORTE PRO ASIENTO 2"
    xlAplicacion.Range("E1").Font.Bold = True
    
    nLineas = nLineas + 4
    With lswAsiento2
        For i = 1 To .ListItems.Count
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, 1) = .ListItems(i).Text
            xlHoja1.Cells(nLineas, 2) = .ListItems(i).SubItems(1)
            xlHoja1.Cells(nLineas, 3) = Format(.ListItems(i).SubItems(2), "#0.00")
            xlHoja1.Cells(nLineas, 4) = Format(.ListItems(i).SubItems(3), "#0.00")
            xlHoja1.Cells(nLineas, 5) = "'" & .ListItems(i).SubItems(4) & ""
            xlHoja1.Cells(nLineas, 6) = .ListItems(i).SubItems(5)
            xlHoja1.Cells(nLineas, 7) = .ListItems(i).SubItems(6)
            xlHoja1.Cells(nLineas, 8) = .ListItems(i).SubItems(7)
            xlHoja1.Cells(nLineas, 9) = .ListItems(i).SubItems(8)
            xlHoja1.Cells(nLineas, 10) = .ListItems(i).SubItems(9)
            xlHoja1.Cells(nLineas, 11) = .ListItems(i).SubItems(10)
            xlHoja1.Cells(nLineas, 12) = .ListItems(i).SubItems(11)
        Next
    End With
    nLineas = nLineas + 1
 
    
    nLineas = nLineas + 4
    xlHoja1.Cells(nLineas, 2) = "Huancayo " & Day(gdFecSis) & " de " & Format(gdFecSis, "MMMM") & " del " & Year(gdFecSis)
    
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(3, 1)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(nLineas, 6)).Font.Size = 9
    
    'xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsarchivo
               
    MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsarchivo, vbInformation, "Mensaje"
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
        
    Set xlAplicacion = Nothing

End Sub

Private Sub cmdImprimir_Click()
    Dim fs As Scripting.FileSystemObject
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim nLineaInicio As Integer
    Dim nLineas As Integer
    Dim nLineasTemp As Integer
    
    Dim i As Integer
    Dim nTotal As Double
    
    Dim glsarchivo As String
    
    
    glsarchivo = "Rep_ProAsiento_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    Set xlHoja1 = xlLibro.Worksheets.Add


    xlAplicacion.Range("A1").ColumnWidth = 15
    xlAplicacion.Range("B1").ColumnWidth = 13
    xlAplicacion.Range("C1").ColumnWidth = 10
    xlAplicacion.Range("D1").ColumnWidth = 11
    xlAplicacion.Range("E1").ColumnWidth = 19
    xlAplicacion.Range("F1").ColumnWidth = 10
    xlAplicacion.Range("G1").ColumnWidth = 35
    xlAplicacion.Range("H1").ColumnWidth = 16
    xlAplicacion.Range("I1").ColumnWidth = 7.5
    xlAplicacion.Range("J1").ColumnWidth = 11
    
    xlAplicacion.Range("A5").Value = "Fecha"
    xlAplicacion.Range("B5").Value = "Cta Contable"
    xlAplicacion.Range("C5").Value = "Debe"
    xlAplicacion.Range("D5").Value = "Haber"
    xlAplicacion.Range("E5").Value = "Cuenta Cliente"
    xlAplicacion.Range("F5").Value = "Codigo Ope"
    xlAplicacion.Range("G5").Value = "Desc Operación"
    xlAplicacion.Range("H5").Value = "Descrip Movimiento"
    xlAplicacion.Range("I5").Value = "Usuario"
    xlAplicacion.Range("J5").Value = "Documento"
    xlAplicacion.Range("A5:K5").Font.Bold = True
    
    nLineas = 1
    xlHoja1.Cells(nLineas, 5) = "REPORTE PRO ASIENTO"
    xlAplicacion.Range("E1").Font.Bold = True
    
    nLineas = nLineas + 4
    With LstAsiento
        For i = 1 To .ListItems.Count
            nLineas = nLineas + 1
            xlHoja1.Cells(nLineas, 1) = .ListItems(i).Text
            xlHoja1.Cells(nLineas, 2) = .ListItems(i).SubItems(1)
            xlHoja1.Cells(nLineas, 3) = Format(.ListItems(i).SubItems(2), "#0.00")
            xlHoja1.Cells(nLineas, 4) = Format(.ListItems(i).SubItems(3), "#0.00")
            xlHoja1.Cells(nLineas, 5) = "'" & .ListItems(i).SubItems(4) & ""
            xlHoja1.Cells(nLineas, 6) = .ListItems(i).SubItems(5)
            xlHoja1.Cells(nLineas, 7) = .ListItems(i).SubItems(6)
            xlHoja1.Cells(nLineas, 8) = .ListItems(i).SubItems(7)
            xlHoja1.Cells(nLineas, 9) = .ListItems(i).SubItems(8)
            xlHoja1.Cells(nLineas, 10) = .ListItems(i).SubItems(9)
            xlHoja1.Cells(nLineas, 11) = .ListItems(i).SubItems(10)
            xlHoja1.Cells(nLineas, 12) = .ListItems(i).SubItems(11)
        Next
    End With
    nLineas = nLineas + 1
 
    
    nLineas = nLineas + 4
    xlHoja1.Cells(nLineas, 2) = "Huancayo " & Day(gdFecSis) & " de " & Format(gdFecSis, "MMMM") & " del " & Year(gdFecSis)
    
    xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(3, 1)).Font.Size = 10
    xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(nLineas, 6)).Font.Size = 9
    
    'xlHoja1.SaveAs App.Path & "\SPOOLER\" & glsarchivo
               
    MsgBox "Se ha generado el Archivo en " & App.Path & "\SPOOLER\" & glsarchivo, vbInformation, "Mensaje"
    xlAplicacion.Visible = True
    xlAplicacion.Windows(1).Visible = True
        
    Set xlAplicacion = Nothing

End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Desea salir del formulario?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
   Set oCon = New COMConecta.DCOMConecta
   oCon.AbreConexion
   cargaagencias
   TxtFechaI = Date
   TxtFechaF = Date
End Sub

Public Sub cargaagencias()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "select cAgeDescripcion,cAgeCod from dbo.Agencias"
    Set rs = oCon.CargaRecordSet(sql)
    CboAgencia.Clear
    If Not rs.EOF And Not rs.BOF Then
        Do Until rs.EOF
            CboAgencia.AddItem rs(0) & Space(100) & rs(1)
            rs.MoveNext
        Loop
    End If
    CboAgencia.AddItem "TODOS" & Space(100) & "0"
    CboAgencia.ListIndex = 0
End Sub

Private Sub LstAsiento_Click()
    If LstAsiento.ListItems.Count > 0 Then
        FrmCuentas.Inicio LstAsiento.ListItems.Item(LstAsiento.SelectedItem.Index).SubItems(5)
    End If
End Sub


Private Sub TxtCtaCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     CmdBuscar.SetFocus
   End If
End Sub
