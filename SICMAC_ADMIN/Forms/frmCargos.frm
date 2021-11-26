VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCargos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargos"
   ClientHeight    =   6015
   ClientLeft      =   1635
   ClientTop       =   2730
   ClientWidth     =   7605
   Icon            =   "frmCargos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   5520
      TabIndex        =   12
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2280
      TabIndex        =   9
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   1200
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "E&liminar"
      Height          =   330
      Left            =   4440
      TabIndex        =   11
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdeditar 
      Caption         =   "E&ditar"
      Height          =   330
      Left            =   3360
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "N&uevo"
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   135
      Left            =   5520
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCargos.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   330
      Left            =   6600
      TabIndex        =   19
      Top             =   5640
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Niveles"
      TabPicture(0)   =   "frmCargos.frx":0530
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCargos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOrden"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMaximo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCadCod"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "EditMoney1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "EditMoney2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "MSHFlexGrid1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCardesc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtOrden"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbCatCod"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&Cargos"
      TabPicture(1)   =   "frmCargos.frx":054C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblcCarDes"
      Tab(1).Control(1)=   "lblCargo"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "lblSueldo"
      Tab(1).Control(4)=   "txtCarSue"
      Tab(1).Control(5)=   "txtcCarDes"
      Tab(1).Control(6)=   "txtcCarCod"
      Tab(1).Control(7)=   "cmbCargos"
      Tab(1).Control(8)=   "Flex"
      Tab(1).ControlCount=   9
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   14
         Top             =   840
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5741
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox cmbCatCod 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   5040
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtOrden 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   4680
         Width           =   615
      End
      Begin VB.TextBox txtCardesc 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   3480
         Width           =   6135
      End
      Begin VB.ComboBox cmbCargos 
         Height          =   315
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   6480
      End
      Begin VB.TextBox txtcCarCod 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   -74040
         TabIndex        =   15
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtcCarDes 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74040
         MaxLength       =   255
         TabIndex        =   16
         Top             =   4560
         Width           =   5775
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5106
         _Version        =   393216
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin SisCon.EditMoney EditMoney2 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   4260
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
      End
      Begin SisCon.EditMoney EditMoney1 
         Height          =   390
         Left            =   1320
         TabIndex        =   3
         Top             =   3840
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   688
      End
      Begin SisCon.EditMoney txtCarSue 
         Height          =   360
         Left            =   -74040
         TabIndex        =   17
         Top             =   4920
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   635
      End
      Begin VB.Label lblSueldo 
         Caption         =   "Sueldo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   4973
         Width           =   735
      End
      Begin VB.Label lblCadCod 
         Caption         =   "Categorias"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5070
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblMaximo 
         Caption         =   "Sueldo.Max."
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblOrden 
         Caption         =   "Orden"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Sueldo.Niv"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3908
         Width           =   855
      End
      Begin VB.Label lblCargos 
         Caption         =   "Descripcripción"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3525
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nivel :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   510
         Width           =   735
      End
      Begin VB.Label lblCargo 
         Caption         =   "Cargos:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   4230
         Width           =   615
      End
      Begin VB.Label lblcCarDes 
         Caption         =   "Descrip."
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   4590
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lbEditado As Boolean
Dim lbEditadoGen As Boolean
Dim lsCodigo As String
Dim lnIniOpe As Integer

Private Sub cmbCargos_Click()
   Dim sqlC1 As String
   Dim rsC1 As ADODB.Recordset
   Set rsC1 = New ADODB.Recordset
   
    sqlC1 = "Select cCarCod as Codigo,cCarDes as Descripcion, convert(numeric(8,2),cCarSue) as Sueldo, convert(numeric(8,2),nMonMaximo) as Sue_Maximo,nOrden as Orden,cCatCod From Cargos where cCarCod like '" & Left(Me.cmbCargos, 3) & "___' and cCarCod not like '" & Left(cmbCargos.Text, 3) & "' order by cCarCod"
    rsC1.open sqlC1, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not RSVacio(rsC1) Then
        Set Me.Flex.DataSource = rsC1
    
        Flex.ColWidth(0) = 1
        Flex.ColWidth(1) = 1500
        Flex.ColWidth(2) = 3300
        Flex.ColWidth(3) = 1800
        Flex.ColWidth(4) = 1
        Flex.ColWidth(5) = 1
        Flex.ColWidth(6) = 1
        
        Flex.ColAlignment(3) = 7
        Flex.ColAlignment(4) = 7
        Flex.ColAlignment(5) = 7
    Else
        Flex.Rows = 1
        Flex.Rows = 2
        Flex.FixedRows = 1
    End If
    
    rsC1.Close
    Set rsC1 = Nothing
End Sub

Private Sub cmbCatCod_Change()
    GetData Left(cmbCatCod, 3)
End Sub

Private Sub cmbCatCod_Click()
    cmbCatCod_Change
End Sub

Private Sub cmbCatCod_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then Me.ctrMant1.SetFocus
    If KeyAscii = 13 Then Me.txtCardesc.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    Dim sqlC As String
    
    If Me.SSTab1.Tab = 0 Then
        If Me.EditMoney1.Value < 0 Then
            MsgBox "Monto no valido.", vbInformation, "Aviso"
            EditMoney1.SetFocus
            Exit Sub
        ElseIf Me.EditMoney2.Value < 0 Then
            MsgBox "Monto no valido.", vbInformation, "Aviso"
            EditMoney2.SetFocus
            Exit Sub
        ElseIf Trim(Me.txtCardesc) = "" Then
            MsgBox "Debe ingresar una descripción.", vbInformation, "Aviso"
            txtCardesc.SetFocus
            Exit Sub
        End If
        
        If SSTab1.Tab <> lnIniOpe Then
            If lnIniOpe = 0 Then
                MsgBox "Ud. ha elegido actualizar un Nivel, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            Else
                MsgBox "Ud. ha elegido actualizar un Cargo, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            End If
            SSTab1.Tab = lnIniOpe
            Exit Sub
        Else
            lnIniOpe = -1
        End If
        
        If lbEditadoGen = False Then
            sqlC = " Insert Cargos (cCarCod,cCarDes,cCarSue,nOrden,nMonMaximo,cCatCod)"
            sqlC = sqlC & " Values('" & lsCodigo & "','" & Me.txtCardesc.Text & "'," & Me.EditMoney1.Value & "," & Me.txtOrden & "," & Me.EditMoney2.Value & ",'')"
            dbCmact.execute sqlC
        Else
            sqlC = " Update Cargos Set cCarDes = '" & Me.txtCardesc.Text & "', cCarSue = " & Me.EditMoney1.Value & ", nOrden= " & Me.txtOrden & ",nMonMaximo = " & Me.EditMoney2.Value
            sqlC = sqlC & " Where cCarCod = '" & Trim(Me.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)) & "'"
            dbCmact.execute sqlC
        End If

        lbEditadoGen = False
        Me.txtCardesc.Text = ""
        EditMoney1.Value = 0
        EditMoney2.Value = 0
        Me.txtOrden.Text = "0"
        Me.MSHFlexGrid1.Enabled = True

        Refresca_C
    Else
        If Trim(Me.cmbCargos) = "" Then
            MsgBox "Debe elegir un nivel.", vbInformation, "Aviso"
            Me.cmbCargos.SetFocus
            Exit Sub
        ElseIf Trim(txtcCarDes) = "" Then
            MsgBox "Debe ingresar una descripción.", vbInformation, "Aviso"
            txtcCarDes.SetFocus
            Exit Sub
        End If
        
        If SSTab1.Tab <> lnIniOpe Then
            If lnIniOpe = 0 Then
                MsgBox "Ud. ha elegido actualizar un Nivel, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            Else
                MsgBox "Ud. ha elegido actualizar un Cargo, para poder actualizar un Nivel, debe terminar antes la operacion antes iniciada.", vbInformation, "Aviso"
            End If
            SSTab1.Tab = lnIniOpe
            Exit Sub
        Else
            lnIniOpe = -1
        End If
        
        If Not lbEditado Then
            sqlC = " Insert Cargos (cCarCod,cCarDes,cCarSue,nOrden,nMonMaximo,cCatCod)"
            sqlC = sqlC & " Values('" & Me.txtcCarCod.Text & "','" & Me.txtcCarDes.Text & "'," & Me.txtCarSue.Value & ",0,0,'')"
            dbCmact.execute sqlC
        Else
            sqlC = " Update Cargos Set cCarDes = '" & Me.txtcCarDes.Text & "', cCarSue = " & Me.txtCarSue.Value & ""
            sqlC = sqlC & " Where cCarCod = '" & Trim(Me.Flex.TextMatrix(Flex.Row, 1)) & "'"
            dbCmact.execute sqlC
        End If
        
        txtcCarCod.Text = ""
        txtcCarDes.Text = ""
        txtCarSue.Value = 0
        txtcCarDes.SetFocus
        cmbCargos_Click
    End If
    
    Activa True
End Sub

Private Sub cmdCancelar_Click()
    If Me.SSTab1.Tab = 0 Then
        lbEditadoGen = False
        Me.txtCardesc.Text = ""
        EditMoney1.Value = 0
        EditMoney2.Value = 0
        Me.txtOrden.Text = "0"
        Me.MSHFlexGrid1.Enabled = True
    Else
        txtcCarDes = ""
        txtCarSue.Value = 0
        txtCarSue.SetFocus
    End If
    
    Activa True
End Sub

Private Sub cmdeditar_Click()
    lnIniOpe = SSTab1.Tab
    If Me.SSTab1.Tab = 0 Then
        Me.txtCardesc.Text = MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 2)
        Me.EditMoney1.Value = MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 3)
        Me.EditMoney2.Value = MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 4)
        Me.txtOrden = MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5)
        lbEditadoGen = True
    Else
        Me.txtcCarCod.Text = Me.Flex.TextMatrix(Flex.Row, 1)
        Me.txtcCarDes.Text = Me.Flex.TextMatrix(Flex.Row, 2)
        Me.txtCarSue.Value = CCur(Me.Flex.TextMatrix(Flex.Row, 3))
        Me.txtcCarDes.SetFocus
        lbEditado = True
    End If
        
    Activa False
End Sub

Private Sub cmdEliminar_Click()
    Dim sqlD As String
    If Me.SSTab1.Tab = 0 Then
        If MsgBox("El Nivel y todos sus cargos van a ser eliminados, desea proseguir ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        sqlD = "Delete Cargos Where cCarCod like '" & Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 1) & "%'"
        dbCmact.execute sqlD
        Refresca_C
    Else
        If Trim(cmbCargos) = "" Then
            MsgBox "Tiene que indicar un nivel.", vbInformation, "Aviso"
            Me.cmbCargos.SetFocus
            Exit Sub
        End If
        
        If MsgBox("El Cargo va ser eliminado, desea proseguir ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        sqlD = "Delete Cargos Where cCarCod like '" & Me.Flex.TextMatrix(Me.Flex.Row, 1) & "'"
        dbCmact.execute sqlD
        cmbCargos_Click
    End If
    
End Sub

Private Sub cmdImprimir_Click()
    R.Text = GetRepo
    frmPrevio.Previo R, "Niveles y Cargos", True, 66
End Sub

Private Sub cmdNuevo_Click()
    Dim lsNumCur As String
    
    lnIniOpe = SSTab1.Tab
    If Me.SSTab1.Tab = 0 Then
        Me.MSHFlexGrid1.Enabled = False
        lbEditadoGen = False
        Me.txtCardesc.Text = ""
        EditMoney1.Value = 0
        EditMoney2.Value = 0
        Me.txtOrden.Text = "0"
        
        If Me.MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) = "" Then
            lsNumCur = "001"
        Else
            lsNumCur = FillNum(Trim(Str(CCur(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1) + 1))), 3, "0")
        End If
        lsCodigo = lsNumCur
        Me.txtCardesc.SetFocus
    Else
        If Not IsNumeric(Flex.TextMatrix(Flex.Rows - 1, 1)) Then
            lsNumCur = "1"
        Else
            lsNumCur = Trim(Str(CCur(Mid(Flex.TextMatrix(Flex.Rows - 1, 1), 4, 3) + 1)))
        End If
        txtcCarCod = Left(Me.cmbCargos, 3) & FillNum(lsNumCur, 3, "0")
        txtcCarDes = ""
        txtCarSue.Value = 0
        lbEditado = False
        
        Me.txtcCarDes.SetFocus
    End If
    Activa False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub EditMoney1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.EditMoney2.SetFocus
End Sub

Private Sub EditMoney2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtOrden.SetFocus
End Sub

Private Sub Flex_DblClick()
    cmdeditar_Click
End Sub

Private Sub Form_Load()
    Dim sqlN As String
    Dim rsN As New ADODB.Recordset
    
    If Not AbreConexion Then Exit Sub
    
    Activa True
    
    Refresca_C
    cmbCargos_Click
    
    Me.SSTab1.Tab = 0
End Sub

Private Sub GetData(psCarCod As String)
    Dim sqlCar As String
    Dim sqlCarDet As String
    Dim rsCar As New ADODB.Recordset
    Dim rsCarDet As New ADODB.Recordset
    
    sqlCar = "Select cCarDes,cCarSue,nOrden,nMonMaximo,cCatCod from Cargos where cCarCod = '" & Trim(psCarCod) & "' and cCarCod like '___'"
    sqlCarDet = "Select cCarCod, cCarDes, cCarSue from Cargos where cCarCod <> '" & Trim(psCarCod) & "' and cCarCod like '" & Trim(psCarCod) & "___'"
    
    rsCar.open sqlCar, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not RSVacio(rsCar) Then
        txtCardesc = rsCar!cCarDes
        txtOrden = IIf(IsNull(rsCar!nOrden), 0, rsCar!nOrden)
        EditMoney1.Value = Format(rsCar!cCarSue, "#,##0.00")
        If IsNull(rsCar!nMonMaximo) Then
            EditMoney2.Value = Format(0, "#,##0.00")
        Else
            EditMoney2.Value = Format(rsCar!nMonMaximo, "#,##0.00")
        End If
        
        If IsNull(rsCar!cCatCod) Then
            cmbCatCod.ListIndex = -1
        Else
            UbicaCombo cmbCatCod, rsCar!cCatCod
        End If
        
    End If
    
    rsCar.Close
    Set rsCar = Nothing
    
    rsCarDet.open sqlCarDet, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    ClearScreen
        
    If Not RSVacio(rsCarDet) Then
    
        While Not rsCarDet.EOF
            If Flex.TextMatrix(Flex.Rows - 1, 0) <> "" Then Flex.Rows = Flex.Rows + 1
            Flex.TextMatrix(Flex.Rows - 1, 0) = rsCarDet!cCarCod
            Flex.TextMatrix(Flex.Rows - 1, 1) = rsCarDet!cCarDes
            Flex.TextMatrix(Flex.Rows - 1, 2) = rsCarDet!cCarSue
            rsCarDet.MoveNext
        Wend
    End If
    
    rsCarDet.Close
    Set rsCarDet = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub MSHFlexGrid1_DblClick()
    cmdeditar_Click
End Sub

Private Sub SSTab1_DblClick()
    Refresca_C
End Sub

Private Sub SSTab1_GotFocus()
    Refresca_C
End Sub

Private Sub txtCardesc_GotFocus()
    txtCardesc.SelStart = 0
    txtCardesc.SelLength = 50
End Sub

Private Sub txtCardesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditMoney1.SetFocus
    Else
        KeyAscii = intfMayusculas(KeyAscii)
    End If
End Sub

Private Sub txtCarSue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.CmdAceptar.SetFocus
End Sub

Private Sub txtcCarDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCarSue.SetFocus
    Else
        KeyAscii = intfMayusculas(KeyAscii)
    End If
End Sub

Private Sub ClearScreen(Optional pbBan As Boolean = False)
    Flex.Rows = 1
    Flex.Rows = 2
    Flex.FixedRows = 1
    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 4300
    Flex.ColWidth(2) = 1150
    
    Flex.TextMatrix(0, 0) = "Código"
    Flex.TextMatrix(0, 1) = "Descripción"
    Flex.TextMatrix(0, 2) = "Sueldo"

    If pbBan Then
        cmbCatCod = ""
        txtCarSue.Value = 0
        txtcCarCod = ""
        txtCardesc = ""
        txtcCarDes = ""
    End If

End Sub

Private Function GetRepo() As String
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    
    Dim sqlC As String
    Dim rsC As New ADODB.Recordset
    Dim lsCadena As String
    
    Dim lnItem As Long
    Dim lnPagina As Long
    
    Dim lsCodigo As String * 10
    Dim lsNivel As String * 35
    Dim lsCargo As String * 50
    Dim lsRemuneracion As String
    
    lsNegritaOn = Chr$(27) + Chr$(71)
    lsNegritaOff = Chr$(27) + Chr$(72)

    
    sqlC = "Select cCarCod,cCarDes,cCarSue from Cargos order by cCarCod"
    rsC.open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    lnPagina = 0
    lnItem = 1
    If Not RSVacio(rsC) Then
        lsCadena = lsCadena & lsNegritaOn & CabeceraPagina("Niveles y Cargos.", lnPagina, lnItem, "NNN")
        lsCadena = lsCadena & Encabezado("Codigos;8; ;2;Nivel;10; ;20;Cargo;10; ;50;Remuneración;15; ;5;", lnItem) & lsNegritaOff
        While Not rsC.EOF
            lsCodigo = rsC!cCarCod
            lsNivel = IIf(Len(Trim(rsC!cCarCod)) = 3, lsNegritaOn & Trim(rsC!cCarDes) & lsNegritaOff, "")
            lsCargo = IIf(Len(Trim(rsC!cCarCod)) = 3, "", Trim(rsC!cCarDes))
            lsRemuneracion = IIf(Len(Trim(rsC!cCarCod)) = 3, "", Format(rsC!cCarSue, "#,##0.00"))
            
            lsCadena = lsCadena & lsCodigo & "  " & lsNivel & "  " & lsCargo & Space(20 - Len(lsRemuneracion)) & lsRemuneracion & Chr(10)
            lnItem = lnItem + 1
            
            If lnItem = 52 Then
                lsCadena = lsCadena & Chr(12)
                lsCadena = lsCadena & lsNegritaOn & CabeceraPagina("Niveles y Cargos.", lnPagina, lnItem, "NNN")
                lsCadena = lsCadena & Encabezado("Codigos;8; ;2;Nivel;10; ;20;Cargo;10; ;50;Remuneración;15; ;5;", lnItem) & lsNegritaOff
            End If
            rsC.MoveNext
        Wend
        
    End If
    GetRepo = lsCadena
End Function

Private Sub txtOrden_GotFocus()
    txtOrden.SelStart = 0
    txtOrden.SelLength = 100
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.CmdAceptar.SetFocus
    Else
        KeyAscii = intfNumEnt(KeyAscii)
    End If
End Sub

Private Sub Refresca_C()
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    
    sqlC = "Select cCarCod,cCarDes From Cargos where cCarCod like '___' order by cCarCod"
    rsC.open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not RSVacio(rsC) Then
        Me.cmbCargos.Clear
        While Not rsC.EOF
            cmbCargos.AddItem Trim(rsC!cCarCod) & " " & Trim(rsC!cCarDes)
            rsC.MoveNext
        Wend
    End If
    
    rsC.Close
    Set rsC = Nothing
    
    Set rsC = New ADODB.Recordset
    
    sqlC = "Select cCarCod as Codigo,cCarDes as Descripcion, convert(numeric(8,2),cCarSue) as Sueldo, convert(numeric(8,2),nMonMaximo) as Sue_Maximo,nOrden as Orden,cCatCod From Cargos where cCarCod like '___' order by cCarCod"
    rsC.open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Set Me.MSHFlexGrid1.DataSource = rsC

    rsC.Close
    Set rsC = Nothing

    MSHFlexGrid1.ColWidth(0) = 1
    MSHFlexGrid1.ColWidth(1) = 1
    MSHFlexGrid1.ColWidth(2) = 2800
    MSHFlexGrid1.ColWidth(3) = 1550
    MSHFlexGrid1.ColWidth(4) = 1550
    MSHFlexGrid1.ColWidth(5) = 1000
    MSHFlexGrid1.ColWidth(6) = 1
    
    MSHFlexGrid1.ColAlignment(3) = 7
    MSHFlexGrid1.ColAlignment(4) = 7
    MSHFlexGrid1.ColAlignment(5) = 7

End Sub

Private Sub Activa(psActiva As Boolean)
    CmdAceptar.Enabled = Not psActiva
    cmdCancelar.Enabled = Not psActiva
    cmdNuevo.Enabled = psActiva
    CmdEditar.Enabled = psActiva
    cmdEliminar.Enabled = psActiva
    Me.MSHFlexGrid1.Enabled = psActiva
    Flex.Enabled = psActiva
End Sub
