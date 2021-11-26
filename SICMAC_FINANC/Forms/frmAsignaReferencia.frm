VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAsignaReferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Pendientes: Actualización de Referencias"
   ClientHeight    =   6015
   ClientLeft      =   930
   ClientTop       =   1020
   ClientWidth     =   10590
   Icon            =   "frmAsignaReferencia.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgePend 
      Caption         =   "&Pendientes de  Agencia"
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
      Left            =   7590
      TabIndex        =   18
      Top             =   5580
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CheckBox chkAge 
      Caption         =   "Operaciones de Agencias"
      Height          =   225
      Left            =   150
      TabIndex        =   13
      Top             =   5610
      Width           =   2145
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fechas :"
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
      Height          =   645
      Left            =   135
      TabIndex        =   10
      Top             =   180
      Width           =   3810
      Begin MSMask.MaskEdBox TxtFecIni 
         Height          =   315
         Left            =   690
         TabIndex        =   0
         Top             =   225
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecFin 
         Height          =   315
         Left            =   2535
         TabIndex        =   1
         Top             =   225
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio :"
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Final :"
         Height          =   210
         Left            =   2025
         TabIndex        =   11
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.Frame Frame3 
      Height          =   660
      Left            =   3990
      TabIndex        =   8
      Top             =   180
      Width           =   4005
      Begin VB.TextBox TxtCuenta 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   930
         TabIndex        =   2
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton CmdIniciar 
         Caption         =   "&Aplicar"
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
         Left            =   2580
         TabIndex        =   3
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta :"
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
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   255
         Width           =   750
      End
   End
   Begin MSAdodcLib.Adodc ADOCta 
      Height          =   330
      Left            =   10710
      Top             =   1350
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Referencias :"
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
      Height          =   1695
      Left            =   135
      TabIndex        =   7
      Top             =   3750
      Width           =   10350
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFRef 
         Height          =   1335
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   2355
         _Version        =   393216
         FixedCols       =   0
         TextStyleFixed  =   3
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Asientos :"
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
      Height          =   2910
      Left            =   135
      TabIndex        =   6
      Top             =   825
      Width           =   10350
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFAsientos 
         Height          =   2595
         Left            =   120
         TabIndex        =   4
         Top             =   195
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin MSAdodcLib.Adodc ADORef 
      Height          =   330
      Left            =   10740
      Top             =   1785
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame fraAge 
      Height          =   525
      Left            =   2460
      TabIndex        =   14
      Top             =   5430
      Visible         =   0   'False
      Width           =   5025
      Begin Sicmact.TxtBuscar txtAgeCod 
         Height          =   315
         Left            =   930
         TabIndex        =   15
         Top             =   150
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
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
         sTitulo         =   ""
      End
      Begin VB.Label Label4 
         Caption         =   "Agencias"
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   180
         Width           =   765
      End
      Begin VB.Label lblAgeDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1650
         TabIndex        =   16
         Top             =   150
         Width           =   3285
      End
   End
   Begin VB.Menu mnurefref 
      Caption         =   "Referencias"
      Visible         =   0   'False
      Begin VB.Menu mnuActRef 
         Caption         =   "Acti&var Referencia"
      End
      Begin VB.Menu mnuref 
         Caption         =   "&Referenciar"
      End
      Begin VB.Menu mnucancelar 
         Caption         =   "&Cancelar"
      End
      Begin VB.Menu mnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefSaldo 
         Caption         =   "&Definir como Saldo"
      End
      Begin VB.Menu mnuRegPendiente 
         Caption         =   "Regulariza &Pendiente"
      End
   End
   Begin VB.Menu mnuref2 
      Caption         =   "Refere2"
      Visible         =   0   'False
      Begin VB.Menu mnueliref 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "frmAsignaReferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nActRef() As Integer
Dim oCon    As DConecta
Dim I       As Integer
Dim j       As Integer

Dim sSql As String
Dim rsDat   As ADODB.Recordset

Private Function ValidaDatos() As Boolean
Dim CadTemp As String
'valida Fechas
   ValidaDatos = False
    CadTemp = ValidaFecha(TxtFecIni.Text)
    If Len(Trim(CadTemp)) > 0 Then
        MsgBox CadTemp, vbInformation, "Aviso"
        TxtFecIni.SetFocus
        Exit Function
    End If
    CadTemp = ValidaFecha(TxtFecFin.Text)
    If Len(Trim(CadTemp)) > 0 Then
        MsgBox CadTemp, vbInformation, "Aviso"
        TxtFecFin.SetFocus
        Exit Function
    End If
    If Len(Trim(txtCuenta.Text)) <= 3 Then
        MsgBox "Falta Ingresar la Cuenta Contable", vbInformation, "Aviso"
        txtCuenta.SetFocus
        Exit Function
    End If
    ValidaDatos = True
End Function

Private Sub CargaDatosRef(ByVal lnMovNro As Long, ByVal psAgeCod As String, lsMovNro As String)
Dim rs As ADODB.Recordset
    MSFRef.Rows = 2
    MSFRef.Clear
    If lnMovNro = 0 Then
        Exit Sub
    End If
'    If psAgeCod <> "" Then
'        Dim oNeg As New NNegOpePendientes
'        If gbBitCentral Then
'            Set rs = oNeg.CargaOpeRegulaVentanillaPendCentral(lnMovNro, psAgeCod, Mid(txtCuenta, 3, 1))
'        Else
'            Set rs = oNeg.CargaOpeRegulaVentanillaPend(lsMovNro, psAgeCod, txtCuenta)
'        End If
'    Else
        oCon.AbreConexion
        sSql = "SELECT M.cmovNro, M.cMovDesc, SUM(ISNULL(me.nMovMeImporte,mc.nMovImporte)), m.nMovNro " _
             & "FROM Movref MR JOIN Mov M On MR.nMovNro = M.nMovNro " _
             & "               JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
             & "          LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
             & "WHERE nMovNroRef = '" & lnMovNro & "' and mc.cCtaContCod = '" & txtCuenta & "' " _
             & "GROUP BY M.cmovNro, M.cMovDesc, m.nMovNro "
        Set rs = oCon.CargaRecordSet(sSql)
'    End If
    If Not rs.EOF Then
       Set MSFRef.Recordset = rs
    End If
    'Ordena Flex de Referencia
    MSFRef.ColWidth(0) = 2500
    MSFRef.TextMatrix(0, 0) = "MOVIMIENTO"
    MSFRef.ColWidth(1) = 5100
    MSFRef.TextMatrix(0, 1) = "DESCRIPCION"
    MSFRef.ColWidth(3) = 0
    MSFRef.ColWidth(7) = 0
    RSClose rs
    oCon.CierraConexion
End Sub

Private Sub chkAge_Click()
If chkAge.value = vbChecked Then
    fraAge.Visible = True
    txtAgeCod.SetFocus
    Me.cmdAgePend.Visible = True
Else
    fraAge.Visible = False
    txtAgeCod = ""
    Me.cmdAgePend.Visible = False
End If
End Sub

Private Sub cmdAgePend_Click()
Dim rsN As ADODB.Recordset
If Not ValidaDatos Then
    Exit Sub
End If
Screen.MousePointer = 11
If Me.chkAge.value = vbChecked Then
   If Me.txtAgeCod = "" Then
      MsgBox "No se indico Agencia donde se realizaron las Operaciones", vbInformation, "¡Aviso!"
   Else
     Dim oAna As New NAnalisisCtas
      Set rsN = oAna.GetOpePendientesNegocioRef(gbBitCentral, Me.TxtFecIni, Me.TxtFecFin, Mid(txtCuenta, 3, 1), Me.txtAgeCod, Left(Me.txtCuenta, 2) & "_" & Mid(txtCuenta, 4))
      RecordSetAdiciona rsDat, rsN
      Set oAna = Nothing
   End If
End If
If rsDat.RecordCount > 0 Then
   Set MSFAsientos.Recordset = rsDat
End If
FormatoFlex
Screen.MousePointer = 0

End Sub

Private Sub CmdIniciar_Click()
    If ValidaDatos Then
        Call CargaDatos
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    ReDim nActRef(0)
    Set oCon = New DConecta
    Set rsDat = New ADODB.Recordset
    Dim oAge As New DActualizaDatosArea
    txtAgeCod.rs = oAge.GetAgencias(, True)
    Set oAge = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCon = Nothing
End Sub

Private Sub mnuActRef_Click()
    For I = 0 To 3
        MSFAsientos.Col = I
        MSFAsientos.CellBackColor = vbYellow
    Next I
    ReDim Preserve nActRef(UBound(nActRef) + 1)
    nActRef(UBound(nActRef)) = MSFAsientos.Row
End Sub

Private Sub mnucancelar_Click()
For j = 1 To UBound(nActRef)
    MSFAsientos.Row = nActRef(j)
    For I = 0 To 3
        MSFAsientos.Col = I
        MSFAsientos.CellBackColor = vbWhite
    Next I
Next
ReDim nActRef(0)
End Sub

Private Sub mnuDefSaldo_Click()
Dim sSql As String
Dim nRow As Integer
Dim prs  As ADODB.Recordset
Dim oMov As DMov
Set oMov = New DMov
nRow = MSFAsientos.Row
sSql = "SELECT * FROM MovPendientesRend " _
     & "WHERE nMovNro = " & MSFAsientos.TextMatrix(nRow, 5) & " and cCtaContcod = '" & txtCuenta & "' "
Set prs = oMov.CargaRecordSet(sSql)
If Not prs.EOF Then
    oMov.ActualizaMovPendientesRend MSFAsientos.TextMatrix(nRow, 5), txtCuenta, prs!nSaldo - Abs(nVal(MSFAsientos.TextMatrix(nRow, 3)))
Else
    oMov.InsertaMovPendientesRend MSFAsientos.TextMatrix(nRow, 5), txtCuenta, Abs(MSFAsientos.TextMatrix(nRow, 3))
End If
Set oMov = Nothing
MSFAsientos.TextMatrix(nRow, 4) = Abs(MSFAsientos.TextMatrix(nRow, 3))
End Sub

Private Sub mnueliref_Click()
Dim sSql As String
Dim nPos1 As Integer
Dim nPos2 As Integer
Dim oMov  As DMov
    If nVal(MSFAsientos.TextMatrix(MSFAsientos.Row, 5)) <> 0 Then
        If MsgBox("Desea Eliminar el registro ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Set oMov = New DMov
            nPos1 = MSFAsientos.Row
            nPos2 = MSFRef.Row
            oMov.EliminaMovRef Trim(MSFRef.TextMatrix(nPos2, 3)), Trim(MSFAsientos.TextMatrix(nPos1, 5))
            oMov.ActualizaMovPendientesRend MSFAsientos.TextMatrix(nPos1, 5), txtCuenta, nVal(MSFRef.TextMatrix(nPos2, 2))
            Set oMov = Nothing
            Call CargaDatosRef(MSFAsientos.TextMatrix(nPos1, 5), MSFAsientos.TextMatrix(nPos1, 7), MSFAsientos.TextMatrix(nPos1, 6))
            MSFAsientos.TextMatrix(nPos1, 4) = nVal(MSFAsientos.TextMatrix(nPos1, 4)) + nVal(MSFRef.TextMatrix(nPos2, 2))
        End If
    End If
End Sub

Private Sub mnuref_Click()
Dim sSql As String
Dim nActRef2 As Integer
Dim R As New ADODB.Recordset
Dim bEnc As Boolean
Dim oMov As DMov
Dim oNeg As New NNegOpePendientes
    If MSFAsientos.CellBackColor = vbYellow Then
       MsgBox "No puede Referenciar el Mismo Movimiento", vbInformation, "Aviso"
    End If
    If MsgBox("Desea Referenciar este Movimiento con el Movimiento Selecionado Anteriormente ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        For I = 0 To 3
            MSFAsientos.Col = I
            MSFAsientos.CellBackColor = vbGreen
        Next I
        
        'Actualiza MovRef
        nActRef2 = MSFAsientos.Row
        Set oMov = New DMov
        oMov.BeginTrans
        For j = 1 To UBound(nActRef)
            sSql = "Select * from Movref Where nMovNro = " & Trim(MSFAsientos.TextMatrix(nActRef2, 5)) & " And nMovNroRef = " & Trim(MSFAsientos.TextMatrix(nActRef(j), 5))
            Set R = oMov.CargaRecordSet(sSql)
            bEnc = Not R.EOF
            RSClose R
            If bEnc Then
                MsgBox "Referencia Ya Existe", vbInformation, "Aviso"
            Else
                If CDate(GetFechaMov(MSFAsientos.TextMatrix(nActRef2, 0) & MSFAsientos.TextMatrix(nActRef2, 1), True)) < CDate(GetFechaMov(MSFAsientos.TextMatrix(nActRef(j), 0) & MSFAsientos.TextMatrix(nActRef(j), 1), True)) Then
                   If MsgBox("El Movimiento de Referencia No puede ser anterior al movimiento Referenciado. ¿Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                      GoTo NoAsignaReferencia
                   End If
               End If
               If Len(MSFAsientos.TextMatrix(nActRef(j), 1)) < 16 Then
                  oNeg.ActualizaNegocioReferencia MSFAsientos.TextMatrix(nActRef(j), 5), gsCodCMAC & txtAgeCod, MSFAsientos.TextMatrix(MSFAsientos.Row, 6)
                  MSFAsientos.TextMatrix(nActRef(j), 4) = 0
               Else
                  oMov.ActualizaMovPendientesRend MSFAsientos.TextMatrix(nActRef(j), 5), txtCuenta, nVal(MSFAsientos.TextMatrix(nActRef2, 3))
                  MSFAsientos.TextMatrix(nActRef(j), 4) = nVal(MSFAsientos.TextMatrix(nActRef(j), 4)) - nVal(MSFAsientos.TextMatrix(nActRef2, 3))
               End If
               oMov.InsertaMovRef MSFAsientos.TextMatrix(nActRef2, 5), MSFAsientos.TextMatrix(nActRef(j), 5), IIf(Me.chkAge.Visible, txtAgeCod, "")
NoAsignaReferencia:
            End If
            MSFAsientos.Row = nActRef2
            For I = 0 To 3
                MSFAsientos.Col = I
                MSFAsientos.CellBackColor = vbWhite
            Next I
            MSFAsientos.Row = nActRef(j)
            For I = 0 To 3
                MSFAsientos.Col = I
                MSFAsientos.CellBackColor = vbWhite
            Next I
            MSFAsientos.Row = nActRef2
        Next j
        oMov.CommitTrans
        'Refrescar datos de referencia
        Call CargaDatosRef(MSFAsientos.TextMatrix(nActRef(UBound(nActRef)), 5), MSFAsientos.TextMatrix(nActRef(UBound(nActRef)), 7), MSFAsientos.TextMatrix(nActRef(UBound(nActRef)), 6))
        
        ReDim nActRef(LBound(nActRef))
    End If
Set oNeg = Nothing
Exit Sub
AsignaReferenciaErr:
    oMov.RollbackTrans
End Sub

Private Sub mnuRegPendiente_Click()
Dim sSql As String
Dim nRow As Integer
Dim oMov As DMov
nRow = MSFAsientos.Row
Set oMov = New DMov
oMov.ActualizaMovPendientesRend MSFAsientos.TextMatrix(nRow, 5), txtCuenta, nVal(MSFAsientos.TextMatrix(nRow, 4))
Set oMov = Nothing
MSFAsientos.TextMatrix(nRow, 4) = "0"
End Sub

Private Sub MSFAsientos_Click()
Dim nPos As Integer
    nPos = MSFAsientos.Row
    CargaDatosRef MSFAsientos.TextMatrix(nPos, 5), MSFAsientos.TextMatrix(nPos, 7), MSFAsientos.TextMatrix(nPos, 6)
End Sub

Private Sub MSFAsientos_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = 2 Then   '   Copiar  [Ctrl+C]
   Flex_PresionaKey MSFAsientos, KeyCode, Shift
End If
End Sub

Private Sub MSFAsientos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
    If Button = 2 Then
        If UBound(nActRef) = 0 Then
            mnuref.Enabled = False
            mnucancelar = False
            PopupMenu mnurefref
        Else
            mnuref.Enabled = True
            mnucancelar.Enabled = True
            PopupMenu mnurefref
        End If
    End If
End Sub

Private Sub MSFRef_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = 2 Then   '   Copiar  [Ctrl+C]
   Flex_PresionaKey MSFRef, KeyCode, Shift
End If
End Sub

Private Sub MSFRef_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuref2
    End If
End Sub

Private Sub txtAgeCod_EmiteDatos()
Me.lblAgeDesc = txtAgeCod.psDescripcion
If lblAgeDesc <> "" Then
    cmdAgePend.SetFocus
End If
End Sub

Private Sub txtCuenta_GotFocus()
    fEnfoque txtCuenta
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii, False)
    If KeyAscii = 13 Then
        CmdIniciar.SetFocus
    End If
End Sub

Private Sub TxtFecFin_GotFocus()
    fEnfoque TxtFecFin
End Sub

Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fEnfoque txtCuenta
        txtCuenta.SetFocus
    End If
End Sub

Private Sub TxtFecIni_GotFocus()
    fEnfoque TxtFecIni
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fEnfoque TxtFecFin
        TxtFecFin.SetFocus
    End If
End Sub

Private Sub CargaDatos()
Dim sSql As String
Dim rsN  As ADODB.Recordset

   oCon.AbreConexion
   Screen.MousePointer = 11
   RSClose rsDat
   Set rsDat = New ADODB.Recordset
   sSql = "SELECT LEFT(m.cMovNro,8) cMovDia, SUBSTRING(m.cMovNro,9,24) cMovHora, m.cMovDesc, " & IIf(Mid(txtCuenta, 3, 1) = "2", "me.nMovMEImporte", "mc.nMovImporte") & " nMovImporte, ISNULL(mRend.nSaldo,0) as nSaldo, m.nMovNro, m.cMovNro, '' cAgeCod " _
        & "FROM mov m JOIN movcta mc on m.nmovnro = mc.nmovnro " & IIf(Mid(Me.txtCuenta, 3, 1) = "2", " JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
        & "      LEFT JOIN (SELECT r.nMovNro, r.nMovNroRef, mc1.cCtaContCod FROM Mov m JOIN MovRef r ON r.nMovNroRef = m.nMovNro JOIN MovCta mc1 ON mc1.nMovNro = m.nMovNro ) mr  ON mr.nMovNro = m.nMovNro and mr.cCtaContCod = mc.cCtaContCod " _
        & "      LEFT JOIN MovPendientesRend mRend ON mRend.nMovNro = m.nMovNro and mRend.cCtaContCod = '" & txtCuenta & "' " _
        & "WHERE LEFT(m.cMovNro,8) BETWEEN '" & Format(Me.TxtFecIni, gsFormatoMovFecha) & "' and '" & Format(TxtFecFin, gsFormatoMovFecha) & "' and m.nMovEstado in (" & gMovEstContabMovContable & "," & gMovEstContabPendiente & ") and m.nMovFlag in (" & gMovFlagVigente & "," & gMovFlagDeExtorno & "," & gMovFlagExtornado & ") and mc.cCtaContCod =  '" & txtCuenta & "'" _
        & "      and (mr.nMovNro IS NULL or ISNULL(mRend.nSaldo,0) <> 0) and not m.cMovNro like '%XXX%' " _
        & "ORDER BY mc.cCtaContCod, m.cMovNro "
    Set rsN = oCon.CargaRecordSet(sSql)
    MSFAsientos.Rows = 2
    MSFAsientos.Clear
    If Not rsN.EOF Then
      RecordSetAdiciona rsDat, rsN
    End If
    If rsN.RecordCount > 0 Then
        If rsDat.RecordCount > 0 Then
           Set MSFAsientos.Recordset = rsDat
        End If
        FormatoFlex
    End If
   Screen.MousePointer = 0
End Sub

Private Sub FormatoFlex()

   MSFAsientos.TextMatrix(0, 0) = "FECHA"
   MSFAsientos.ColWidth(0) = 800
   MSFAsientos.ColAlignment(0) = 6
   MSFAsientos.TextMatrix(0, 1) = "MOVIMIENTO"
   MSFAsientos.ColWidth(1) = 1600
   MSFAsientos.ColAlignment(1) = 1
   MSFAsientos.TextMatrix(0, 2) = "DESCRIPCION"
   MSFAsientos.ColWidth(2) = 5200
   MSFAsientos.TextMatrix(0, 3) = "IMPORTE"
   MSFAsientos.ColWidth(3) = 1000
   MSFAsientos.ColAlignment(3) = 6
   MSFAsientos.TextMatrix(0, 4) = "SALDO"
   MSFAsientos.ColWidth(4) = 1000
   MSFAsientos.ColAlignment(4) = 6
   MSFAsientos.TextMatrix(0, 5) = "nMovNro"
   MSFAsientos.ColWidth(5) = 0
   MSFAsientos.TextMatrix(0, 6) = "cMovNro"
   MSFAsientos.ColWidth(6) = 0

End Sub

