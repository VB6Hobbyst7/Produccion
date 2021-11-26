VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmContabManVarNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Cuenta Contable"
   ClientHeight    =   6855
   ClientLeft      =   2670
   ClientTop       =   1215
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContabManVarNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc ADOCtas 
      Height          =   330
      Left            =   6300
      Top             =   555
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar SBEstado 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   6525
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   3330
      Left            =   45
      TabIndex        =   17
      Top             =   3180
      Width           =   6030
      Begin VB.CommandButton CmdGrabaCta 
         Caption         =   "&Grabar Variables Cuenta"
         Height          =   330
         Left            =   225
         TabIndex        =   8
         Top             =   2955
         Width           =   2745
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   4575
         TabIndex        =   10
         Top             =   2955
         Width           =   1380
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   330
         Left            =   3090
         TabIndex        =   9
         Top             =   2955
         Width           =   1380
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   330
         Left            =   3705
         TabIndex        =   6
         Top             =   195
         Width           =   1635
      End
      Begin VB.TextBox TxtBuscar 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid DGrid 
         Height          =   2235
         Left            =   150
         TabIndex        =   7
         Top             =   600
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   3942
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cCta"
            Caption         =   "Cuentas"
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
            DataField       =   "cDescrip"
            Caption         =   "Descripcion"
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
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4500.284
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3195
      Left            =   45
      TabIndex        =   11
      Top             =   -15
      Width           =   6030
      Begin VB.CommandButton cmdclase 
         Caption         =   "&Aplicar"
         Height          =   330
         Left            =   2025
         TabIndex        =   1
         Top             =   210
         Width           =   1620
      End
      Begin VB.CommandButton cmdgenerar 
         Caption         =   "&Generar"
         Height          =   315
         Left            =   1710
         TabIndex        =   4
         Top             =   2805
         Width           =   2685
      End
      Begin VB.HScrollBar HSCuenta 
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   2460
         Visible         =   0   'False
         Width           =   5760
      End
      Begin VB.HScrollBar HSVar 
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1365
         Visible         =   0   'False
         Width           =   5760
      End
      Begin VB.PictureBox PicVar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   5685
         TabIndex        =   2
         Top             =   900
         Width           =   5745
         Begin VB.Label LblVariable 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "M"
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000001&
            Height          =   270
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   60
            Width           =   465
         End
      End
      Begin VB.PictureBox PicValVar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         ScaleHeight     =   420
         ScaleWidth      =   5700
         TabIndex        =   3
         Top             =   1965
         Width           =   5760
         Begin VB.Label LblVarCta 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "M"
            DragMode        =   1  'Automatic
            ForeColor       =   &H80000001&
            Height          =   270
            Index           =   0
            Left            =   105
            TabIndex        =   16
            Top             =   75
            Width           =   465
         End
      End
      Begin VB.TextBox txtClase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   885
         MaxLength       =   8
         TabIndex        =   0
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "Variables de Cuenta"
         Height          =   285
         Left            =   210
         TabIndex        =   19
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Variables"
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   615
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "Clase :"
         Height          =   285
         Left            =   165
         TabIndex        =   13
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmContabManVarNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumCtrlCta As Integer
Dim PX, PY As Long
Dim PDx As Long
Dim sCTAS() As String * 100
Dim Estado() As Boolean
Dim NumCtas As Long
Const NumVarPic = 11
Const OrigPX = 90
Const OrigPY = 60
Const OrigPDx = 500
Dim lsCtaClase As String
Dim oCon As DConecta

Public Function Inicio(psCtaClase As String)
lsCtaClase = psCtaClase
Me.Show 1
End Function

Private Sub CargaVarCtaExistente(ByVal pClase As String)
Dim sSql As String
Dim R As New ADODB.Recordset
Dim R2 As New ADODB.Recordset
Dim nPos As Integer
Dim I As Integer
Dim j As Integer

    nPos = 0
    sSql = "Select CG.cClase, VC.cAbrev, VC.nCodigo From CtasGen CG Inner join VarCtasCont VC On CG.nCodigo = VC.nCodigo " _
           & " Where cClase = '" & pClase & "' Order By nOrden"
    Set R = oCon.CargaRecordSet(sSql)
        PX = 90
        PY = 60
        PDx = 500
        Do While Not R.EOF
            If nPos > 0 Then
                Call Load(LblVarCta(nPos))
            End If
            LblVarCta(nPos).ForeColor = &H80000001
            LblVarCta(nPos).Alignment = 2
            LblVarCta(nPos).Caption = R!cAbrev & Space(80) & Trim(Str(R!nCodigo))
            LblVarCta(nPos).Top = PY
            LblVarCta(nPos).Left = PX
            LblVarCta(nPos).Visible = True
            PX = PX + PDx
            nPos = nPos + 1
            R.MoveNext
        Loop
    R.Close
    NumCtrlCta = IIf(LblVarCta.Count = 1, -1, LblVarCta.Count - 1)
    
    'Si se ha bajado ya la variable cambiar de color y deshabilitar
    If LblVarCta.Count > 1 Then
        For I = 0 To LblVariable.Count - 1
            For j = 0 To LblVarCta.Count - 1
                If Trim(Left(LblVariable(I).Caption, 5)) = Trim(Left(LblVarCta(j).Caption, 5)) Then
                    LblVariable(I).BackColor = &HFFC0C0
                    LblVariable(I).Enabled = False
                End If
            Next j
        Next I
    End If
End Sub

Private Function ValidaDatos() As Boolean
    If Len(Trim(txtClase.Text)) = 0 Then
        MsgBox "Falta Ingresar la Clase", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    If NumCtrlCta <= 0 Then
        MsgBox "Falta Mas Variable a la Clase", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    ValidaDatos = True
End Function

Private Sub GeneraCuentas(ByVal sClase As String)
Dim I, j As Integer
Dim sCadSel As String
Dim sCadFrom As String
Dim sCadWhere As String
Dim sCadRes As String
Dim sCadTotal As String
Dim ContVar As Integer
Dim R4 As New ADODB.Recordset
Dim sSql As String
Dim sCtaArmada As String
Dim K As Integer
Dim sCodRestric As String

    SBEstado.Panels(1).Text = "Generando Cuentas.."
    If Len(txtClase) = 2 Then
        ContVar = 2
    Else
        ContVar = 1
    End If
    oCon.Ejecutar "Delete TempCtasGen"
    
    sCodRestric = ""
    For j = 0 To LblVarCta.Count - 1
        sCadSel = "INSERT INTO TempCtasGen Select '" & sClase & "'"
        sCadFrom = " FROM "
        sCadWhere = " WHERE "
        sCadRes = " And "
        sCtaArmada = "'" & sClase & "'"
        If ContVar <= LblVarCta.Count Then
            For I = 0 To ContVar - 1
                sCadSel = sCadSel & " + RTRIM(VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ".cValor)"
                sCtaArmada = sCtaArmada & " + RTRIM(VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ".cValor)"
                sCadFrom = sCadFrom & "ValvarCtas " & "VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ", "
                sCadWhere = sCadWhere & "VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ".nCodigo = " & Trim(Right(LblVarCta(I).Caption, 3)) & " and "
                sCodRestric = sCodRestric & "'" & Trim(Right(LblVarCta(I).Caption, 3)) & "',"
                'Arma Cadena de Restriccion
                sSql = "Select * from RestVarCta Where nCodRes = " & Trim(Right(LblVarCta(I).Caption, 3)) & " and nCodigo IN (" & Left(sCodRestric, Len(sCodRestric) - 1) & ")"
                Set R4 = oCon.CargaRecordSet(sSql)
                If Not R4.BOF And Not R4.EOF Then
                    sCadRes = sCadRes & " NOT VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ".cCodValor IN (SELECT cCodValRes FROM restvarcta WHERE nCodRes = VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ".nCodigo and nCodigo = VV" & R4!nCodigo & ".nCodigo  and ccodvalor = VV" & R4!nCodigo & ".ccodvalor and ccodvalres = VV" & Trim(Right(LblVarCta(I).Caption, 3)) & ".ccodvalor) and"
                End If
                R4.Close

'                For k = 0 To ContVar - 1
'                    If k <> ContVar - 1 Then
'                        SSQL = SSQL & Trim(Right(LblVarCta(k).Caption, 3)) & ","
'                    Else
'                        SSQL = SSQL & Trim(Right(LblVarCta(k).Caption, 3)) & ")"
'                    End If
'                Next k
'                R4.Open SSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'                If Not R4.BOF And Not R4.EOF Then
'                    Do While Not R4.EOF
'                        sCadRes = sCadRes & " (VV" & Trim(Str(R4!nCodigo)) & ".cCodValor = '" & Trim(R4!cCodValor) & "' And VV" & Trim(Str(R4!nCodRes)) & ".cCodValor = '" & Trim(R4!cCodValRes) & "') OR "
'                        R4.MoveNext
'                    Loop
'                    sCadRes = Mid(sCadRes, 1, Len(sCadRes) - 3) & ") And ("
'                End If
'               R4.Close
            Next I
            sCadSel = sCadSel & " as cCtaContCod ,VV" & Trim(Right(LblVarCta(I - 1).Caption, 3)) & ".cDescrip, NEWID()"
            sCadFrom = Mid(sCadFrom, 1, Len(sCadFrom) - 2)
            sCadWhere = Mid(sCadWhere, 1, Len(sCadWhere) - 4)
            sCadTotal = sCadSel + sCadFrom + sCadWhere
            
            'Añade las Restricciones
            If Len(sCadRes) > 5 Then
                'sCadRes = "Select " & sCtaArmada & " cCtaContCod " & Space(1) & sCadFrom & sCadWhere & Mid(sCadRes, 1, Len(sCadRes) - 7) + ")"
                'sCadTotal = sCadTotal & " ) as Cta Left Join  (" & sCadRes & ") Res ON res.cCtaContCod = cta.cCtaContCod " & IIf(J = 0, "", "JOIN TempCtasGen TG ON Cta.cCtaContCod LIKE TG.cCta + '%'") & " WHERE res.cCtaContCod is NULL"
                sCadTotal = sCadTotal & Mid(sCadRes, 1, Len(sCadRes) - 3)
            End If
            
            SBEstado.Panels(1).Text = "Generando Cuentas.."
            oCon.Ejecutar sCadTotal
            ContVar = ContVar + 1
        End If
    Next j
    
    ADOCtas.Refresh
    SBEstado.Panels(1).Text = "Proceso ha Finalizado.."
End Sub

Private Sub cmdbuscar_Click()
Dim sSql As String
    If Len(Trim(TxtBuscar.Text)) > 0 Then
        sSql = "cCta like '" & TxtBuscar.Text & "%'"
        ADOCtas.Recordset.Find sSql, , adSearchForward, 1
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdclase_Click()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim NumCtrl As Integer
Dim Dx As Long
Dim X, Y As Integer
    
    If Len(Trim(txtClase.Text)) = 0 Then
        MsgBox "Falta Ingresar la Clase", vbInformation, "Aviso"
        txtClase.SetFocus
        Exit Sub
    End If
    
    X = 90
    Y = 60
    Dx = 500
    NumCtrl = 0
    sSql = "Select * from VarCtasCont"
    Set R = oCon.CargaRecordSet(sSql)
    Do While Not R.EOF
        If NumCtrl > 0 Then
            Call Load(LblVariable(NumCtrl))
        End If
        LblVariable(NumCtrl).Alignment = 2
        LblVariable(NumCtrl).Left = X
        LblVariable(NumCtrl).Top = Y
        LblVariable(NumCtrl).Visible = True
        LblVariable(NumCtrl).BackColor = &H80000018
        LblVariable(NumCtrl).Enabled = True
        LblVariable(NumCtrl).Caption = Trim(R!cAbrev) + Space(80) + Trim(R!nCodigo)
        NumCtrl = NumCtrl + 1
        If NumCtrl > NumVarPic Then
            HSVar.Visible = True
            HSVar.Max = NumCtrl - NumVarPic
            HSVar.value = 0
        End If
        X = X + Dx
        R.MoveNext
    Loop
    R.Close
    'Si existe Cta levantar sus variables
    Call CargaVarCtaExistente(Trim(txtClase.Text))
    cmdclase.Enabled = False
End Sub

Private Sub cmdGenerar_Click()
    If NumCtrlCta > 0 Then
        Call GeneraCuentas(Trim(txtClase.Text))
    Else
        MsgBox "Falta Ingresar mas variables", vbInformation, "Aviso"
    End If
End Sub

Private Sub CmdGrabaCta_Click()
Dim sSql As String
Dim I As Integer
    
    If MsgBox("Desea Grabar las Variables de la Cta ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If ValidaDatos Then
            sSql = "Delete CtasGen Where cClase = '" & Trim(txtClase.Text) & "'"
            oCon.Ejecutar sSql
            For I = 0 To LblVarCta.Count - 1
                sSql = "INSERT INTO CtasGen(nCodigo,cClase,dFecha,cCodUsu) VALUES ('" & Trim(Right(LblVarCta(I).Caption, 2)) _
                      & "','" & Trim(txtClase.Text) & "','" & FechaHora(gdFecSis) & "','" & gsCodUser & "')"
                oCon.Ejecutar sSql
            Next I
        End If
    End If
End Sub

Private Sub CmdGrabar_Click()
Dim sSql As String
Dim I As Integer
MsgBox "Las Cuentas generadas se grabarán en el Plan de Cuentas", vbInformation, "¡¡¡ ADVERTENCIA !!!!"
    If MsgBox("Desea Grabar las Cuentas Generadas ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        If ValidaDatos Then
            If ADOCtas.Recordset.RecordCount > 0 Then
                sSql = "INSERT INTO CtaCont (cCtaContCod, cCtaContDesc, cUltimaActualizacion) SELECT cCta, cDescrip, '" & GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge) & "' FROM TempCtasGen WHERE cCta not in (SELECT cCtaContCod FROM CtaCont)"
                oCon.Ejecutar sSql
                Unload Me
            Else
                MsgBox "No existen Datos para Grabar", vbInformation, "Aviso"
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    Set oCon = New DConecta
    oCon.AbreConexion
    ADOCtas.ConnectionString = gsConnection
    ADOCtas.CommandType = adCmdText
    ADOCtas.RecordSource = "SELECT * FROM TempCtasGen ORDER BY cCta"
    Set DGrid.DataSource = ADOCtas
    LblVarCta(0).Visible = False
    LblVariable(0).Visible = False
    txtClase = lsCtaClase
    NumCtrlCta = -1
    PX = 90
    PY = 60
    PDx = 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oCon = Nothing
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.DragIcon = LoadPicture()   ' Descarga el icono.
End Sub

Private Sub HSCuenta_Change()
Dim I As Integer
Dim TX As Integer
     TX = OrigPX
    For I = HSCuenta.value To LblVarCta.Count - 1
        LblVarCta(I).Left = TX
        LblVarCta(I).Visible = True
        TX = TX + OrigPDx
    Next I
    I = HSCuenta.value - 1
    Do While I >= 0
        LblVarCta(I).Visible = False
        I = I - 1
    Loop
End Sub

Private Sub HSVar_Change()
Dim I As Integer
Dim TX As Integer
    
    TX = OrigPX
    For I = HSVar.value To LblVariable.Count - 1
        LblVariable(I).Left = TX
        LblVariable(I).Visible = True
        TX = TX + OrigPDx
    Next I
    I = HSVar.value - 1
    Do While I >= 0
        LblVariable(I).Visible = False
        I = I - 1
    Loop
End Sub


Private Sub PicValVar_DragDrop(Source As Control, X As Single, Y As Single)
    Source.DragIcon = LoadPicture()   ' Descarga el icono.
    If Source.Name = LblVariable(0).Name Then
        NumCtrlCta = NumCtrlCta + 1
        If NumCtrlCta > 0 Then
            Call Load(LblVarCta(NumCtrlCta))
        End If
        LblVarCta(NumCtrlCta).ForeColor = &H80000001
        LblVarCta(NumCtrlCta).Alignment = 2
        LblVarCta(NumCtrlCta).Caption = Source.Caption
        LblVarCta(NumCtrlCta).Top = PY
        LblVarCta(NumCtrlCta).Left = PX
        LblVarCta(NumCtrlCta).Visible = True
        Source.Enabled = False
        Source.BackColor = &HFFC0C0
        PX = PX + PDx
        If NumCtrlCta > NumVarPic - 1 Then
            HSCuenta.Visible = True
            HSCuenta.Max = NumCtrlCta - (NumVarPic - 1)
            HSCuenta.value = 0
        End If
    End If
End Sub

Private Sub PicVar_DragDrop(Source As Control, X As Single, Y As Single)
Dim I As Integer
Dim Valor As String
Dim TX As Integer
    If Trim(Source.Name) = LblVarCta(0).Name Then
        Source.DragIcon = LoadPicture()   ' Descarga el icono.
        If NumCtrlCta = 0 Then
            LblVarCta(NumCtrlCta).Visible = False
            Valor = Trim(Right(Source.Caption, 5))
            NumCtrlCta = NumCtrlCta - 1
        Else
            If NumCtrlCta > 0 Then
                NumCtrlCta = NumCtrlCta - 1
                If NumCtrlCta <= NumVarPic - 1 Then 'Deshabilita Scroll
                    HSCuenta.Visible = False
                    TX = OrigPX
                    For I = 0 To LblVarCta.Count - 1
                        LblVarCta(I).Left = TX
                        LblVarCta(I).Visible = True
                        TX = TX + OrigPDx
                    Next I
                    PX = LblVarCta(LblVarCta.Count - 1).Left + OrigPDx
                End If
            End If
            Valor = Trim(Right(Source.Caption, 5))
            For I = Source.Index To NumCtrlCta
                LblVarCta(I) = LblVarCta(I + 1)
            Next I
            Call Unload(LblVarCta(NumCtrlCta + 1))
        End If
        For I = 0 To LblVariable.Count - 1
            If Trim(Right(LblVariable(I).Caption, 5)) = Valor Then
                LblVariable(I).Enabled = True
                LblVariable(I).BackColor = &H80000018
                Exit For
            End If
        Next I
        PX = PX - PDx
    End If
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub txtClase_Change()
Dim I As Integer
    If LblVarCta.Count >= 2 Then
        For I = 1 To LblVarCta.Count - 1
            Call Unload(LblVarCta(I))
        Next I
    End If
    If LblVariable.Count >= 2 Then
        For I = 1 To LblVariable.Count - 1
            Call Unload(LblVariable(I))
        Next I
    End If
    LblVarCta(0).Visible = False
    LblVariable(0).Visible = False
    PX = 90
    PY = 60
    PDx = 500
    NumCtas = 0
    cmdclase.Enabled = True
End Sub

Private Sub txtClase_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmdclase.SetFocus
    End If
End Sub
