VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHPagoAMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provisión de Pago Asistencia Medica Privada"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmRHPagoAMP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReporte 
      Caption         =   "&Reporte de Provisión"
      Height          =   360
      Left            =   3030
      TabIndex        =   27
      Top             =   2790
      Width           =   2190
   End
   Begin VB.CheckBox chkRepo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Con Reportes"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   6135
      Value           =   1  'Checked
      Width           =   2145
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   75
      Left            =   2865
      TabIndex        =   8
      Top             =   6270
      Visible         =   0   'False
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   132
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"frmRHPagoAMP.frx":030A
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
   Begin VB.CommandButton cmdSalir 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4830
      TabIndex        =   5
      Top             =   6150
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   15
      TabIndex        =   6
      Top             =   0
      Width           =   6015
      Begin MSMask.MaskEdBox mskFam 
         Height          =   300
         Left            =   4590
         TabIndex        =   15
         Top             =   2265
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPer 
         Height          =   300
         Left            =   4590
         TabIndex        =   14
         Top             =   1860
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPolizaFam 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   2265
         Width           =   1290
      End
      Begin VB.TextBox txtPolizaPer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1665
         TabIndex        =   1
         Top             =   1860
         Width           =   1290
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   285
         Left            =   1305
         TabIndex        =   0
         Top             =   1440
         Width           =   345
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "&Provision de Pago"
         Height          =   360
         Left            =   780
         TabIndex        =   3
         Top             =   2790
         Width           =   2190
      End
      Begin VB.TextBox txtPersona 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "F.Doc Familiar :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3090
         TabIndex        =   17
         Top             =   2288
         Width           =   1455
      End
      Begin VB.Label lblFecDocPer 
         Caption         =   "F.Doc Personal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3090
         TabIndex        =   16
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label lblPolizaFam 
         Caption         =   "Familiares :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   2288
         Width           =   1140
      End
      Begin VB.Label lblPolizaPersonal 
         Caption         =   "Personal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   1883
         Width           =   1140
      End
      Begin VB.Label lblTitPersona 
         Caption         =   "Persona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   1230
         Width           =   1140
      End
      Begin VB.Label lblPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1665
         TabIndex        =   10
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Line Line1 
         BorderStyle     =   0  'Transparent
         X1              =   1140
         X2              =   1200
         Y1              =   2325
         Y2              =   1305
      End
      Begin VB.Label lblPago 
         Caption         =   $"frmRHPagoAMP.frx":038A
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
         Height          =   990
         Left            =   180
         TabIndex        =   7
         Top             =   165
         Width           =   5610
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   15
      TabIndex        =   18
      Top             =   3150
      Width           =   6015
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   4905
         TabIndex        =   26
         Top             =   2055
         Width           =   975
      End
      Begin VB.ListBox lstEmpleados 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   120
         TabIndex        =   25
         Top             =   1410
         Width           =   4725
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   330
         Left            =   4905
         TabIndex        =   24
         Top             =   1590
         Width           =   975
      End
      Begin VB.TextBox txtEmp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1065
         Width           =   1170
      End
      Begin VB.CommandButton cmdBuscar1 
         Caption         =   "..."
         Height          =   285
         Left            =   1305
         TabIndex        =   19
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label lblAyuda 
         Caption         =   $"frmRHPagoAMP.frx":04AE
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
         Height          =   645
         Left            =   180
         TabIndex        =   23
         Top             =   165
         Width           =   5610
      End
      Begin VB.Line Line2 
         BorderStyle     =   0  'Transparent
         X1              =   1140
         X2              =   1200
         Y1              =   2325
         Y2              =   1305
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1665
         TabIndex        =   22
         Top             =   1065
         Width           =   4215
      End
      Begin VB.Label lblempleado 
         Caption         =   "Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   135
         TabIndex        =   21
         Top             =   840
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmRHPagoAMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    If Me.txtEmp.Text = "" Then
        MsgBox "Debe ingresar un empleado valido.", vbInformation, "Aviso"
        Me.cmdBuscar1.SetFocus
        Exit Sub
    End If
    Me.lstEmpleados.AddItem Me.txtEmp.Text & " - " & Me.lblNombre.Caption
    Me.txtEmp.Text = ""
    Me.lblNombre.Caption = ""
End Sub

Private Sub cmdBuscar_Click()
    frmBuscaCli.Inicia Me, False
    Me.txtPersona.Text = CodGrid
    Me.lblPersona.Caption = NomGrid
    If CodGrid <> "" Then Me.txtPolizaPer.SetFocus
End Sub

Private Sub cmdBuscar1_Click()
    frmBuscaCli.Inicia Me, False, , True
    Me.txtEmp.Text = CodGrid
    Me.lblNombre.Caption = NomGrid
    If CodGrid <> "" Then Me.cmdAgregar.SetFocus
End Sub

Private Sub cmdEliminar_Click()
    If lstEmpleados.ListIndex = -1 Then Exit Sub
    Me.lstEmpleados.RemoveItem lstEmpleados.ListIndex
End Sub

Private Sub cmdPago_Click()
    Dim lsMov1 As String
    Dim lsMov2 As String
    Dim lsCadena As String
    Dim sql As String
    Dim lsPersonas As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lnMontoComFam As Currency
    Dim lnMontoComPer As Currency
    
    If Me.txtPersona.Text = "" Then
        MsgBox "Debe ingresar a la empresa que brinda el servicio de Seguro.", vbInformation, "Aviso"
        cmdBuscar.SetFocus
        Exit Sub
    ElseIf Me.txtPolizaPer.Text = "" Then
        MsgBox "Debe ingresar el numero de Poliza del Personal.", vbInformation, "Aviso"
        txtPolizaPer.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskPer.Text) Then
        MsgBox "Debe ingresar la fecha de emision de la Poliza del Personal.", vbInformation, "Aviso"
        mskPer.SetFocus
        Exit Sub
    ElseIf Me.txtPolizaFam.Text = "" Then
        MsgBox "Debe ingresar el numero de Poliza de Familiares.", vbInformation, "Aviso"
        txtPolizaFam.SetFocus
        Exit Sub
    ElseIf Me.txtPolizaFam.Text = Me.txtPolizaPer.Text Then
        If MsgBox("Las Poliza son iguales. Desa Continuar ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            txtPolizaFam.SetFocus
            Exit Sub
        End If
    ElseIf Not IsDate(Me.mskFam.Text) Then
        MsgBox "Debe ingresar la fecha de emision de la Poliza del Familiares.", vbInformation, "Aviso"
        mskFam.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Ud. va realizar la provision de pago a " & Me.lblPersona.Caption & ", y generara los asientos contables corresondientes." & Chr(13) & "Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lsPersona = "''"
    If Me.lstEmpleados.ListCount > 0 Then
        For i = 0 To Me.lstEmpleados.ListCount - 1
            If i = 0 Then
                lsPersona = "'" & Left(Me.lstEmpleados.List(i), 10) & "'"
            Else
                lsPersona = lsPersona & ",'" & Left(Me.lstEmpleados.List(i), 10) & "'"
            End If
        Next i
        
        sql = " Select MontoPer, MontoFam From " _
            & " (Select Round((count(*) * (select nRanIniTab from dbcomunes..tablacod where cCodTab = 'FL01')) * 0.03 * 1.18,2) As MontoPer" _
            & " From Empleado" _
            & " Where cempest = '3' And cCodPers In (" & lsPersona & ")) As AAA," _
            & " (Select (Sum(nRanIniTab) - Round((count(*) * (select nRanIniTab From dbcomunes..tablacod where cCodTab = 'FL01')),2)) * 0.03 * 1.18 MontoFam" _
            & " From Empleado E" _
            & " Inner Join dbComunes..TablaCod TC On TC.cCodTab like 'ES__' And TC.cValor = E.cTipAsiMed" _
            & " where cempest = '3' And nRanIniTab  <> 0" _
            & " And cCodPers In (" & lsPersona & ")) AS BBB"
        rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
        
        lnMontoComFam = rs!MontoFam
        lnMontoComPer = rs!MontoPer
        
        rs.Close
    End If
    
    lsMov1 = GeneraAsientoAsistenciaMedica(629901, False, Me.txtPersona.Text, Me.txtPolizaPer.Text, Format(Me.mskPer.Text, gcFormatoFecha))
    lsMov2 = GeneraAsientoAsistenciaMedica(629902, True, Me.txtPersona.Text, Me.txtPolizaFam.Text, Format(Me.mskFam.Text, gcFormatoFecha))
    
    lsCadena = ""
    lsCadena = lsCadena & ImprimeAsientoContable(lsMov1) & Chr(12)
    lsCadena = lsCadena & ImprimeAsientoContable(lsMov2)
    
    R.Text = lsCadena
    frmPrevio.Previo R, Caption, True, 66
    
    If Me.chkRepo.Value = 1 Then
        lsCadena = GetRepAsistenciaMedica(False, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaPer.Text, Me.mskPer.Text, lnMontoComPer) & Chr(12)
        lsCadena = lsCadena & GetRepAsistenciaMedica(True, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaFam.Text, Me.mskFam.Text, lnMontoComFam) & Chr(12)
        R.Text = lsCadena
        frmPrevio.Previo R, Caption, True, 66
    End If
End Sub

Private Sub cmdReporte_Click()
    Dim lsCadena As String
    Dim sql As String
    Dim lsPersonas As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lnMontoComFam As Currency
    Dim lnMontoComPer As Currency
    
    If Me.txtPersona.Text = "" Then
        MsgBox "Debe ingresar a la empresa que brinda el servicio de Seguro.", vbInformation, "Aviso"
        cmdBuscar.SetFocus
        Exit Sub
    ElseIf Me.txtPolizaPer.Text = "" Then
        MsgBox "Debe ingresar el numero de Poliza del Personal.", vbInformation, "Aviso"
        txtPolizaPer.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskPer.Text) Then
        MsgBox "Debe ingresar la fecha de emision de la Poliza del Personal.", vbInformation, "Aviso"
        mskPer.SetFocus
        Exit Sub
    ElseIf Me.txtPolizaFam.Text = "" Then
        MsgBox "Debe ingresar el numero de Poliza de Familiares.", vbInformation, "Aviso"
        txtPolizaFam.SetFocus
        Exit Sub
    ElseIf Me.txtPolizaFam.Text = Me.txtPolizaPer.Text Then
        If MsgBox("Las Poliza son iguales. Desa Continuar ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            txtPolizaFam.SetFocus
            Exit Sub
        End If
    ElseIf Not IsDate(Me.mskFam.Text) Then
        MsgBox "Debe ingresar la fecha de emision de la Poliza del Familiares.", vbInformation, "Aviso"
        mskFam.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Ud. va realizar la provision de pago a " & Me.lblPersona.Caption & ", y generara los asientos contables corresondientes." & Chr(13) & "Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lsPersona = "''"
    If Me.lstEmpleados.ListCount > 0 Then
        For i = 0 To Me.lstEmpleados.ListCount - 1
            If i = 0 Then
                lsPersona = "'" & Left(Me.lstEmpleados.List(i), 10) & "'"
            Else
                lsPersona = lsPersona & ",'" & Left(Me.lstEmpleados.List(i), 10) & "'"
            End If
        Next i
        
        sql = " Select MontoPer, MontoFam From " _
            & " (Select Round((count(*) * (select nRanIniTab from dbcomunes..tablacod where cCodTab = 'FL01')) * 0.03 * 1.18,2) As MontoPer" _
            & " From Empleado" _
            & " Where cempest = '3' And cCodPers In (" & lsPersona & ")) As AAA," _
            & " (Select (Sum(nRanIniTab) - Round((count(*) * (select nRanIniTab From dbcomunes..tablacod where cCodTab = 'FL01')),2)) * 0.03 * 1.18 MontoFam" _
            & " From Empleado E" _
            & " Inner Join dbComunes..TablaCod TC On TC.cCodTab like 'ES__' And TC.cValor = E.cTipAsiMed" _
            & " where cempest = '3' And nRanIniTab  <> 0" _
            & " And cCodPers In (" & lsPersona & ")) AS BBB"
        rs.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
        
        lnMontoComFam = rs!MontoFam
        lnMontoComPer = rs!MontoPer
        
        rs.Close
    End If
    
    lsCadena = ""
    lsCadena = GetRepAsistenciaMedica(False, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaPer.Text, Me.mskPer.Text, lnMontoComPer) & Chr(12)
    lsCadena = lsCadena & GetRepAsistenciaMedica(True, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaFam.Text, Me.mskFam.Text, lnMontoComFam) & Chr(12)
    R.Text = lsCadena
    frmPrevio.Previo R, Caption, True, 66
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    AbreConexion
    Me.lstEmpleados.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub mskFam_GotFocus()
    mskFam.SelStart = 0
    mskFam.SelLength = 50
End Sub

Private Sub mskFam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdPago.SetFocus
    End If
End Sub

Private Sub mskPer_GotFocus()
    mskPer.SelStart = 0
    mskPer.SelLength = 50
End Sub

Private Sub mskPer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPolizaFam.SetFocus
    End If
End Sub

Private Sub txtPolizaPer_GotFocus()
    txtPolizaPer.SelStart = 0
    txtPolizaPer.SelLength = 50
End Sub

Private Sub txtPolizaPer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = intfNumero(KeyAscii)
    Else
        txtPolizaPer = Format(txtPolizaPer, "000000000")
        Me.mskPer.SetFocus
    End If
End Sub

Private Sub txtPolizaFam_GotFocus()
    txtPolizaFam.SelStart = 0
    txtPolizaFam.SelLength = 50
End Sub

Private Sub txtPolizaFam_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = intfNumero(KeyAscii)
    Else
        txtPolizaFam = Format(txtPolizaFam, "000000000")
        Me.mskFam.SetFocus
    End If
End Sub

