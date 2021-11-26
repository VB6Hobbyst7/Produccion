VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
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
      TabIndex        =   23
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
      TabIndex        =   3
      Top             =   6135
      Value           =   1  'Checked
      Width           =   2145
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   75
      Left            =   2865
      TabIndex        =   7
      Top             =   6270
      Visible         =   0   'False
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   132
      _Version        =   393217
      Enabled         =   -1  'True
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
      TabIndex        =   4
      Top             =   6150
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   15
      TabIndex        =   5
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtAjusteF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5220
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   2265
         Width           =   540
      End
      Begin VB.TextBox txtAjusteP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5220
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1875
         Width           =   540
      End
      Begin Sicmact.TxtBuscar txtPersona 
         Height          =   285
         Left            =   195
         TabIndex        =   24
         Top             =   1440
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   503
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin MSMask.MaskEdBox mskFam 
         Height          =   300
         Left            =   3480
         TabIndex        =   13
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
         Left            =   3480
         TabIndex        =   12
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
         Left            =   1125
         TabIndex        =   1
         Top             =   2265
         Width           =   1290
      End
      Begin VB.TextBox txtPolizaPer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         TabIndex        =   0
         Top             =   1875
         Width           =   1290
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "&Provision de Pago"
         Height          =   360
         Left            =   780
         TabIndex        =   2
         Top             =   2790
         Width           =   2190
      End
      Begin VB.Label Label1 
         Caption         =   "Doc.Fam :"
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
         Left            =   2565
         TabIndex        =   15
         Top             =   2295
         Width           =   1455
      End
      Begin VB.Label lblFecDocPer 
         Caption         =   "Doc.Pers:"
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
         Left            =   2565
         TabIndex        =   14
         Top             =   1890
         Width           =   930
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   6
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
      TabIndex        =   16
      Top             =   3150
      Width           =   6015
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   4905
         TabIndex        =   22
         Top             =   2055
         Width           =   975
      End
      Begin VB.ListBox lstEmpleados 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   120
         TabIndex        =   21
         Top             =   1410
         Width           =   4725
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   330
         Left            =   4905
         TabIndex        =   20
         Top             =   1590
         Width           =   975
      End
      Begin Sicmact.TxtBuscar txtEmp 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1065
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   7
         sTitulo         =   ""
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
        Me.txtEmp.SetFocus
        Exit Sub
    End If
    Me.lstEmpleados.AddItem Me.txtEmp.Text & " - " & Me.lblNombre.Caption
    Me.txtEmp.Text = ""
    Me.lblNombre.Caption = ""
End Sub

Private Sub CmdBuscar_Click()
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
    Dim Sql As String
    Dim lsPersonas As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    Dim oAsist As DActualizaAsistMedicaPrivada
    Set oAsist = New DActualizaAsistMedicaPrivada
    
    Dim lnMontoComFam As Currency
    Dim lnMontoComPer As Currency
    Dim oAsiento As NContImprimir
    Set oAsiento = New NContImprimir

    If Me.txtPersona.Text = "" Then
        MsgBox "Debe ingresar a la empresa que brinda el servicio de Seguro.", vbInformation, "Aviso"
        Me.txtPersona.SetFocus
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
        If MsgBox("Las Poliza son iguales. Desea Continuar ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            txtPolizaFam.SetFocus
            Exit Sub
        End If
    ElseIf Not IsDate(Me.mskFam.Text) Then
        MsgBox "Debe ingresar la fecha de emision de la Poliza del Familiares.", vbInformation, "Aviso"
        mskFam.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Ud. va realizar la provision de pago a " & Me.lblPersona.Caption & Chr(13) & ", y generara los asientos contables corresondientes." & Chr(13) & "Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    lsPersona = "''"
    If Me.lstEmpleados.ListCount > 0 Then
        For i = 0 To Me.lstEmpleados.ListCount - 1
            If i = 0 Then
                lsPersona = "'" & Left(Me.lstEmpleados.List(i), 10) & "'"
            Else
                lsPersona = lsPersona & ",'" & Left(Me.lstEmpleados.List(i), 10) & "'"
            End If
        Next i
        
        Set rs = oAsist.GetDevAMP(lsPersona)
        
        lnMontoComFam = IIf(IsNull(rs!MontoFam), 0, rs!MontoFam)
        lnMontoComPer = IIf(IsNull(rs!MontoPer), 0, rs!MontoPer)
        
        rs.Close
    End If
    
    lsMov1 = oAsist.GeneraAsientoAsistenciaMedica(629901, False, Me.txtPersona.Text, Me.txtPolizaPer.Text, Format(Me.mskPer.Text, gcFormatoFecha), gdFecSis, gsCodUser, gsCodAge, Me.txtAjusteP.Text)
    lsMov2 = oAsist.GeneraAsientoAsistenciaMedica(629902, True, Me.txtPersona.Text, Me.txtPolizaFam.Text, Format(Me.mskFam.Text, gcFormatoFecha), gdFecSis, gsCodUser, gsCodAge, Me.txtAjusteF.Text)
    
    lsCadena = ""
    lsCadena = lsCadena & oAsiento.ImprimeAsientoContable(lsMov1, 66, 80) & oImpresora.gPrnSaltoPagina
    lsCadena = lsCadena & oAsiento.ImprimeAsientoContable(lsMov2, 66, 80)
    
    oPrevio.Show lsCadena, Caption, True
    
    If Me.chkRepo.value = 1 Then
        lsCadena = oAsist.GetRepAsistenciaMedica(gdFecSis, gsEmpresa, gsNomAge, Me.lblPersona.Caption, False, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaPer.Text, Me.mskPer.Text, lnMontoComPer, Me.txtAjusteP.Text) & oImpresora.gPrnSaltoPagina
        lsCadena = lsCadena & oAsist.GetRepAsistenciaMedica(gdFecSis, gsEmpresa, gsNomAge, Me.lblPersona.Caption, True, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaFam.Text, Me.mskFam.Text, lnMontoComFam, Me.txtAjusteF.Text) & oImpresora.gPrnSaltoPagina
        R.Text = lsCadena
        oPrevio.Show lsCadena, Caption, True
    End If
End Sub

Private Sub cmdReporte_Click()
    Dim lsCadena As String
    Dim Sql As String
    Dim lsPersonas As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    
    Dim lnMontoComFam As Currency
    Dim lnMontoComPer As Currency
    
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    
    Dim oAsist As DActualizaAsistMedicaPrivada
    Set oAsist = New DActualizaAsistMedicaPrivada
    
    If Me.txtPersona.Text = "" Then
        MsgBox "Debe ingresar a la empresa que brinda el servicio de Seguro.", vbInformation, "Aviso"
        txtPersona.SetFocus
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
        If MsgBox("Las Poliza son iguales. Desea Continuar ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            txtPolizaFam.SetFocus
            Exit Sub
        End If
    ElseIf Not IsDate(Me.mskFam.Text) Then
        MsgBox "Debe ingresar la fecha de emision de la Poliza del Familiares.", vbInformation, "Aviso"
        mskFam.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtAjusteP) Then
        MsgBox "Debe ingresar un ajuste valido.", vbInformation, "Aviso"
        txtAjusteP.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtAjusteF) Then
        MsgBox "Debe ingresar un ajuste valido.", vbInformation, "Aviso"
        txtAjusteF.SetFocus
        Exit Sub
    End If
    
    lsPersona = "''"
    If Me.lstEmpleados.ListCount > 0 Then
        For i = 0 To Me.lstEmpleados.ListCount - 1
            If i = 0 Then
                lsPersona = "'" & Left(Me.lstEmpleados.List(i), 13) & "'"
            Else
                lsPersona = lsPersona & ",'" & Left(Me.lstEmpleados.List(i), 13) & "'"
            End If
        Next i
        
        Sql = " Select Isnull(MontoPer,0) MontoPer, Isnunll(MontoFam,0) MontoFam " _
            & " From  (Select Round((count(*) * (select dbo.GetAsistTit())) * 0.03 * 1.18,2)" _
            & " As MontoPer From rrhh Where nRHEstado Like '[78]%' And cPersCod In (" & lsPersona & ")) As AAA," _
            & " (Select (Sum(nRHAsistMedPrivMonto) - Round((count(*) * (select dbo.GetAsistTit())),2)) * 0.03 * 1.18 MontoFam" _
            & " From RRHH E" _
            & " Inner Join RHEmpleado RHE On RHE.cPersCod = E.cPersCod" _
            & " Inner Join RHAsistMedPrivTabla TC On TC.cRHAsistMedPrivCod = RHE.cRHEmplAMPCod" _
            & " where nRHEstado Like  '[78]%' And nRHAsistMedPrivMonto  <> 0 And RHE.cPersCod In (" & lsPersona & ")) AS BBB"

        Set rs = oAsist.GetDevAMP(lsPersona)
        
        lnMontoComFam = rs!MontoFam
        lnMontoComPer = rs!MontoPer
        
        rs.Close
    End If
    
    lsCadena = ""
    lsCadena = oAsist.GetRepAsistenciaMedica(gdFecSis, gsEmpresa, gsNomAge, Me.lblPersona.Caption, False, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaPer.Text, Me.mskPer.Text, lnMontoComPer, txtAjusteP) & oImpresora.gPrnSaltoPagina
    lsCadena = lsCadena & oAsist.GetRepAsistenciaMedica(gdFecSis, gsEmpresa, gsNomAge, Me.lblPersona.Caption, True, Me.txtPersona.Text, Me.lblPersona.Caption, Me.txtPolizaFam.Text, Me.mskFam.Text, lnMontoComFam, txtAjusteF) & oImpresora.gPrnSaltoPagina
    oPrevio.Show lsCadena, Caption, True
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCon As DConecta
Set oCon = New DConecta

    oCon.AbreConexion
    Me.lstEmpleados.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim oCon As DConecta
Set oCon = New DConecta
    oCon.CierraConexion
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

Private Sub txtAjusteP_GotFocus()
    txtAjusteP.SelStart = 0
    txtAjusteP.SelLength = 50
End Sub

Private Sub txtAjusteP_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtAjusteP, KeyAscii)
End Sub

Private Sub txtAjusteF_GotFocus()
    txtAjusteF.SelStart = 0
    txtAjusteF.SelLength = 50
End Sub

Private Sub txtAjusteF_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtAjusteF, KeyAscii)
End Sub

Private Sub txtEmp_EmiteDatos()
    Me.lblNombre.Caption = Me.txtEmp.psDescripcion
End Sub

Private Sub txtPersona_EmiteDatos()
    Me.lblPersona.Caption = Me.txtPersona.psDescripcion
End Sub

Private Sub txtPersona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPolizaPer.SetFocus
    End If
End Sub

Private Sub txtPolizaPer_GotFocus()
    txtPolizaPer.SelStart = 0
    txtPolizaPer.SelLength = 50
End Sub

Private Sub txtPolizaPer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = NumerosEnteros(KeyAscii)
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
        KeyAscii = NumerosEnteros(KeyAscii)
    Else
        txtPolizaFam = Format(txtPolizaFam, "000000000")
        Me.mskFam.SetFocus
    End If
End Sub

