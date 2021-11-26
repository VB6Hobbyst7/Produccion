VERSION 5.00
Begin VB.Form frmLogProvCtas 
   BackColor       =   &H00400040&
   BorderStyle     =   0  'None
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   Icon            =   "frmLogProvCtas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4950
      MaskColor       =   &H008080FF&
      TabIndex        =   9
      Top             =   5600
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3750
      TabIndex        =   8
      Top             =   5600
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   55
      TabIndex        =   28
      Top             =   5400
      Width           =   6030
   End
   Begin VB.Frame fraCta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUENTAS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4650
      Left            =   50
      TabIndex        =   1
      Top             =   800
      Width           =   6060
      Begin VB.TextBox txtCtaCCIME 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   4080
         Width           =   5655
      End
      Begin VB.TextBox txtCtaDetracMN 
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   5655
      End
      Begin VB.TextBox txtCtaCCIMN 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   5655
      End
      Begin VB.ComboBox cboCtaDolares 
         BackColor       =   &H0080FF80&
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3480
         Width           =   5655
      End
      Begin VB.ComboBox cboCtaSoles 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   820
         Width           =   5655
      End
      Begin VB.TextBox txtCtaME 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   5655
      End
      Begin VB.TextBox txtCtaMN 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   5655
      End
      Begin VB.CheckBox chkME 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5580
         TabIndex        =   12
         Top             =   1425
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CheckBox chkMN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5580
         TabIndex        =   11
         Top             =   465
         Visible         =   0   'False
         Width           =   270
      End
      Begin Sicmact.TxtBuscar txtME 
         Height          =   360
         Left            =   2880
         TabIndex        =   5
         Top             =   4440
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   3413
         _ExtentY        =   635
         Appearance      =   0
         BackColor       =   8454016
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
         TipoBusqueda    =   2
      End
      Begin Sicmact.TxtBuscar txtMN 
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   4440
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   3413
         _ExtentY        =   635
         Appearance      =   0
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
         TipoBusqueda    =   2
         sTitulo         =   ""
      End
      Begin Sicmact.TxtBuscar txtBanco1 
         Height          =   360
         Left            =   240
         TabIndex        =   15
         Top             =   450
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   635
         Appearance      =   0
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
         sTitulo         =   ""
      End
      Begin Sicmact.TxtBuscar txtBanco2 
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   635
         Appearance      =   0
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
         sTitulo         =   ""
      End
      Begin Sicmact.TxtBuscar txtBanco1Detrac 
         Height          =   360
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   635
         Appearance      =   0
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
         sTitulo         =   ""
      End
      Begin VB.Label lblBancoDetrac 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Banco1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   29
         Top             =   1970
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CCI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cta Detraccion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CCI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblBanco2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Banco2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   18
         Top             =   3060
         Width           =   3135
      End
      Begin VB.Label lblBanco1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Banco1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   17
         Top             =   400
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moneda Extranjera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label lblMN 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moneda Nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   200
         Width           =   2295
      End
   End
   Begin Sicmact.TxtBuscar txtAge 
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   3413
      _ExtentY        =   635
      Appearance      =   0
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
      sTitulo         =   ""
   End
   Begin VB.Label lblNomPers 
      BackColor       =   &H00400040&
      Caption         =   "cpersNombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   60
      TabIndex        =   10
      Top             =   400
      Width           =   6060
   End
   Begin VB.Label lblPersCod 
      BackColor       =   &H00400040&
      Caption         =   "cperscod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6060
   End
   Begin VB.Label lblAgencia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   4365
   End
End
Attribute VB_Name = "frmLogProvCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodigoProveedor As String
Dim lsPersCod As String
Dim lsPersNombre As String
Dim oConstSist As NConstSistemas 'WIOR 20130131
Dim sPersCodCMACMaynas As String 'WIOR 20130131

Public Sub Ini(psPersCod As String, psPersNombre As String, psCtaCodMN As String, psCtaCodME As String, psCtaProv As String)
    lsCodigoProveedor = psCtaProv
    lsPersCod = psPersCod
    lsPersNombre = psPersNombre
    'WIOR 20130131 ***************************
    Set oConstSist = New NConstSistemas
    sPersCodCMACMaynas = oConstSist.LeeConstSistema(41)
    Set oConstSist = Nothing
    'WIOR FIN ********************************
    Me.Show 1
End Sub
'WIOR 20130131 ***************************
Private Sub cboCtaDolares_Click()
Me.txtCtaME.Text = Trim(Me.cboCtaDolares.Text)
End Sub

Private Sub cboCtaSoles_Click()
Me.txtCtaMN.Text = Trim(Me.cboCtaSoles.Text)
End Sub
'WIOR FIN ********************************
Private Sub cmdAceptar_Click()
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    If Not ValidaDatos Then Exit Sub
    'If MsgBox("¿Está seguro de grabar las cuentas?, solo se grabaran las cuentas marcadas con check, solo se puede grabar una cuenta por moneda.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub 'Comentado xPASI20150914
    If MsgBox("¿Está seguro de grabar las cuentas?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    'oProv.SetProvCtas Me.lblPersCod.Caption, IIf(Me.chkMN.value = 1, Me.txtMN.Text, ""), IIf(Me.chkME.value = 1, Me.txtME.Text, ""), GetMovNro(gsCodUser, gsCodAge)
    'oProv.SetProvCtasBancos Me.lblPersCod.Caption, Trim(Me.txtBanco1.Text), txtCtaMN.Text, Trim(Me.txtBanco2.Text), Me.txtCtaME.Text, GetMovNro(gsCodUser, gsCodAge)
    oProv.SetProvCtasBancos Me.lblPersCod.Caption, Trim(Me.txtBanco1.Text), txtCtaMN.Text, Trim(Me.txtBanco2.Text), Me.txtCtaME.Text, GetMovNro(gsCodUser, gsCodAge), Trim(txtCtaCCIMN.Text), Trim(txtCtaCCIME), Trim(txtBanco1Detrac.Text), Trim(txtCtaDetracMN) 'PASIERS0472015*******************
    Unload Me
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If Len(Trim(txtBanco1.Text)) > 0 Then
        If Trim(txtBanco1.Text) = sPersCodCMACMaynas And cboCtaSoles.ListIndex = -1 Then
            MsgBox "Asegurese de haber seleccionado la Cuenta de CMACMaynas en Soles.", vbInformation, "Aviso"
            cboCtaSoles.SetFocus
            Exit Function
        Else
            If Len(Trim(txtCtaMN.Text)) = 0 Then
                MsgBox "Asegurese de haber Ingresado la Cuenta de Banco en Soles.", vbInformation, "Aviso"
                txtCtaMN.SetFocus
                Exit Function
            End If
        End If
    End If
   If Len(Trim(txtBanco2.Text)) > 0 Then
        If Trim(txtBanco2.Text) = sPersCodCMACMaynas And cboCtaDolares.ListIndex = -1 Then
            MsgBox "Asegurese de haber seleccionado la Cuenta de CMACMaynas en Dólares.", vbInformation, "Aviso"
            cboCtaDolares.SetFocus
            Exit Function
        Else
            If Len(Trim(txtCtaME.Text)) = 0 Then
                MsgBox "Asegurese de haber Ingresado la Cuenta de Banco en Dólares.", vbInformation, "Aviso"
                txtCtaME.SetFocus
                Exit Function
            End If
        End If
    End If
    ValidaDatos = True
End Function
Private Sub cmdCancelar_Click()
    lblBanco1.Caption = ""
    lblBanco2.Caption = ""
    txtCtaMN.Text = ""
    txtBanco1.Text = ""
    txtCtaME.Text = ""
    txtBanco2.Text = ""
    Unload Me
End Sub

Private Sub Form_Load()
'    Dim oCon As DConstantes
'    Set oCon = New DConstantes
'    Me.txtAge.rs = oCon.GetAgencias(, , True)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

Me.lblPersCod.Caption = lsPersCod
Me.lblNomPers.Caption = lsPersNombre
Me.Caption = "Cuentas de Proveedor"
   
If lsCodigoProveedor <> "" Then
    'Cargamos Cuentas del Proveedor
    '******************************
    lblBanco1.Caption = GetDescBancoProv(lsCodigoProveedor, 1)
    lblBanco2.Caption = GetDescBancoProv(lsCodigoProveedor, 2)
    lblBancoDetrac.Caption = GetDescBancoProv(lsCodigoProveedor, 0, True) 'PASIERS0472015
    
    Set rs = GetCtaBancoProv(lsCodigoProveedor, 1)
    'txtCtaMN.Text = rs(1)
    txtBanco1.Text = rs(0)
    'WIOR 20130131 *************************************
    Dim bCuentaAct As Boolean
    bCuentaAct = False
      If Trim(txtBanco1.Text) = sPersCodCMACMaynas Then
          Me.txtCtaMN.Visible = False
          Dim RsCtaMN As ADODB.Recordset
          Set RsCtaMN = GetCuentasProveedor(lsPersCod, 1)
          cboCtaSoles.Clear
          Do While Not RsCtaMN.EOF
                If Trim(RsCtaMN!cCtaCod) = Trim(rs(1)) Then
                    bCuentaAct = True
                End If
                cboCtaSoles.AddItem Trim(RsCtaMN!cCtaCod)
                RsCtaMN.MoveNext
          Loop
          'cboCtaSoles.Left = 2640 Comentado xPASI20150611
          cboCtaSoles.Visible = True 'PASI20150611
          txtCtaMN.Text = rs(1)
          If txtCtaMN.Text <> "" And bCuentaAct Then
            cboCtaSoles.Text = rs(1)
          Else
            If txtCtaMN.Text <> "" Then
                cboCtaSoles.AddItem rs(1)
                cboCtaSoles.Text = rs(1)
            End If
          End If
          txtCtaCCIMN.Text = rs(2) 'PASIERS0472015
      Else
          Me.txtCtaMN.Visible = True
          'cboCtaSoles.Left = 6120 comentado xPASI20150611
          cboCtaSoles.Visible = False 'PASI20150611
          cboCtaSoles.Clear
          txtCtaMN.Text = rs(1)
          txtCtaCCIMN.Text = rs(2) 'PASIERS0472015
      End If
    'WIOR FIN ******************************************
    
    'PASIERS0472015********************************
    Set rs = GetCtaBancoProv(lsCodigoProveedor, 3)
    txtBanco1Detrac.Text = rs(0)
    txtCtaDetracMN.Text = rs(1)
    'END PASI******************************************
    
    Set rs = GetCtaBancoProv(lsCodigoProveedor, 2)
    'txtCtaME.Text = rs(1)
    txtBanco2.Text = rs(0)
    'WIOR 20130131 *************************************
      bCuentaAct = False
      If Trim(txtBanco2.Text) = sPersCodCMACMaynas Then
          Me.txtCtaME.Visible = False
          Dim RsCtaME As ADODB.Recordset
          Set RsCtaME = GetCuentasProveedor(lsPersCod, 2)
          cboCtaDolares.Clear
          Do While Not RsCtaME.EOF
                If Trim(RsCtaME!cCtaCod) = Trim(rs(1)) Then
                    bCuentaAct = True
                End If
                cboCtaDolares.AddItem Trim(RsCtaME!cCtaCod)
                RsCtaME.MoveNext
          Loop
          'cboCtaDolares.Left = 2640 Comentado xPASI20150611
          cboCtaDolares.Visible = True 'PASI20150611
          txtCtaME.Text = rs(1)
          If txtCtaME.Text <> "" And bCuentaAct Then
            cboCtaDolares.Text = rs(1)
          Else
            If txtCtaME.Text <> "" Then
                cboCtaDolares.AddItem rs(1)
                cboCtaDolares.Text = rs(1)
            End If
          End If
          txtCtaCCIME.Text = rs(2) 'PASIERS0472015
      Else
          Me.txtCtaME.Visible = True
          'cboCtaDolares.Left = 6120 comentado xPASI20150611
          cboCtaDolares.Visible = False 'PASI20150611
          cboCtaDolares.Clear
          txtCtaME.Text = rs(1)
          txtCtaCCIME.Text = rs(2) 'PASIERS0472015
      End If
    'WIOR FIN ******************************************
    lsCodigoProveedor = ""
End If
        
    'Cargamos Bancos
    '*****************
'    txtBanco1.psRaiz = "BANCOS"
'    txtBanco1.rs = GetInstFinancieras("01")
'
'    txtBanco2.psRaiz = "BANCOS"
'    txtBanco2.rs = GetInstFinancieras("01")
'   COMENTADO POR WIOR 20130131
    
    'WIOR 20130131 ****************************************
    txtBanco1.psRaiz = "BANCOS Y CMACS"
    txtBanco1.rs = GetInstFinancieras("0[13]", True)
    
    txtBanco2.psRaiz = "BANCOS Y CMACS"
    txtBanco2.rs = GetInstFinancieras("0[13]", True)
    'WIOR FIN *********************************************
    'PASIERS0472015***********************************
    txtBanco1Detrac.psRaiz = "BANCOS Y CMACS"
    txtBanco1Detrac.rs = GetInstFinancieras("0[1]", True)
    'END PASI

End Sub


Public Function GetInstFinancieras(Optional ByVal psFiltroTipoCtaIF As String = "", Optional ByVal pbMasFinan As Boolean = False) As Recordset 'WIOR 20130131 AGREGO pbMasFinan
Dim oConec As DConecta
Dim Sql As String
Dim rs As ADODB.Recordset
Dim lsIFFiltro As String
Dim lsFiltroCta As String
Dim Pos As String
Dim lsCadAux As String
Dim lsFiltroTipoIF As String
lsIFFiltro = ""
If psFiltroTipoCtaIF <> "" Then
    lsIFFiltro = " WHERE I.cIfTpo LIKE '" & psFiltroTipoCtaIF & "' "
End If

Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function


'WIOR 20130131 ****************************************
If pbMasFinan Then
    Sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
        & " From " _
        & "         (   SELECT  I.cIFTpo,I.cIFTpo cPersCod, c.cConsDescripcion cPersNombre, 1 as Nivel " _
        & "             FROM    INSTITUCIONFINANC I JOIN Constante c ON c.nConsValor = convert(int,I.cIFTpo) and c.nconscod like '" & gCGTipoIF & "' " _
        & "         " & lsIFFiltro & "" _
        & "             GROUP BY I.cIFTpo, c.cConsDescripcion " _
        & "             UNION ALL " _
        & "             SELECT  I.cIFTpo,P.cPersCod cPersCod, CONVERT(CHAR(200),P.cPersNombre ) AS cPersNombre , 2 AS Nivel " _
        & "             FROM    INSTITUCIONFINANC I  " _
        & "                     JOIN PERSONA P ON P.cPersCod = I.cPersCod  " & lsIFFiltro _
        & "         ) AS INSTFIN  " _
        & "     ORDER BY cIFTpo,cPersCod "
Else
    Sql = "SELECT  cPersCod  AS cCodigo, cPersNombre as Descripcion, Nivel " _
        & " From " _
        & "         (   SELECT  I.cIFTpo cPersCod, c.cConsDescripcion cPersNombre, 1 as Nivel " _
        & "             FROM    INSTITUCIONFINANC I JOIN Constante c ON c.nConsValor = convert(int,I.cIFTpo) and c.nconscod like '" & gCGTipoIF & "' " _
        & "         " & lsIFFiltro & "" _
        & "             GROUP BY I.cIFTpo, c.cConsDescripcion " _
        & "             UNION ALL " _
        & "             SELECT  P.cPersCod cPersCod, CONVERT(CHAR(200),P.cPersNombre ) AS cPersNombre , 2 AS Nivel " _
        & "             FROM    INSTITUCIONFINANC I  " _
        & "                     JOIN PERSONA P ON P.cPersCod = I.cPersCod  " & lsIFFiltro _
        & "         ) AS INSTFIN  " _
        & "     ORDER BY cPersCod "
End If
'WIOR FIN *********************************************
   
Set rs = oConec.CargaRecordSet(Sql)
Set GetInstFinancieras = rs
oConec.CierraConexion
Set oConec = Nothing
End Function

Private Sub txtAge_EmiteDatos()
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    
    Me.lblAgencia.Caption = Me.txtAge.psDescripcion
    
    If txtAge <> "" Then
        Me.txtMN.rs = oProv.GetProvCtas(Me.lblPersCod.Caption, Me.txtAge.Text, gMonedaNacional, gbBitCentral)
        Me.txtME.rs = oProv.GetProvCtas(Me.lblPersCod.Caption, Me.txtAge.Text, gMonedaExtranjera, gbBitCentral)
    End If
    
End Sub
Private Sub txtBanco1_EmiteDatos()
  lblBanco1 = txtBanco1.psDescripcion
  'WIOR 20130131 *************************************
    If Trim(txtBanco1.Text) = sPersCodCMACMaynas Then
        txtCtaMN.Visible = False
        txtCtaMN.Text = ""
        Dim RsCta As ADODB.Recordset
        Set RsCta = GetCuentasProveedor(lsPersCod, 1)
        cboCtaSoles.Clear
        Do While Not RsCta.EOF
            cboCtaSoles.AddItem Trim(RsCta!cCtaCod)
            RsCta.MoveNext
        Loop
        'cboCtaSoles.Left = 2640 comentado PASI20150611
        cboCtaSoles.Visible = True
    Else
        Me.txtCtaMN.Text = ""
        Me.txtCtaMN.Visible = True
        'cboCtaSoles.Left = 6120 comentado xPASI20150611
        cboCtaSoles.Visible = False 'PASI20150611
        cboCtaSoles.Clear
    End If
  'WIOR FIN ******************************************
End Sub
Private Sub txtBanco2_EmiteDatos()
  lblBanco2 = txtBanco2.psDescripcion
  'WIOR 20130131 *************************************
    If Trim(txtBanco2.Text) = sPersCodCMACMaynas Then
        txtCtaME.Visible = False
        txtCtaME.Text = ""
        Dim RsCta As ADODB.Recordset
        Set RsCta = GetCuentasProveedor(lsPersCod, 2)
        cboCtaDolares.Clear
        Do While Not RsCta.EOF
            cboCtaDolares.AddItem Trim(RsCta!cCtaCod)
            RsCta.MoveNext
        Loop
        'cboCtaDolares.Left = 2640 comentado xPASI20150611
        cboCtaDolares.Visible = True 'PASI20150611
    Else
        txtCtaME.Text = ""
        Me.txtCtaME.Visible = True
        'cboCtaDolares.Left = 6120 comentado xPASI20150611
        cboCtaDolares.Visible = False
        cboCtaDolares.Clear
    End If
  'WIOR FIN ******************************************
End Sub

Private Sub txtME_GotFocus()
    If txtAge.Text = "" Then
        MsgBox "Debe elegir la agencia donde buscar las cuentas.", vbInformation, "Aviso"
        Me.txtAge.SetFocus
    End If
End Sub

Private Sub txtMN_GotFocus()
    If txtAge.Text = "" Then
        MsgBox "Debe elegir la agencia donde buscar las cuentas.", vbInformation, "Aviso"
        Me.txtAge.SetFocus
    End If
End Sub

'WIOR 20130131 *************************************
Public Function GetCuentasProveedor(ByVal psPersCod As String, ByVal pnMoneda As Integer) As Recordset
Dim oConec As DConecta
Dim Sql As String
Dim rs As ADODB.Recordset


Set oConec = New DConecta
Set rs = New ADODB.Recordset
If oConec.AbreConexion() = False Then Exit Function

    Sql = "exec stp_sel_GetCuentasProveedor '" & psPersCod & "','" & pnMoneda & "' "

Set rs = oConec.CargaRecordSet(Sql)
Set GetCuentasProveedor = rs
oConec.CierraConexion
Set oConec = Nothing
End Function
'WIOR FIN ******************************************


