VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredAgricoSelec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actividad Agropecuaria"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "frmCredAgricoSelec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscarCooperativa 
      Caption         =   "..."
      Height          =   400
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar al Acreedor"
      Top             =   1100
      Width           =   420
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4680
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4680
      Width           =   1170
   End
   Begin VB.ComboBox cmbActividad 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin TabDlg.SSTab sstSubTipo 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Detalle de Actividad"
      TabPicture(0)   =   "frmCredAgricoSelec.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSubTipo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbSubtipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAnimales"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTotalHec"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtHectProd"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtHectProd 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtTotalHec 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtAnimales 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cmbSubtipo 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   3255
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   6000
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Animales (CGV):"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   2400
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Hectáreas en Producción:"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Total de Hectáreas:"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label lblSubTipo 
         AutoSize        =   -1  'True
         Caption         =   "Cultivo:"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   645
         Width           =   525
      End
   End
   Begin VB.Label lblCooperativa 
      AutoSize        =   -1  'True
      Caption         =   "Cooperativa"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblCodCooperativa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   120
      TabIndex        =   13
      Tag             =   "txtcodigo"
      Top             =   1100
      Width           =   1245
   End
   Begin VB.Label lblNomCooperativa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   1440
      TabIndex        =   12
      Tag             =   "txtnombre"
      Top             =   1100
      Width           =   4455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Actividad:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmCredAgricoSelec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lbRegistrar As Boolean
Private lnTipoAct As Long
Private lnSubTipoAct As Long
Private lnTotalHect As Double
Private lnHectProd As Double
Private lnAnimales As Long
Private lsCodCooperativa As String
Private lsNomCooperativa As String
Private lnMinAnimal As Long
Private fbConvenio As Boolean

Property Let Registrar(pReg As Boolean)
   lbRegistrar = pReg
End Property
Property Get Registrar() As Boolean
    Registrar = lbRegistrar
End Property

Property Let TipoAct(pTpoAct As Long)
   lnTipoAct = pTpoAct
End Property
Property Get TipoAct() As Long
    TipoAct = lnTipoAct
End Property

Property Let SubTipoAct(pSubTpoAct As Long)
   lnSubTipoAct = pSubTpoAct
End Property
Property Get SubTipoAct() As Long
    SubTipoAct = lnSubTipoAct
End Property

Property Let TotalHect(pnTotalHect As Double)
   lnTotalHect = pnTotalHect
End Property
Property Get TotalHect() As Double
    TotalHect = lnTotalHect
End Property

Property Let HectProd(pnHectProd As Double)
   lnHectProd = pnHectProd
End Property
Property Get HectProd() As Double
    HectProd = lnHectProd
End Property

Property Let Animales(pnAnimales As Double)
   lnAnimales = pnAnimales
End Property
Property Get Animales() As Double
    Animales = lnAnimales
End Property

Property Let CodCooperativa(psCodCooperativa As String)
   lsCodCooperativa = psCodCooperativa
End Property
Property Get CodCooperativa() As String
    CodCooperativa = lsCodCooperativa
End Property

Property Let NomCooperativa(psNomCooperativa As String)
   lsNomCooperativa = psNomCooperativa
End Property
Property Get NomCooperativa() As String
    NomCooperativa = lsNomCooperativa
End Property
's
Private Sub cmbActividad_Click()
Dim oParam As COMDCredito.DCOMParametro
Dim rsParam As ADODB.Recordset
Dim nTipo As Integer

    Set oParam = New COMDCredito.DCOMParametro
    
    sstSubTipo.Caption = "Detalle de Actividad " & Trim(Left(cmbActividad.Text, 25))
    
    nTipo = CInt(Trim(Right(cmbActividad.Text, 4)))
    Set rsParam = oParam.ObtenerParametrosAgro(nTipo, , "1")
    
    If fbConvenio Then
        lblCooperativa.Top = 840
        lblCodCooperativa.Top = 1100
        lblNomCooperativa.Top = 1100
        cmdBuscarCooperativa.Top = 1100
        sstSubTipo.Top = 1680
        
        If nTipo = 1 Then
            sstSubTipo.Height = 2055
            lblSubTipo.Caption = "Cultivo:"
            cmdAceptar.Top = 3840
            cmdCancelar.Top = 3840
            Me.Height = 4800
        Else
            sstSubTipo.Height = 2895
            lblSubTipo.Caption = "Subtipo:"
            cmdAceptar.Top = 4680
            cmdCancelar.Top = 4680
            Me.Height = 5655
        End If
    Else
        lblCooperativa.Top = 8000
        lblCodCooperativa.Top = 8000
        lblNomCooperativa.Top = 8000
        cmdBuscarCooperativa.Top = 8000
        sstSubTipo.Top = 840
        
        If nTipo = 1 Then
            lblSubTipo.Caption = "Cultivo:"
            sstSubTipo.Height = 2055
            cmdAceptar.Top = 3000
            cmdCancelar.Top = 3000
            Me.Height = 3960
        Else
            lblSubTipo.Caption = "Subtipo:"
            sstSubTipo.Height = 2895
            cmdAceptar.Top = 3840
            cmdCancelar.Top = 3840
            Me.Height = 4800
        End If
    End If
    
    cmbSubtipo.Clear
    Do While Not rsParam.EOF
        cmbSubtipo.AddItem Trim(rsParam!cSubTipo) & Space(100) & Trim(str(rsParam!nSubTipo))
        rsParam.MoveNext
    Loop
    rsParam.Close

    
End Sub

Private Sub cmbSubtipo_Click()
Dim oParam As COMDCredito.DCOMParametro
Dim rsParam As ADODB.Recordset
Dim nSubTipo As Integer

Set oParam = New COMDCredito.DCOMParametro
nSubTipo = CInt(Trim(Right(cmbSubtipo.Text, 4)))
Set rsParam = oParam.ObtenerParametrosAgro(0, , , nSubTipo)

If Not (rsParam.EOF And rsParam.BOF) Then
    lnMinAnimal = CLng(rsParam!nMin)
Else
    lnMinAnimal = 0
End If

End Sub

Private Sub CmdAceptar_Click()
If ValidaDatos Then
    'If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") =vbNo  Then Exit Sub
        lnTipoAct = CInt(Trim(Right(cmbActividad.Text, 4)))
        lnSubTipoAct = CInt(Trim(Right(cmbSubtipo.Text, 4)))
        lnTotalHect = CDbl(txtTotalHec.Text)
        lnHectProd = CDbl(txtHectProd.Text)
        
        If fbConvenio Then
            lsCodCooperativa = Trim(lblCodCooperativa.Caption)
            lsNomCooperativa = Trim(lblNomCooperativa.Caption)
        Else
            lsCodCooperativa = ""
            lsNomCooperativa = ""
        End If
        
        If lnTipoAct = 1 Then
            lnAnimales = 0
        Else
            lnAnimales = CLng(txtAnimales.Text)
        End If
        
        lbRegistrar = True
        Unload Me
End If
End Sub
Private Function ValidaDatos() As Boolean
If cmbActividad.Text = "" Then
    MsgBox "Seleccione el Tipo de Actividad", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If fbConvenio Then
    If lblCodCooperativa.Caption = "" Or lblNomCooperativa.Caption = "" Then
        MsgBox "Ingrese la cooperativa a la que Pertenece.", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
Else
    lblCodCooperativa.Caption = ""
    lblNomCooperativa.Caption = ""
End If

If cmbSubtipo.Text = "" Then
    MsgBox "Seleccione el SubTipo o Cultivo de Actividad", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If CDbl(txtTotalHec.Text) <= 0 Then
    MsgBox "Total de Hectáreas debe ser Mayor a 0", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If CDbl(txtHectProd.Text) <= 0 Then
    MsgBox "Hectáreas en Producción debe ser Mayor a 0", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If CDbl(txtTotalHec.Text) < CDbl(txtHectProd.Text) Then
    MsgBox "Hectáreas en Producción debe ser menor o igual al Total de Hectáreas", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If CInt(Trim(Right(cmbActividad.Text, 4))) = 2 Then
    If Trim(txtAnimales.Text) = "" Or Trim(txtAnimales.Text) = "0" Then
        MsgBox "Ingrese la Cantidad de Animales", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    If CLng(Trim(txtAnimales.Text)) < lnMinAnimal Then
        MsgBox "El número minímo de animales para " & Trim(Left(cmbSubtipo.Text, 25)) & " es " & lnMinAnimal & ".", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
End If

ValidaDatos = True
End Function



Private Sub cmdBuscarCooperativa_Click()
Dim oPersBuscada As COMDPersona.UCOMPersona
Set oPersBuscada = New COMDPersona.UCOMPersona
    lblCodCooperativa.Caption = ""
    lblNomCooperativa.Caption = ""
    
    Set oPersBuscada = frmBuscaPersona.Inicio
    If oPersBuscada Is Nothing Then Exit Sub
    lblCodCooperativa.Caption = oPersBuscada.sPersCod
    lblNomCooperativa.Caption = oPersBuscada.sPersNombre

    Set oPersBuscada = Nothing
End Sub

Private Sub cmdCancelar_Click()
    lbRegistrar = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oConstate As COMDConstantes.DCOMConstantes
    Set oConstate = New COMDConstantes.DCOMConstantes
    
    
    Llenar_Combo_con_Recordset oConstate.RecuperaConstantes(7067), cmbActividad

    cmbActividad.ListIndex = IndiceListaCombo(cmbActividad, lnTipoAct)
    cmbSubtipo.ListIndex = IndiceListaCombo(cmbSubtipo, lnSubTipoAct)
    
    txtTotalHec.Text = lnTotalHect
    txtHectProd.Text = lnHectProd
    txtAnimales.Text = lnAnimales
    lblCodCooperativa.Caption = lsCodCooperativa
    lblNomCooperativa.Caption = lsNomCooperativa
    
End Sub

Public Sub InsertaActDatos(ByVal psCtaCod As String, ByVal pnTipo As Integer, ByVal pnSubTipo As Integer, ByVal pnTotalHect As Double, ByVal pnHectProd As Double, ByVal pnAnimal As Long, ByVal psCodCooperativa As String)
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito

Call oCredito.InsActDatosAgricolasCred(psCtaCod, pnTipo, pnSubTipo, pnTotalHect, pnHectProd, pnAnimal, psCodCooperativa)
End Sub

Public Sub Inicia(Optional ByVal psCtaCod As String = "", Optional psTipoCred As String = "")
If Trim(psCtaCod) <> "" Then
    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsCredito As ADODB.Recordset
    
    Set oCredito = New COMDCredito.DCOMCredito
    
    Set rsCredito = oCredito.MostrarDatosAgricolasCred(psCtaCod)
    
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        lnTipoAct = CInt(rsCredito!nTipo)
        lnSubTipoAct = CInt(rsCredito!nSubTipo)
        lnTotalHect = CDbl(rsCredito!nTotalHect)
        lnHectProd = CDbl(rsCredito!nHectProd)
        lnAnimales = CLng(rsCredito!nAnimales)
        lsCodCooperativa = Trim(rsCredito!cCodCoopera)
        lsNomCooperativa = Trim(rsCredito!CooperaDesc)
    Else
        lnTipoAct = 1
        lnSubTipoAct = 0
        lnTotalHect = 0
        lnHectProd = 0
        lnAnimales = 0
        lsCodCooperativa = ""
        lsNomCooperativa = ""
    End If
    
    If psTipoCred = "602" Then 'Agropecuario por Convenio
        fbConvenio = True
    Else
        fbConvenio = False
    End If
    
Else
    If psTipoCred = "602" Then 'Agropecuario por Convenio
        fbConvenio = True
    Else
        fbConvenio = False
        If lnTipoAct = 0 Then
            lnTipoAct = 1
        End If
    End If
    Me.Show 1
End If

End Sub
Public Sub EliminaAgricolasCred(ByVal psCtaCod As String)
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito

Call oCredito.EliminaAgricolasCred(psCtaCod)
End Sub

Private Sub txtAnimales_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtHectProd_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtHectProd, KeyAscii)
End Sub

Private Sub txtTotalHec_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTotalHec, KeyAscii)
End Sub
