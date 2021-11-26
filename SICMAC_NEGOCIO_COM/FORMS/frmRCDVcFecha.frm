VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRCDVcFecha 
   Caption         =   "Cambios de RCD"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   9945
   Icon            =   "frmRCDVcFecha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9945
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   17
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   16
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   5640
      Width           =   1335
   End
   Begin TabDlg.SSTab sTabRCD 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "frmRCDVcFecha.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Crédito"
      TabPicture(1)   =   "frmRCDVcFecha.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   9495
         Begin VB.TextBox txtDocCom 
            Height          =   330
            Left            =   2040
            TabIndex        =   30
            Top             =   3120
            Width           =   1575
         End
         Begin VB.ComboBox cboTDC 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2760
            Width           =   3615
         End
         Begin VB.ComboBox cboMagnitud 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2040
            Width           =   3615
         End
         Begin VB.TextBox txtPersNom 
            Height          =   330
            Left            =   2040
            TabIndex        =   27
            Top             =   240
            Width           =   7215
         End
         Begin VB.TextBox txtApellidoPaterno 
            Height          =   330
            Left            =   2040
            TabIndex        =   26
            Top             =   600
            Width           =   7215
         End
         Begin VB.TextBox txtApellidoMaterno 
            Height          =   330
            Left            =   2040
            TabIndex        =   25
            Top             =   960
            Width           =   7215
         End
         Begin VB.TextBox txtNombre1 
            Height          =   330
            Left            =   2040
            TabIndex        =   24
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtNombre2 
            Height          =   330
            Left            =   6720
            TabIndex        =   23
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtApellidoCasada 
            Height          =   330
            Left            =   2040
            TabIndex        =   22
            Top             =   1320
            Width           =   7215
         End
         Begin VB.TextBox txtCodSBS 
            Height          =   330
            Left            =   2040
            TabIndex        =   21
            Top             =   3480
            Width           =   1575
         End
         Begin VB.TextBox txtRelInst 
            Height          =   330
            Left            =   6720
            TabIndex        =   20
            Top             =   2400
            Width           =   1575
         End
         Begin VB.ComboBox cboSexo 
            Height          =   315
            ItemData        =   "frmRCDVcFecha.frx":0342
            Left            =   6720
            List            =   "frmRCDVcFecha.frx":034F
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2760
            Width           =   1575
         End
         Begin MSMask.MaskEdBox txtFechaNaci 
            Height          =   330
            Left            =   2040
            TabIndex        =   31
            Top             =   2400
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            Caption         =   "Documento Compl.:"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   3180
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Documento Compl.:"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   2800
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Nacimiento:"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   2460
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Magnitud :"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2100
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Nombre Completo:"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Apellido Paterno/Raz. S:"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   660
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   "Apellido Materno:"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Nombre 1:"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1740
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Nombre 2"
            Height          =   255
            Left            =   6000
            TabIndex        =   36
            Top             =   1755
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Apellido Casada:"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Codigo SBS:"
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   3540
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "cRelInst"
            Height          =   255
            Left            =   6000
            TabIndex        =   33
            Top             =   2460
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Sexo :"
            Height          =   255
            Left            =   6000
            TabIndex        =   32
            Top             =   2800
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Otros Datos"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   4
         Top             =   2640
         Width           =   9135
         Begin VB.ComboBox cboUbigeo 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   5415
         End
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label Label14 
            Caption         =   "Ubigeo"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Agencia"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   320
            Width           =   735
         End
      End
      Begin VB.Frame frame3 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   9135
         Begin SICMACT.FlexEdit FeCreditos 
            Height          =   1935
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   8415
            _extentx        =   14843
            _extenty        =   3413
            cols0           =   6
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "Item-Credito-Cuenta-Saldo-Agencia-Ubigeo"
            encabezadosanchos=   "0-2300-2000-1500-1000-1000"
            font            =   "frmRCDVcFecha.frx":03CF
            font            =   "frmRCDVcFecha.frx":03FB
            font            =   "frmRCDVcFecha.frx":0427
            font            =   "frmRCDVcFecha.frx":0453
            font            =   "frmRCDVcFecha.frx":047F
            fontfixed       =   "frmRCDVcFecha.frx":04AB
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-X-X-X-4-X"
            listacontroles  =   "0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-L-R-C-C"
            formatosedit    =   "0-0-0-4-0-0"
            textarray0      =   "Item"
            lbeditarflex    =   -1
            lbflexduplicados=   0
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Busqueda"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtNumSec 
         Height          =   330
         Left            =   3480
         MaxLength       =   8
         TabIndex        =   11
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8280
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboNomTabla 
         Height          =   315
         ItemData        =   "frmRCDVcFecha.frx":04D9
         Left            =   6000
         List            =   "frmRCDVcFecha.frx":04E3
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtFechaCierre 
         Height          =   330
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label18 
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   5280
         TabIndex        =   45
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Secuencia :"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRCDVcFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRs01 As ADODB.Recordset
Dim oRs02 As ADODB.Recordset

Private Sub cboAgencia_Click()
Dim i As Integer
If Trim(cboAgencia.Text) <> "" Then
    For i = 1 To FeCreditos.Rows - 1
        FeCreditos.TextMatrix(i, 4) = "00" & Trim(Right(cboAgencia.Text, 4))
    Next i
End If
End Sub

Private Sub cboUbigeo_Click()
Dim i As Integer
If Trim(cboUbigeo.Text) <> "" Then
    For i = 1 To FeCreditos.Rows - 1
        FeCreditos.TextMatrix(i, 5) = Trim(Right(cboUbigeo.Text, 20))
    Next i
End If
End Sub

Private Sub cmdBuscar_Click()
    'JUEZ 20150310 ***************************
    If Trim(txtNumSec.Text) = "" Then
        MsgBox "No digitó el numero de secuencia", vbInformation, "Aviso"
        txtNumSec.SetFocus
        Exit Sub
    End If
    If Trim(cboNomTabla.Text) = "" Then
        MsgBox "No seleccionó el tipo de búsqueda", vbInformation, "Aviso"
        cboNomTabla.SetFocus
        Exit Sub
    End If
    'END JUEZ ********************************
    Call ObtenerVC
End Sub

Private Sub CmdModificar_Click()
Dim oRCD As COMNCredito.NCOMRCD
Dim i As Integer
Set oRCD = New COMNCredito.NCOMRCD
    'Call oRCD.ActualizarVc01(gdFecData, txtNumSec.Text, Trim(Right(cboMagnitud.Text, 2)), txtFechaNaci.Text, Trim(Right(cboTDC.Text, 4)), txtDocCom.Text, txtPersNom.Text, txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtApellidoCasada.Text, txtNombre1.Text, txtNombre2.Text, txtCodSBS.Text, txtRelInst.Text, Trim(Right(cboSexo.Text, 2))) 'JUEZ 20140410 Se agregó cboSexo
    Call oRCD.ActualizarVc01(gdFecData, txtNumSec.Text, Trim(Right(cboNomTabla.Text, 6)), Trim(Right(cboMagnitud.Text, 2)), txtFechaNaci.Text, Trim(Right(cboTDC.Text, 4)), txtDocCom.Text, txtPersNom.Text, txtApellidoPaterno.Text, txtApellidoMaterno.Text, txtApellidoCasada.Text, txtNombre1.Text, txtNombre2.Text, txtCodSBS.Text, txtRelInst.Text, Trim(Right(cboSexo.Text, 2))) 'JUEZ 20150310 Se agregó Trim(cboNomTabla.Text)
    For i = 1 To FeCreditos.Rows - 1
        'Call oRCD.ActualizarVc02(gdFecData, txtNumSec.Text, FeCreditos.TextMatrix(i, 4), FeCreditos.TextMatrix(i, 5))
        Call oRCD.ActualizarVc02(gdFecData, txtNumSec.Text, Trim(Right(cboNomTabla.Text, 6)), FeCreditos.TextMatrix(i, 4), FeCreditos.TextMatrix(i, 5)) 'JUEZ 20150310 Se agregó Trim(cboNomTabla.Text)
    Next i
    MsgBox "Los datos se guardaron correctamente", vbApplicationModal
End Sub

Private Sub cmdNuevo_Click()
    Set oRs01 = Nothing
    Set oRs02 = Nothing
    txtNumSec.Enabled = True
    LimpiaFlex FeCreditos
    txtFechaNaci.Text = "__/__/____"
    cboTDC.ListIndex = IndiceListaCombo(cboTDC, -1)
    cboMagnitud.ListIndex = IndiceListaCombo(cboMagnitud, -1)
    cboAgencia.ListIndex = IndiceListaCombo(cboAgencia, -1)
    txtDocCom.Text = ""
    txtPersNom.Text = ""
    txtApellidoPaterno.Text = ""
    txtApellidoMaterno.Text = ""
    txtApellidoCasada.Text = ""
    txtNombre1.Text = ""
    txtNombre2.Text = ""
    txtRelInst.Text = "" 'ALPA 20120615
    cboSexo.ListIndex = IndiceListaCombo(cboSexo, -1) 'JUEZ 20140410
    cmdModificar.Enabled = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub ObtenerVC()
Dim oRCD As COMNCredito.NCOMRCD
Set oRCD = New COMNCredito.NCOMRCD
Set oRs01 = New ADODB.Recordset
Set oRs02 = New ADODB.Recordset

'Set oRs01 = oRCD.ObtenerVC(gdFecData, txtNumSec.Text, "01")
'Set oRs02 = oRCD.ObtenerVC(gdFecData, txtNumSec.Text, "02")
Set oRs01 = oRCD.ObtenerVC(gdFecData, txtNumSec.Text, Trim(Right(cboNomTabla.Text, 6)), "01") 'JUEZ 20150310 Se agregó Trim(cboNomTabla.Text)
Set oRs02 = oRCD.ObtenerVC(gdFecData, txtNumSec.Text, Trim(Right(cboNomTabla.Text, 6)), "02") 'JUEZ 20150310 Se agregó Trim(cboNomTabla.Text)
Call MostrarVc01
Call MostrarFlexCreditos
End Sub
Private Sub MostrarVc01()
If Not (oRs01.BOF Or oRs01.EOF) Then
    txtNumSec.Enabled = False
    cboMagnitud.ListIndex = IndiceListaCombo(cboMagnitud, oRs01!cMagEmp)
    txtFechaNaci.Text = oRs01!dFecNac
    cboTDC.ListIndex = IndiceListaCombo(cboTDC, oRs01!cTiDociComp)
    txtDocCom.Text = Trim(oRs01!cNuDociComp)
    txtPersNom.Text = oRs01!cPersNom
    txtApellidoPaterno.Text = oRs01!cApePat
    txtApellidoMaterno.Text = oRs01!capemat
    txtApellidoCasada.Text = oRs01!cApeCasada
    txtNombre1.Text = oRs01!cNombre1
    txtNombre2.Text = oRs01!cNombre2
    txtCodSBS.Text = oRs01!ccodsbs
    txtRelInst.Text = IIf(IsNull(oRs01!cRelInst), 0, oRs01!cRelInst) 'ALPA 20120615
    cboSexo.ListIndex = IndiceListaCombo(cboSexo, oRs01!cPersGenero) 'JUEZ 20140410
    cmdModificar.Enabled = True
Else
    MsgBox "Registro no encontrado", vbInformation, "Aviso"
    txtNumSec.Enabled = True
    txtNumSec.SetFocus
End If
End Sub

Private Sub Form_Load()
    
    Call RecorrerConstante1004
    Call RecorrerConstante1015
    Call RecorrerAgencia
    Call RecorrerUbigeo
    txtFechaCierre.Text = gdFecData
    cmdModificar.Enabled = False
End Sub

Private Sub RecorrerConstante1004()
    Dim oRs As ADODB.Recordset
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    
    Set oRs = oCons.RecuperaConstantes(1004)
    
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
            cboMagnitud.AddItem UCase(oRs!cConsDescripcion) & Space(250) & oRs!nConsValor
            oRs.MoveNext
        Loop
    End If
    
End Sub

Private Sub RecorrerConstante1015()
    Dim oRs As ADODB.Recordset
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    
    cboTDC.AddItem "SIN DOCUMENTO COMPLEMENTARIO" & Space(250) & "  "
    
    Set oRs = oCons.RecuperaConstantes(1015)
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
          cboTDC.AddItem oRs!cConsDescripcion & Space(250) & IIf(Len(oRs!nConsValor) = 1, "0", "") & CStr(oRs!nConsValor)
         oRs.MoveNext
        Loop
    End If
End Sub
Private Sub RecorrerAgencia()
    Dim oRs As ADODB.Recordset
    Dim oCons As COMDConstantes.DCOMAgencias
    Set oCons = New COMDConstantes.DCOMAgencias
    
    Set oRs = oCons.ObtieneAgencias
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
          cboAgencia.AddItem oRs!cConsDescripcion & Space(250) & IIf(Len(oRs!nConsValor) = 1, "0", "") & CStr(oRs!nConsValor)
         oRs.MoveNext
        Loop
    End If
End Sub
Private Sub RecorrerUbigeo()
    Dim oRs As ADODB.Recordset
    Dim oCons As COMNCredito.NCOMRCD
    Set oCons = New COMNCredito.NCOMRCD
    
    Set oRs = oCons.ObtenerUbigeoRCD
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
          cboUbigeo.AddItem oRs!cDesc & Space(250) & CStr(oRs!cCodigo)
         oRs.MoveNext
        Loop
    End If
End Sub


Private Sub MostrarFlexCreditos()
    Dim lsAgencia As String
    Dim lsUbigeo As String

    Dim i As Integer
    LimpiaFlex FeCreditos
    i = 1
    Do While Not oRs02.EOF
        FeCreditos.AdicionaFila
        FeCreditos.TextMatrix(oRs02.Bookmark, 1) = oRs02!cCtaCod
        FeCreditos.TextMatrix(oRs02.Bookmark, 2) = oRs02!cCtaCnt
        FeCreditos.TextMatrix(oRs02.Bookmark, 3) = oRs02!nSaldo
        FeCreditos.TextMatrix(oRs02.Bookmark, 4) = oRs02!cCodAge
        FeCreditos.TextMatrix(oRs02.Bookmark, 5) = oRs02!cUbicGeo
        lsAgencia = Right(oRs02!cCodAge, 2)
        lsUbigeo = oRs02!cUbicGeo
        oRs02.MoveNext
    Loop
    cboAgencia.ListIndex = IndiceListaCombo(cboAgencia, lsAgencia)
    cboUbigeo.ListIndex = IndiceListaCombo(cboUbigeo, lsUbigeo)
End Sub

Private Sub txtNumSec_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub txtNumSec_LostFocus()
    If Len(Trim(txtNumSec.Text)) < 8 Then
        txtNumSec.Text = String(8 - Len(Trim(txtNumSec.Text)), "0") & Trim(txtNumSec.Text)
    End If
End Sub
