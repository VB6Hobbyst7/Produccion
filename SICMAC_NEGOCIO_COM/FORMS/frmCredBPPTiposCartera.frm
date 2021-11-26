VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPTiposCartera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Tipos de Cartera"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frmCredBPPTiposCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Cartera"
      TabPicture(0)   =   "frmCredBPPTiposCartera.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "feTipoCartera"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdQuitar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEditar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCartera"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.TextBox txtCartera 
         Height          =   300
         Left            =   960
         TabIndex        =   14
         Top             =   600
         Width           =   4155
      End
      Begin VB.Frame Frame1 
         Caption         =   " Tipos de Créditos "
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   9255
         Begin VB.CheckBox ChkCredCorpo 
            Caption         =   "Créditos Corporativos"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1995
         End
         Begin VB.CheckBox ChkPequenaEmp 
            Caption         =   "Créditos a Pequeñas Empresas"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   2595
         End
         Begin VB.CheckBox ChkConsumoNoRev 
            Caption         =   "Créditos de Consumo no Revolvente"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   2955
         End
         Begin VB.CheckBox ChkHipotecario 
            Caption         =   "Créditos Hipotecario para Vivienda"
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   1080
            Width           =   2955
         End
         Begin VB.CheckBox ChkGrandeEmp 
            Caption         =   "Créditos a Grandes Empresas"
            Height          =   255
            Left            =   3240
            TabIndex        =   9
            Top             =   360
            Width           =   2475
         End
         Begin VB.CheckBox ChkMicro 
            Caption         =   "Créditos a Microempresas"
            Height          =   255
            Left            =   3240
            TabIndex        =   8
            Top             =   720
            Width           =   2595
         End
         Begin VB.CheckBox ChkMedianaEmp 
            Caption         =   "Créditos a Medianas Empresas"
            Height          =   255
            Left            =   6240
            TabIndex        =   7
            Top             =   360
            Width           =   2475
         End
         Begin VB.CheckBox ChkConsumoRev 
            Caption         =   "Créditos de Consumo Revolvente"
            Height          =   255
            Left            =   6240
            TabIndex        =   6
            Top             =   720
            Width           =   2715
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6240
            TabIndex        =   5
            Top             =   1080
            Width           =   1170
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   5280
         Width           =   1050
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   2
         Top             =   5280
         Width           =   1050
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8160
         TabIndex        =   1
         Top             =   5280
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feTipoCartera 
         Height          =   2415
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   4260
         Cols0           =   11
         HighLight       =   1
         EncabezadosNombres=   "-Id-Tipo Cartera-Corporativos-Grandes Emp.-Medianas Emp.-Pequeñas Emp.-Microempresas-Consumo Rev.-Consumo No Rev.-Hipotecario"
         EncabezadosAnchos=   "300-0-2630-1200-1300-1350-1350-1400-1200-1500-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
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
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   5280
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Cartera :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   630
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCredBPPTiposCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCreBPPTiposCartera
'** Descripción : Formulario para la Administracion de los Tipos de Cartera
'**               creado segun RFC099-2012
'** Creación : JUEZ, 20121005 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim IdTipoEdita As String

Private Sub cmdCancelar_Click()
    Call ManejaControles(True, True, False)
    Call EditaControles(False)
    IdTipoEdita = ""
End Sub

Private Function ManejaControles(ByVal pbHabilita As Boolean, ByVal pbEditar As Boolean, ByVal pbCancela As Boolean)
    'txtCartera.Enabled = pbHabilita
    'ChkCredCorpo.Enabled = pbHabilita
    'ChkGrandeEmp.Enabled = pbHabilita
    'ChkMedianaEmp.Enabled = pbHabilita
    'ChkPequenaEmp.Enabled = pbHabilita
    'ChkMicro.Enabled = pbHabilita
    'ChkConsumoRev.Enabled = pbHabilita
    'ChkConsumoNoRev.Enabled = pbHabilita
    'ChkHipotecario.Enabled = pbHabilita
    cmdQuitar.Visible = pbEditar
    cmdEditar.Visible = pbEditar
    cmdCancelar.Visible = pbCancela
End Function

Private Function EditaControles(ByVal pbEdita As Boolean)
    If pbEdita Then
        Dim oDCredBPP As COMNCredito.NCOMBPPR
        Dim rsCartera As ADODB.Recordset
        Set oDCredBPP = New COMNCredito.NCOMBPPR

        IdTipoEdita = feTipoCartera.TextMatrix(feTipoCartera.Row, 1)
        Set rsCartera = oDCredBPP.RecuperaCredTiposCartera(IdTipoEdita)
        Set oDCredBPP = Nothing

        txtCartera.Text = rsCartera!cTipoCartera
        ChkCredCorpo.value = rsCartera!CredCorporativo
        ChkGrandeEmp.value = rsCartera!CredGrandeEmp
        ChkMedianaEmp.value = rsCartera!CredMedianaEmp
        ChkPequenaEmp.value = rsCartera!CredPequenaEmp
        ChkMicro.value = rsCartera!CredMicroempresa
        ChkConsumoRev.value = rsCartera!CredConsumoRev
        ChkConsumoNoRev.value = rsCartera!CredConsumoNoRev
        ChkHipotecario.value = rsCartera!CredHipotecario
        Set rsCartera = Nothing
    Else
        txtCartera.Text = ""
        ChkCredCorpo.value = 0
        ChkGrandeEmp.value = 0
        ChkMedianaEmp.value = 0
        ChkPequenaEmp.value = 0
        ChkMicro.value = 0
        ChkConsumoRev.value = 0
        ChkConsumoNoRev.value = 0
        ChkHipotecario.value = 0
        IdTipoEdita = ""
    End If
End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdEditar_Click()
    Call ManejaControles(False, False, True)
    Call EditaControles(True)
End Sub

Private Sub cmdGuardar_Click()
    If ValidaDatos Then
        Dim oCredBPP As COMNCredito.NCOMBPPR
        Set oCredBPP = New COMNCredito.NCOMBPPR

        If IdTipoEdita = "" Then
            Call oCredBPP.dInsertaCredTipoCartera(Trim(txtCartera.Text), ChkCredCorpo.value, ChkGrandeEmp.value, ChkMedianaEmp.value, _
                                                    ChkPequenaEmp.value, ChkMicro.value, ChkConsumoRev.value, ChkConsumoNoRev.value, _
                                                    ChkHipotecario.value)
            MsgBox "Los datos se registraron correctamente", vbInformation, "Aviso"
        Else
            Call oCredBPP.dActualizaCredTipoCartera(IdTipoEdita, Trim(txtCartera.Text), ChkCredCorpo.value, ChkGrandeEmp.value, ChkMedianaEmp.value, _
                                                    ChkPequenaEmp.value, ChkMicro.value, ChkConsumoRev.value, ChkConsumoNoRev.value, _
                                                    ChkHipotecario.value)
            MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
        End If

        Call ManejaControles(True, True, False)
        Call EditaControles(False)
        IdTipoEdita = ""
        Call CargaDatos
    End If
End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feTipoCartera.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        feTipoCartera.EliminaFila feTipoCartera.Row
        Dim oCredBPP As COMNCredito.NCOMBPPR
        Set oCredBPP = New COMNCredito.NCOMBPPR
        Call oCredBPP.dEliminaCredTipoCartera(feTipoCartera.TextMatrix(feTipoCartera.Row, 1))
        Call CargaDatos
    End If
End Sub

Private Sub Form_Load()
    IdTipoEdita = ""
    Call CargaDatos
End Sub

Private Sub CargaDatos()
    Dim oCredBPP As COMNCredito.NCOMBPPR
    Dim rsCartera As ADODB.Recordset
    Dim lnFila As Integer
    Set oCredBPP = New COMNCredito.NCOMBPPR

    Set rsCartera = oCredBPP.RecuperaCredTiposCartera
    Set oCredBPP = Nothing
    Call LimpiaFlex(feTipoCartera)
    Do While Not rsCartera.EOF
        feTipoCartera.AdicionaFila
        lnFila = feTipoCartera.Row
        feTipoCartera.TextMatrix(lnFila, 1) = rsCartera!cIdTipo
        feTipoCartera.TextMatrix(lnFila, 2) = rsCartera!cTipoCartera
        feTipoCartera.TextMatrix(lnFila, 3) = IIf(rsCartera!CredCorporativo = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 4) = IIf(rsCartera!CredGrandeEmp = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 5) = IIf(rsCartera!CredMedianaEmp = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 6) = IIf(rsCartera!CredPequenaEmp = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 7) = IIf(rsCartera!CredMicroempresa = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 8) = IIf(rsCartera!CredConsumoRev = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 9) = IIf(rsCartera!CredConsumoNoRev = 0, "", "X")
        feTipoCartera.TextMatrix(lnFila, 10) = IIf(rsCartera!CredHipotecario = 0, "", "X")
        rsCartera.MoveNext
    Loop
    rsCartera.Close
    Set rsCartera = Nothing
End Sub

Private Function ValidaDatos()
    ValidaDatos = False

    If feTipoCartera.lbEditarFlex Then
        If feTipoCartera.TextMatrix(feTipoCartera.Row, 2) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 3) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 4) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 5) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 6) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 7) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 8) = "" And _
            feTipoCartera.TextMatrix(feTipoCartera.Row, 9) = "" Then
                MsgBox "Debe seleccionar al menos un Tipo de Crédito", vbInformation, "Aviso"
                ValidaDatos = False
                Exit Function
        End If
    Else
        If Trim(txtCartera.Text = "") Then
             MsgBox "Debe ingresar la descripcion del Tipo de Cartera", vbInformation, "Aviso"
             txtCartera.SetFocus
             ValidaDatos = False
             Exit Function
        End If

        If (ChkCredCorpo.value = 0 And ChkGrandeEmp.value = 0 And ChkMedianaEmp.value = 0 And ChkPequenaEmp.value = 0 And _
            ChkMicro.value = 0 And ChkConsumoRev.value = 0 And ChkConsumoNoRev.value = 0 And ChkHipotecario.value = 0) Then
             MsgBox "Debe seleccionar al menos un Tipo de Crédito", vbInformation, "Aviso"
             ValidaDatos = False
             Exit Function
        End If
    End If

    ValidaDatos = True
End Function

Private Sub txtCartera_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub

Private Sub txtCartera_LostFocus()
    txtCartera.Text = UCase(txtCartera.Text)
End Sub


