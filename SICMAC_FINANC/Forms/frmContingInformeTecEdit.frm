VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingInformeTecEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Informe Tecnico"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   Icon            =   "frmContingInformeTecEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   1320
      TabIndex        =   31
      Top             =   4080
      Width           =   1050
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   345
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTabConting 
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6853
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Informe Técnico"
      TabPicture(0)   =   "frmContingInformeTecEdit.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraActivo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraPasivo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame FraPasivo 
         Caption         =   " Datos del Informe Técnico "
         Height          =   2055
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   10095
         Begin VB.TextBox txtNombreRefP 
            Height          =   285
            Left            =   1080
            TabIndex        =   37
            Top             =   1520
            Width           =   2775
         End
         Begin VB.TextBox txtEntidadP 
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   7080
            TabIndex        =   36
            Top             =   1125
            Width           =   2835
         End
         Begin VB.TextBox txtDocOrigenP 
            Height          =   285
            Left            =   1080
            TabIndex        =   35
            Top             =   1120
            Width           =   2415
         End
         Begin VB.TextBox txtNroInformeP 
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Top             =   360
            Width           =   1935
         End
         Begin Sicmact.TxtBuscar txtBuscaPersP 
            Height          =   330
            Left            =   1080
            TabIndex        =   38
            Top             =   725
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Enabled         =   0   'False
            TipoBusqueda    =   7
            TipoBusPers     =   1
            EnabledText     =   0   'False
         End
         Begin MSMask.MaskEdBox txtFecInformeP 
            Height          =   285
            Left            =   4440
            TabIndex        =   39
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecOrigenP 
            Height          =   300
            Left            =   4800
            TabIndex        =   40
            Top             =   1125
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblMonedaDemP 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   8040
            TabIndex        =   60
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblDemandaP 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   8560
            TabIndex        =   57
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label lblMontoP 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   8560
            TabIndex        =   56
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label lblMonedaP 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   8040
            TabIndex        =   55
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblTipoP 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   4800
            TabIndex        =   54
            Top             =   1515
            Width           =   1515
         End
         Begin VB.Label lblCalifP 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   7800
            TabIndex        =   53
            Top             =   1515
            Width           =   2115
         End
         Begin VB.Label Label25 
            Caption         =   "Nombre Referencial : "
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "Entidad : "
            Height          =   255
            Left            =   6360
            TabIndex        =   51
            Top             =   1170
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Doc. Origen : "
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Fecha Doc. :"
            Height          =   255
            Left            =   3720
            TabIndex        =   49
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Monto Real Pérdida (gastos) :"
            Height          =   255
            Left            =   5880
            TabIndex        =   48
            Top             =   795
            Width           =   2175
         End
         Begin VB.Label lblNombrePersP 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   47
            Top             =   725
            Width           =   2840
         End
         Begin VB.Label Label31 
            Caption         =   "Fecha Informe :"
            Height          =   255
            Left            =   3240
            TabIndex        =   46
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label32 
            Caption         =   "Personal a Cargo : "
            Height          =   480
            Left            =   120
            TabIndex        =   45
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label33 
            Caption         =   "Nº Informe :"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label34 
            Caption         =   "Calificación :"
            Height          =   255
            Left            =   6720
            TabIndex        =   43
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblLabelP 
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   4080
            TabIndex        =   42
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label35 
            Caption         =   "Monto Demanda (control) :"
            Height          =   255
            Left            =   5880
            TabIndex        =   41
            Top             =   435
            Width           =   2175
         End
      End
      Begin VB.Frame FraActivo 
         Caption         =   " Datos del Informe Técnico "
         Height          =   2055
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   10095
         Begin VB.TextBox txtNroInformeA 
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtDocOrigenA 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            Top             =   1120
            Width           =   2775
         End
         Begin VB.TextBox txtEntidadA 
            Height          =   285
            Left            =   7680
            TabIndex        =   14
            Top             =   1125
            Width           =   2115
         End
         Begin VB.TextBox txtNombreRefA 
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   1520
            Width           =   5175
         End
         Begin Sicmact.TxtBuscar TxtBuscarPersA 
            Height          =   330
            Left            =   1080
            TabIndex        =   17
            Top             =   725
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Enabled         =   0   'False
            TipoBusqueda    =   7
            TipoBusPers     =   1
            EnabledText     =   0   'False
         End
         Begin MSMask.MaskEdBox txtFecInformeA 
            Height          =   285
            Left            =   5040
            TabIndex        =   18
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecOrigenA 
            Height          =   300
            Left            =   5040
            TabIndex        =   19
            Top             =   1125
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblMonedaA 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   7680
            TabIndex        =   59
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblMontoA 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   315
            Left            =   8460
            TabIndex        =   58
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label lblCalifA 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   7680
            TabIndex        =   32
            Top             =   360
            Width           =   2115
         End
         Begin VB.Label Label12 
            Caption         =   "Calificación :"
            Height          =   255
            Left            =   6720
            TabIndex        =   29
            Top             =   405
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Nº Informe :"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Personal a Cargo : "
            Height          =   480
            Left            =   120
            TabIndex        =   27
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha Informe :"
            Height          =   255
            Left            =   3840
            TabIndex        =   26
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label lblNombrePersA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   25
            Top             =   720
            Width           =   3435
         End
         Begin VB.Label lblLabelA 
            Caption         =   "Monto :"
            Height          =   255
            Left            =   6720
            TabIndex        =   24
            Top             =   795
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Doc. :"
            Height          =   255
            Left            =   4080
            TabIndex        =   23
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Doc. Origen : "
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Entidad : "
            Height          =   255
            Left            =   6720
            TabIndex        =   21
            Top             =   1170
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre Referencial : "
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   " Datos Generales "
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10095
         Begin VB.Label txtDesc 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   615
            Left            =   6840
            TabIndex        =   11
            Top             =   360
            Width           =   2955
         End
         Begin VB.Label lblUsuario 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5520
            TabIndex        =   10
            Top             =   675
            Width           =   1155
         End
         Begin VB.Label lblFechaReg 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   5520
            TabIndex        =   9
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   1080
            TabIndex        =   8
            Top             =   680
            Width           =   390
         End
         Begin VB.Label lblMontoReg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   1520
            TabIndex        =   7
            Top             =   680
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "Usuario : "
            Height          =   255
            Left            =   4200
            TabIndex        =   6
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Origen :"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   395
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Monto : "
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Registro :"
            Height          =   255
            Left            =   4200
            TabIndex        =   3
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label lblOrigen 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00808080&
            Height          =   270
            Left            =   1080
            TabIndex        =   2
            Top             =   360
            Width           =   2595
         End
      End
   End
End
Attribute VB_Name = "frmContingInformeTecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingInformeTecEdit
'** Descripción : Form para ditar Informes Tecnicos para Contingencias creado segun RFC056-2012
'** Creación : JUEZ, 20120709 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String
Dim nItem As Integer

Public Function EditarActivo(ByVal pnNumItem As Integer, ByVal psNumRegistro As String)
    nItem = pnNumItem
    sNumRegistro = psNumRegistro
    Dim rsConting As ADODB.Recordset
    Set oConting = New DContingencia
    Set rsConting = oConting.BuscaContigenciaSeleccionada(psNumRegistro)
    
    lblOrigen.Caption = rsConting!cOrigen
    lblUsuario.Caption = rsConting!cUserReg
    lblFechaReg.Caption = Format(rsConting!dFechaReg, "dd/mm/yyyy")
    lblMoneda.Caption = rsConting!cmoneda
    lblMontoReg.Caption = Format(rsConting!nMonto, "#,##0.00")
    txtDesc.Caption = " " & Trim(rsConting!cContigDesc)
    Set rsConting = Nothing
    
    Set rs = oConting.RecuperaInfTecnicoxItem(pnNumItem, psNumRegistro)
    
    txtNroInformeA.Text = rs!cNroInforme
    txtFecInformeA.Text = Format(rs!dFechaRegInfTec, "dd/mm/yyyy")
    lblCalifA.Caption = rs!cCalif
    TxtBuscarPersA.Text = rs!cPersCod
    lblMonedaA.Caption = rs!cmoneda
    lblMontoA.Caption = Format(rs!nProvision, "#,##0.00")
    txtDocOrigenA.Text = rs!cDocOrigen
    txtFecOrigenA.Text = Format(rs!dFechaDocOrigen, "dd/mm/yyyy")
    txtEntidadA.Text = rs!cEntidad
    txtNombreRefA.Text = rs!cNombreRef
    
    FraActivo.Visible = True
    FraPasivo.Visible = False
    
    TxtBuscarPersA.Enabled = True
    txtBuscarPersA_EmiteDatos
    
    Me.Caption = "Activos Contingentes: Editar Informe Tecnico"
    Me.Show 1
End Function

Public Function EditarPasivo(ByVal pnNumItem As Integer, ByVal psNumRegistro As String)
    nItem = pnNumItem
    sNumRegistro = psNumRegistro
    Dim rsConting As ADODB.Recordset
    Set oConting = New DContingencia
    Set rsConting = oConting.BuscaContigenciaSeleccionada(psNumRegistro)
    
    lblOrigen.Caption = rsConting!cOrigen
    lblUsuario.Caption = rsConting!cUserReg
    lblFechaReg.Caption = Format(rsConting!dFechaReg, "dd/mm/yyyy")
    lblMoneda.Caption = rsConting!cmoneda
    lblMontoReg.Caption = Format(rsConting!nMonto, "#,##0.00")
    txtDesc.Caption = " " & rsConting!cContigDesc
    Set rsConting = Nothing
    
    Set rs = oConting.RecuperaInfTecnicoxItem(pnNumItem, psNumRegistro)
    
    txtNroInformeP.Text = rs!cNroInforme
    txtFecInformeP.Text = Format(rs!dFechaRegInfTec, "dd/mm/yyyy")
    lblMonedaP.Caption = rs!cmoneda
    lblMontoP.Caption = Format(rs!nProvision, "#,##0.00")
    txtBuscaPersP.Text = rs!cPersCod
    lblMonedaDemP.Caption = rs!cMonedaDem
    lblDemandaP.Caption = Format(rs!nDemandaLab, "#,##0.00")
    txtDocOrigenP.Text = rs!cDocOrigen
    txtFecOrigenP.Text = rs!dFechaDocOrigen
    txtEntidadA.Text = rs!cEntidad
    txtNombreRefA.Text = rs!cNombreRef
    lblTipoP.Caption = rs!cTipo
    lblCalifP.Caption = rs!cCalif
    
    FraActivo.Visible = False
    FraPasivo.Visible = True
    
    txtBuscaPersP.Enabled = True
    txtBuscaPersP_EmiteDatos
    
    Me.Caption = "Pasivos Contingentes: Editar Informe Tecnico"
    Me.Show 1
End Function

Private Sub cmdActualizar_Click()
    If ValidaDatos Then
        If MsgBox("Está seguro de actualizar los datos del Informe Técnico? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
        Set oConting = New DContingencia
        Dim cUser As String
        cUser = oConting.ObtenerUserxCodPersEncargado(TxtBuscarPersA.Text)
        If cUser <> "" Then
            If FraActivo.Enabled = True Then
                Call oConting.ActualizaInformeTecnico(nItem, sNumRegistro, Trim(txtNroInformeA.Text), Trim(txtFecInformeA.Text), _
                cUser, Trim(txtDocOrigenA.Text), Trim(txtFecOrigenA.Text), Trim(txtEntidadA.Text), Trim(txtNombreRefA.Text))
            Else
                Call oConting.ActualizaInformeTecnico(nItem, sNumRegistro, Trim(txtNroInformeP.Text), Trim(txtFecInformeP.Text), _
                cUser, Trim(txtDocOrigenP.Text), Trim(txtFecOrigenP.Text), Trim(txtEntidadP.Text), Trim(txtNombreRefP.Text))
            End If
            MsgBox "Se ha actuliazado con exito los datos del Informe Técnico", vbInformation, "Aviso"
            Call cmdCancelar_Click
            frmContingInformeTecCons.Extorno sNumRegistro
        Else
            MsgBox "La persona elegida no puede ser la encargada", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub txtBuscarPersA_EmiteDatos()
    If TxtBuscarPersA.Text = "" Then Exit Sub
    
    Dim rsDatos As New ADODB.Recordset
    Set oGen = New DGeneral
    Set rsDatos = oGen.GetDataUser(TxtBuscarPersA.Text, True)
    If rsDatos.RecordCount = 0 Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
        TxtBuscarPersA.Text = ""
    Else
        lblNombrePersA.Caption = rsDatos!cPersNombre
    End If
End Sub

Private Sub txtBuscaPersP_EmiteDatos()
    If txtBuscaPersP.Text = "" Then Exit Sub
    
    Dim rsDatos As New ADODB.Recordset
    Set oGen = New DGeneral
    Set rsDatos = oGen.GetDataUser(txtBuscaPersP.Text, True)
    If rsDatos.RecordCount = 0 Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
        txtBuscaPersP.Text = ""
    Else
        lblNombrePersP.Caption = rsDatos!cPersNombre
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub txtDocOrigenA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecOrigenA.SetFocus
    End If
End Sub

Private Sub txtDocOrigenP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecOrigenP.SetFocus
    End If
End Sub

Private Sub txtEntidadA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNombreRefA.SetFocus
    End If
End Sub

Private Sub txtEntidadP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNombreRefP.SetFocus
    End If
End Sub

Private Sub txtFecInformeA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBuscarPersA.SetFocus
    End If
End Sub

Private Sub txtFecInformeA_LostFocus()
    Dim sMsj As String
    sMsj = ValidaFecha(txtFecInformeA.Text)
    If Not Trim(sMsj) = "" Then
        MsgBox sMsj, vbInformation, "Aviso"
        If txtFecInformeA.Enabled Then
            txtFecInformeA.SetFocus
            Exit Sub
        End If
    End If
    If CDate(txtFecInformeA.Text) >= gdFecSis Then
        MsgBox "Fecha del informe No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecInformeA.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFecInformeP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBuscaPersP.SetFocus
    End If
End Sub

Private Sub txtFecInformeP_LostFocus()
    Dim sMsj As String
    sMsj = ValidaFecha(txtFecInformeP.Text)
    If Not Trim(sMsj) = "" Then
        MsgBox sMsj, vbInformation, "Aviso"
        If txtFecInformeP.Enabled Then
            txtFecInformeP.SetFocus
            Exit Sub
        End If
    End If
    If CDate(txtFecInformeP.Text) >= gdFecSis Then
        MsgBox "Fecha del informe No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecInformeP.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFecOrigenA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntidadA.SetFocus
    End If
End Sub

Private Sub txtFecOrigenA_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFecOrigenA.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        If txtFecOrigenA.Enabled Then
            txtFecOrigenA.SetFocus
            Exit Sub
        End If
    End If
    If CDate(txtFecOrigenA.Text) >= gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecOrigenA.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFecOrigenP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntidadP.SetFocus
    End If
End Sub

Private Sub txtFecOrigenP_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFecOrigenP.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        If txtFecOrigenP.Enabled Then txtFecOrigenP.SetFocus
        Exit Sub
    End If
    If CDate(txtFecOrigenP.Text) >= gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecOrigenP.SetFocus
        Exit Sub
    End If
End Sub

Public Function ValidaDatos() As Boolean
    If FraActivo.Enabled = True Then
        If Trim(txtNroInformeA.Text) = "" Then
            MsgBox "Falta ingresar el Nro. de informe", vbInformation, "Aviso"
            txtNroInformeA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecInformeA.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha del informe", vbInformation, "Aviso"
            txtFecInformeA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(TxtBuscarPersA.Text) = "" Then
            MsgBox "Falta ingresar la persona encargada", vbInformation, "Aviso"
            TxtBuscarPersA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtDocOrigenA.Text) = "" Then
            MsgBox "Falta ingresar el documento de origen", vbInformation, "Aviso"
            txtDocOrigenA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecOrigenA.Text) = "" Then
            MsgBox "Falta ingresar la fecha del documento de origen", vbInformation, "Aviso"
            txtFecOrigenA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtEntidadA.Text) = "" Then
            MsgBox "Falta ingresar la entidad", vbInformation, "Aviso"
            txtEntidadA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtNombreRefA.Text) = "" Then
            MsgBox "Falta ingresar el normbre referencial", vbInformation, "Aviso"
            txtNombreRefA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    Else
        If Trim(txtNroInformeP.Text) = "" Then
            MsgBox "Falta ingresar el Nro. de informe", vbInformation, "Aviso"
            txtNroInformeP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecInformeP.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha del informe", vbInformation, "Aviso"
            txtFecInformeP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtBuscaPersP.Text) = "" Then
            MsgBox "Falta ingresar la persona encargada", vbInformation, "Aviso"
            txtBuscaPersP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtDocOrigenP.Text) = "" Then
            MsgBox "Falta ingresar el documento de origen", vbInformation, "Aviso"
            txtDocOrigenP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecOrigenP.Text) = "" Then
            MsgBox "Falta ingresar la fecha del documento de origen", vbInformation, "Aviso"
            txtFecOrigenP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtEntidadP.Text) = "" Then
            MsgBox "Falta ingresar la entidad", vbInformation, "Aviso"
            txtEntidadP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtNombreRefP.Text) = "" Then
            MsgBox "Falta ingresar el normbre referencial", vbInformation, "Aviso"
            txtNombreRefP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    ValidaDatos = True
End Function

Private Sub txtNombreRefA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdActualizar.SetFocus
    End If
End Sub

Private Sub txtNombreRefP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdActualizar.SetFocus
    End If
End Sub

Private Sub txtNroInformeA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecInformeA.SetFocus
    End If
End Sub

Private Sub txtNroInformeP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecInformeP.SetFocus
    End If
End Sub
