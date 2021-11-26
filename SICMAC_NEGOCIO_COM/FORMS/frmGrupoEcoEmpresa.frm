VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGrupoEcoEmpresa 
   Caption         =   "Empresas - Grupo Economico"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8670
   Icon            =   "frmGrupoEcoEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin MSMask.MaskEdBox txtFecha 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CheckBox ckActivar 
      Caption         =   "Activar"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   7800
      TabIndex        =   1
      Top             =   4560
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo Economico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.Frame Frame2 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   8415
         Begin SICMACT.TxtBuscar txtCodEmpresa 
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2295
            _extentx        =   4048
            _extenty        =   661
            appearance      =   1
            appearance      =   1
            font            =   "frmGrupoEcoEmpresa.frx":030A
            appearance      =   1
            tipobusqueda    =   3
            stitulo         =   ""
         End
         Begin VB.Frame fraVinculados 
            Caption         =   "Vinculados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   0
            TabIndex        =   6
            Top             =   1560
            Width           =   8415
            Begin VB.CheckBox ckRepresentante 
               Alignment       =   1  'Right Justify
               Caption         =   "Representante Legal"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5400
               TabIndex        =   21
               Top             =   1200
               Width           =   2175
            End
            Begin SICMACT.TxtBuscar txtCodRepresentanteLegal 
               Height          =   375
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   2295
               _extentx        =   4048
               _extenty        =   661
               appearance      =   1
               appearance      =   1
               font            =   "frmGrupoEcoEmpresa.frx":0336
               appearance      =   1
               tipobusqueda    =   3
               stitulo         =   ""
            End
            Begin VB.Label lblNroDocVinc 
               Height          =   255
               Left            =   1800
               TabIndex        =   20
               Top             =   1200
               Width           =   2175
            End
            Begin VB.Label Label4 
               Caption         =   "Nro. Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   19
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label lblDireccionRL 
               Height          =   255
               Left            =   1800
               TabIndex        =   16
               Top             =   840
               Width           =   6375
            End
            Begin VB.Label Label2 
               Caption         =   "Dirección"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblRepresentanteLegal 
               Height          =   255
               Left            =   2760
               TabIndex        =   12
               Top             =   480
               Width           =   5295
            End
         End
         Begin VB.Label lblDocEmpresa 
            Height          =   255
            Left            =   1920
            TabIndex        =   18
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblDireccion 
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   840
            Width           =   6255
         End
         Begin VB.Label Label1 
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblEmpresa 
            Height          =   255
            Left            =   2760
            TabIndex        =   8
            Top             =   480
            Width           =   5415
         End
      End
      Begin VB.Label lblGrupo 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmGrupoEcoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pnGrupo As Integer
Dim pcGrupo As String
Dim pnTipo As Integer
Public nLogPerNegativa As Integer

Private Sub cmdGrabar_Click()
Dim nRetorno As Integer
Dim oGrup As COMNPersona.ncompersona
Set oGrup = New COMNPersona.ncompersona
'nRetorno = oGrup.GrabarPersGrupoEconomico(pnGrupo, txtCodEmpresa.Text, txtCodRepresentanteLegal.Text, IIf(ckRepresentante.value = 1, 1, 0), 0, 0, 0, 0, "", IIf(pnTipo = 1, 1, ckActivar.value), "", "", 0, txtFecha.Text)
'MsgBox "Datos se registraron correctamente", vbInformation
'Call cmdsalir_Click'JAME20140227 COMENTO
'JAME20140227 ERS167-2013****************************************************
    If Len(Trim(txtCodEmpresa.Text)) = 0 Or Len(Trim(txtCodRepresentanteLegal.Text)) = 0 Then
        MsgBox " Usted tiene que ingresar el nombre ", vbInformation, "Aviso"
        Exit Sub
    End If
     
    If pnGrupo = 2 Then
        If txtCodEmpresa.Text = txtCodRepresentanteLegal.Text Then
            MsgBox "No se permite relacionar a una persona natural y\o juridica con ella misma", vbCritical, "Aviso"
            Exit Sub
        End If
        nRetorno = oGrup.GrabarPersGrupoEconomico(pnGrupo, txtCodEmpresa.Text, txtCodRepresentanteLegal.Text, IIf(ckRepresentante.value = 1, 1, 0), 0, 0, 0, 0, "", IIf(pnTipo = 1, 1, ckActivar.value), "", "", 0, txtFecha.Text)
        MsgBox "Datos se registraron correctamente", vbInformation
        Call cmdsalir_Click
        Exit Sub
    End If
   
    If txtCodEmpresa.Text = txtCodRepresentanteLegal.Text Then
        MsgBox "No se permite relacionar a una persona natural y\o juridica con ella misma", vbCritical, "Aviso"
    Else
        If oGrup.validarDevuelveGrupoEconomicoPersona(txtCodEmpresa.Text) = True Then
            MsgBox "No se permite grabar a la empresa jurìdica por que ya existe en Sin Grupo", vbCritical, "Aviso"
        Else
            nRetorno = oGrup.GrabarPersGrupoEconomico(pnGrupo, txtCodEmpresa.Text, txtCodRepresentanteLegal.Text, IIf(ckRepresentante.value = 1, 1, 0), 0, 0, 0, 0, "", IIf(pnTipo = 1, 1, ckActivar.value), "", "", 0, txtFecha.Text)
            MsgBox "Datos se registraron correctamente", vbInformation
            Call cmdsalir_Click
        End If
    End If
'JAME FIN
End Sub

Private Sub cmdModificar_Click()
Dim nRetorno As Integer
Dim oGrup As COMNPersona.ncompersona
Set oGrup = New COMNPersona.ncompersona
nRetorno = oGrup.GrabarPersGrupoEconomico(pnGrupo, txtCodEmpresa.Text, txtCodRepresentanteLegal.Text, IIf(ckRepresentante.value = 1, 1, 0), 0, 0, 0, 0, "", IIf(pnTipo = 1, 1, ckActivar.value), "", "", 0, txtFecha.Text)
MsgBox "Datos se registraron correctamente", vbInformation
Call cmdsalir_Click
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Public Sub Modificar(ByVal nGrupo As Integer, ByVal cGrupo As String, ByVal cPersCodEmpresa As String, ByVal cPersCodVinculado As String)
    Dim oComPe As COMNPersona.ncompersona
    Set oComPe = New COMNPersona.ncompersona
    
    Dim lsRepresentanteLegal As String
    Dim lsDireccionRL As String
    Dim lsNroDocVinc As String
    
    Dim lsEmpresa As String
    Dim lsDireccion As String
    Dim lsDocEmpresa As String
    
    pnGrupo = nGrupo
    pcGrupo = cGrupo
    lblGrupo.Caption = pcGrupo
    txtCodEmpresa.Text = cPersCodEmpresa
    If oComPe.ValidaPersona(Trim(cPersCodEmpresa), lsEmpresa, lsDireccion, lsDocEmpresa) = False Then
        MsgBox "No existe datos del Cliente"
    Else
        lblEmpresa.Caption = lsEmpresa
        lblDireccion.Caption = lsDireccion
        lblDocEmpresa.Caption = lsDocEmpresa
    End If

    txtCodRepresentanteLegal.Text = cPersCodVinculado
    If oComPe.ValidaPersona(Trim(cPersCodVinculado), lsRepresentanteLegal, lsDireccionRL, lsNroDocVinc) = False Then
        MsgBox "No existe datos del Cliente"
    Else
        lblRepresentanteLegal.Caption = lsRepresentanteLegal
        lblDireccionRL.Caption = lsDireccionRL
        lblNroDocVinc.Caption = lsNroDocVinc
    End If
    If Len(Trim(lblDocEmpresa.Caption)) = 11 And Mid(lblDocEmpresa.Caption, 1, 1) = "2" Then
        ckRepresentante.Enabled = True
    Else
        ckRepresentante.Enabled = False
    End If
    pnTipo = 2
    Call validar_controles(True, pnTipo)
    txtFecha.Text = gdFecSis
    Show 1
End Sub

Public Sub Nuevo(ByVal nGrupo As Integer, ByVal cGrupo As String)
    pnGrupo = nGrupo
    pcGrupo = cGrupo
    lblGrupo.Caption = pcGrupo
    pnTipo = 1
    ckActivar.Visible = False
    Call validar_controles(True, 1)
    txtFecha.Text = gdFecSis
    Show 1
End Sub

Private Sub validar_controles(ByVal bHabilitar As Boolean, nTipo As Integer)
    If nTipo = 1 Then
        cmdGrabar.Enabled = bHabilitar
        cmdModificar.Enabled = Not bHabilitar
        cmdEliminar.Enabled = Not bHabilitar
    ElseIf nTipo = 2 Then
        cmdGrabar.Enabled = Not bHabilitar
        cmdModificar.Enabled = bHabilitar
        cmdEliminar.Enabled = Not bHabilitar
    ElseIf nTipo = 3 Then
        cmdGrabar.Enabled = Not bHabilitar
        cmdModificar.Enabled = Not bHabilitar
        cmdEliminar.Enabled = bHabilitar
    End If
    
    cmdCancelar.Enabled = bHabilitar
    cmdSalir.Enabled = bHabilitar
End Sub

Private Sub Form_Load()
    nLogPerNegativa = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    nLogPerNegativa = 0
End Sub

Private Sub txtCodEmpresa_EmiteDatos()
    lblEmpresa.Caption = txtCodEmpresa.psDescripcion
    lblDireccion.Caption = txtCodEmpresa.sPersDireccion
    lblDocEmpresa.Caption = txtCodEmpresa.sPersNroDoc
    If Len(Trim(txtCodEmpresa.sPersNroDoc)) = 11 And Mid(txtCodEmpresa.sPersNroDoc, 1, 1) = "2" Then
        ckRepresentante.Enabled = True
    Else
        ckRepresentante.Enabled = False
    End If
    
End Sub

Private Sub txtCodRepresentanteLegal_EmiteDatos()
    lblRepresentanteLegal.Caption = txtCodRepresentanteLegal.psDescripcion
    lblDireccionRL.Caption = txtCodRepresentanteLegal.sPersDireccion
    lblNroDocVinc.Caption = txtCodRepresentanteLegal.sPersNroDoc
End Sub
