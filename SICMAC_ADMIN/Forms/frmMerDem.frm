VERSION 5.00
Begin VB.Form frmMerDem 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmMerDem.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7485
      TabIndex        =   8
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Index           =   1
      Left            =   15
      TabIndex        =   7
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Index           =   1
      Left            =   1215
      TabIndex        =   6
      Top             =   4695
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Index           =   1
      Left            =   2415
      TabIndex        =   1
      Top             =   4695
      Width           =   1095
   End
   Begin SicmactAdmin.ctrRRHHGen ctrRRHHGen 
      Height          =   1200
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   8535
      _ExtentX        =   14261
      _ExtentY        =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraPerNoLab 
      Caption         =   "Permisos"
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
      Height          =   3405
      Index           =   1
      Left            =   15
      TabIndex        =   2
      Top             =   1215
      Width           =   8565
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   1
         Left            =   7410
         TabIndex        =   4
         Top             =   2955
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   3
         Top             =   2970
         Width           =   1095
      End
      Begin SicmactAdmin.FlexEdit FlexPerNoLab 
         Height          =   2685
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   8355
         _ExtentX        =   14790
         _ExtentY        =   4736
         Cols0           =   7
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod.Tpo-Tipo-Fecha-Observaciones-bit-bit1"
         EncabezadosAnchos=   "300-800-1800-1500-5000-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmMerDem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ctrRRHHGen_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    'oPersona.ObtieneClientexCodigo ()
    ClearScreen
    If Not oPersona Is Nothing Then
        If lnEstadosTipo = RHEstadosTpoVacaciones Then
            If Not oRRHH.GetRRHHControl(oPersona.sPersCod, TipoControl.TipoControlVacaciones) Then
                MsgBox "La Persona Pertenece a un grupo de empleados al que no se le controla las vacaciones.", vbInformation, "Aviso"
                Set oRRHH = Nothing
                Set oPersona = Nothing
                Exit Sub
            End If
        ElseIf lnEstadosTipo = RHEstadosTpoPermisosLicencias Then
            If Not oRRHH.GetRRHHControl(oPersona.sPersCod, TipoControl.TipoControlPermisos) Then
                MsgBox "La Persona Pertenece a un grupo de empleados al que no se le controla los permisos.", vbInformation, "Aviso"
                Set oRRHH = Nothing
                Set oPersona = Nothing
                Exit Sub
            End If
        End If
        
        ClearScreen
        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        CargaData Me.ctrRRHHGen.psCodigoPersona
    End If
End Sub

Private Sub CargaData(psPersCod As String)
    Dim oPNL As DPeriodoNoLaborado
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Set oPNL = New DPeriodoNoLaborado
    
    Set rsP = oPNL.GetRHPeriodoNoLabPersona(psPersCod, CInt(lnEstadosTipo))
    
    If Not (rsP.EOF And rsP.BOF) Then
        Set Me.FlexPerNoLab(lnIndiceActual).Recordset = rsP
    Else
        FlexPerNoLab(lnIndiceActual).Clear
        FlexPerNoLab(lnIndiceActual).Rows = 2
        FlexPerNoLab(lnIndiceActual).FormaCabecera
    End If
End Sub

