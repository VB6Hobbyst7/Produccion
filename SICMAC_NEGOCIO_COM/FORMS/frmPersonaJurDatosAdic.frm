VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonaJurDatosAdic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Accionariales"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11430
   Icon            =   "frmPersonaJurDatosAdic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   11430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   300
      Left            =   4920
      TabIndex        =   9
      Top             =   3960
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTabs 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Accionistas"
      TabPicture(0)   =   "frmPersonaJurDatosAdic.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fleAccionistas"
      Tab(0).Control(1)=   "cmdEliminar"
      Tab(0).Control(2)=   "cmdEditar"
      Tab(0).Control(3)=   "cmdNuevo"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Directorio"
      TabPicture(1)   =   "frmPersonaJurDatosAdic.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fleDirectorio"
      Tab(1).Control(1)=   "cmdNuevoG"
      Tab(1).Control(2)=   "cmdDEditar"
      Tab(1).Control(3)=   "cmdDEliminar"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Gerencias y Ejecutivos"
      TabPicture(2)   =   "frmPersonaJurDatosAdic.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fleGerencias"
      Tab(2).Control(1)=   "cmdGEliminar"
      Tab(2).Control(2)=   "cmdGEditar"
      Tab(2).Control(3)=   "cmdNuevoE"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Patrimonio en Otras Entidades"
      TabPicture(3)   =   "frmPersonaJurDatosAdic.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "flePatrimonio"
      Tab(3).Control(1)=   "cmdPEliminar"
      Tab(3).Control(2)=   "cmdPEditar"
      Tab(3).Control(3)=   "cmdPatrimonio1"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Participación Patrimonial en Empresas"
      TabPicture(4)   =   "frmPersonaJurDatosAdic.frx":037A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "flePatOtrasEmpresa"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdPEEliminar"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdPEEditar"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdPatrimonio2"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Cargos en Otras Empresas"
      TabPicture(5)   =   "frmPersonaJurDatosAdic.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fleCargos"
      Tab(5).Control(1)=   "cmdNuevoC"
      Tab(5).Control(2)=   "cmdEditarC"
      Tab(5).Control(3)=   "cmdEliminarC"
      Tab(5).ControlCount=   4
      Begin VB.CommandButton cmdEliminarC 
         Cancel          =   -1  'True
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   26
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdEditarC 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   25
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdNuevoC 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   -65160
         TabIndex        =   24
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmdPatrimonio2 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   9840
         TabIndex        =   21
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmdPEEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   9840
         TabIndex        =   20
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdPEEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   9840
         TabIndex        =   19
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdPatrimonio1 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   -65160
         TabIndex        =   18
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmdPEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   17
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdPEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   16
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdNuevoE 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   -65160
         TabIndex        =   15
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmdGEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   14
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdGEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   13
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdDEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   12
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdDEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   11
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdNuevoG 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   -65160
         TabIndex        =   10
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   -65160
         TabIndex        =   8
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   7
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   -65160
         TabIndex        =   6
         Top             =   1200
         Width           =   1000
      End
      Begin SICMACT.FlexEdit fleAccionistas 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5530
         Cols0           =   8
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nombre o Razón Social-RUC-D.O.I.-Nacionalidad-Aporte S/-%-"
         EncabezadosAnchos=   "350-3100-1200-1200-1500-1200-600-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-6-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C-L-R-R-L"
         FormatosEdit    =   "0-1-3-0-1-2-3-0"
         CantEntero      =   15
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fleDirectorio 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   9615
         _ExtentX        =   17595
         _ExtentY        =   5530
         Cols0           =   7
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nombre o Razón Social-RUC-D.O.I.-Cargo-Nacionalidad-"
         EncabezadosAnchos=   "350-3000-1200-1200-2000-1500-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C-L-L-C"
         FormatosEdit    =   "0-1-3-0-1-1-0"
         CantEntero      =   10
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fleGerencias 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   9615
         _ExtentX        =   15478
         _ExtentY        =   5530
         Cols0           =   6
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nombre o Razón Social-RUC-D.O.I.-Cargo y Función-"
         EncabezadosAnchos=   "350-3000-1200-1200-2000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C-L-C"
         FormatosEdit    =   "0-1-3-0-1-0"
         CantEntero      =   10
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flePatrimonio 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   9615
         _ExtentX        =   15478
         _ExtentY        =   5530
         Cols0           =   9
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#--Cod. SBS-Apellidos y Nombres-Razón Social-RUC-Aporte S/-%-"
         EncabezadosAnchos=   "350-0-1200-3000-1800-1200-1200-600-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-X"
         ListaControles  =   "0-0-1-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R-R-R-L"
         FormatosEdit    =   "0-0-0-1-1-3-2-3-1"
         CantEntero      =   15
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flePatOtrasEmpresa 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   9615
         _ExtentX        =   15478
         _ExtentY        =   5530
         Cols0           =   9
         ScrollBars      =   2
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#--Cod. SBS-Apellidos y Nombres-Razón Social-RUC-Aporte S/-%-"
         EncabezadosAnchos=   "350-0-1200-3000-1800-1200-1200-600-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-X"
         ListaControles  =   "0-0-1-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R-R-R-L"
         FormatosEdit    =   "0-0-0-1-1-3-2-3-0"
         CantEntero      =   15
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fleCargos 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   9615
         _ExtentX        =   15478
         _ExtentY        =   5530
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#--Cod. SBS-Nombres o Razón Social-Razón Social-RUC-Cargo-"
         EncabezadosAnchos=   "350-0-1200-3000-1800-1200-2000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5-6-X"
         ListaControles  =   "0-0-1-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R-L-C"
         FormatosEdit    =   "0-0-0-1-1-3-1-0"
         CantEntero      =   10
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Label lblPersElejida 
      Caption         =   "0"
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmPersonaJurDatosAdic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nPestana As Integer
Dim sResul As String
Dim nTipoPestana As String
Dim nTabLleno1, nTabLleno2, nTabLleno3, nTabLleno4 As Integer
Dim sTemp As String

Public oCargarDatos As New UPersona_Cli
Public oCargarPersona As UPersona_Cli

Private Sub cmdDEditar_Click()
If cmdDEditar.Caption = "&Editar" Then

    Call desahabilita(nPestana, False)
    cmdDEditar.Caption = "&Aceptar"
    
    Me.fleDirectorio.lbEditarFlex = True
    
    cmdNuevoG.Enabled = False
    cmdDEliminar.Enabled = False
    fleDirectorio.SetFocus

Else
    If cmdDEditar.Caption = "&Aceptar" Then
        
        If verificaLinea(fleDirectorio) = False And fleDirectorio.rows - 1 = 1 Then Call habilita(True): cmdDEditar.Caption = "&Editar": Me.fleDirectorio.lbEditarFlex = False: cmdNuevoG.Enabled = True: cmdDEliminar.Enabled = True: Exit Sub
        If verificaGrilla(fleDirectorio) = True Then MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO": Me.fleAccionistas.SetFocus: Exit Sub
                
        Call habilita(True)
        cmdDEditar.Caption = "&Editar"
        Me.fleDirectorio.lbEditarFlex = False
        cmdNuevoG.Enabled = True
        cmdDEliminar.Enabled = True
    
    End If
End If
End Sub

Private Sub cmdDEliminar_Click()
fleDirectorio.EliminaFila fleDirectorio.row
End Sub

Private Sub cmdEditar_Click()

If cmdEditar.Caption = "&Editar" Then

    Call desahabilita(nPestana, False)
    cmdEditar.Caption = "&Aceptar"
    
    Me.fleAccionistas.lbEditarFlex = True
    
    cmdNuevo.Enabled = False
    cmdEliminar.Enabled = False
    fleAccionistas.SetFocus

Else
    If cmdEditar.Caption = "&Aceptar" Then
        
        If verificaLinea(fleAccionistas) = False And fleAccionistas.rows - 1 = 1 Then Call habilita(True): cmdEditar.Caption = "&Editar": Me.fleAccionistas.lbEditarFlex = False: cmdNuevo.Enabled = True: cmdEliminar.Enabled = True: Exit Sub
        If verificaGrilla(fleAccionistas) = True Then MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO": Me.fleAccionistas.SetFocus: Exit Sub
        cmdEditar.Caption = "&Editar"
        Me.fleAccionistas.lbEditarFlex = False
        cmdNuevo.Enabled = True
        cmdEliminar.Enabled = True
        
        Call habilita(True)
    
    End If
End If

End Sub

Private Sub cmdEditarC_Click()
If cmdEditarC.Caption = "&Editar" Then
    
    Call desahabilita(nPestana, False)
    cmdEditarC.Caption = "&Aceptar"
    
    Me.fleCargos.lbEditarFlex = True
    
    cmdNuevoC.Enabled = False
    cmdEliminarC.Enabled = False
    fleCargos.SetFocus

Else
    If cmdEditarC.Caption = "&Aceptar" Then
        
        If verificaLinea(fleCargos) = False And fleCargos.rows - 1 = 1 Then Call habilita(True): cmdEditarC.Caption = "&Editar": Me.fleCargos.lbEditarFlex = False: cmdNuevoC.Enabled = True: cmdEliminarC.Enabled = True: Exit Sub
        If verificaGrilla(fleCargos) = True Then MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO": Me.fleAccionistas.SetFocus: Exit Sub
        Call habilita(True)
        cmdEditarC.Caption = "&Editar"
        Me.fleCargos.lbEditarFlex = False
        cmdNuevoC.Enabled = True
        cmdEliminarC.Enabled = True
    
    End If
End If
End Sub

Private Sub CmdEliminar_Click()
fleAccionistas.EliminaFila fleAccionistas.row
End Sub

Private Sub cmdEliminarC_Click()
fleCargos.EliminaFila fleCargos.row
End Sub

Private Sub cmdGEditar_Click()
If cmdGEditar.Caption = "&Editar" Then

    Call desahabilita(nPestana, False)
    cmdGEditar.Caption = "&Aceptar"
    
    Me.fleGerencias.lbEditarFlex = True
    
    cmdNuevoE.Enabled = False
    cmdGEliminar.Enabled = False
    fleDirectorio.SetFocus

Else
    If cmdGEditar.Caption = "&Aceptar" Then
        If verificaLinea(fleGerencias) = False And fleGerencias.rows - 1 = 1 Then Call habilita(True): cmdGEditar.Caption = "&Editar": Me.fleGerencias.lbEditarFlex = False: cmdNuevoE.Enabled = True: cmdGEliminar.Enabled = True: Exit Sub
        If verificaGrilla(fleGerencias) = True Then MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO": Me.fleAccionistas.SetFocus: Exit Sub
        Call habilita(True)
        cmdGEditar.Caption = "&Editar"
        Me.fleGerencias.lbEditarFlex = False
        cmdNuevoE.Enabled = True
        cmdGEliminar.Enabled = True
    
    End If
End If
End Sub

Private Sub cmdGEliminar_Click()
fleGerencias.EliminaFila fleGerencias.row
End Sub

Private Sub cmdNuevo_Click()

If cmdNuevo.Caption = "&Nuevo" Then

    Call desahabilita(nPestana, False)
    
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    
    fleAccionistas.lbEditarFlex = True
    fleAccionistas.AdicionaFila
    fleAccionistas.SetFocus
    
    Me.cmdNuevo.Caption = "&Aceptar"

ElseIf cmdNuevo.Caption = "&Aceptar" Then

    If verificaGrilla(fleAccionistas) = True Then
      MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO"
      Me.fleAccionistas.SetFocus
      Exit Sub

    End If

    Me.fleAccionistas.lbEditarFlex = False

    Call habilita(True)
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    Me.cmdNuevo.Caption = "&Nuevo"

End If

End Sub

Function verificaLinea(ByVal grilla As FlexEdit) As Boolean
Dim j As Integer
For j = 1 To grilla.cols - 1
  If grilla.TextMatrix(1, j) <> "" Then
    verificaLinea = True
    Exit Function
  End If
Next j
verificaLinea = False
End Function
Function verificaGrilla(ByVal grilla As FlexEdit) As Boolean
Dim i, j As Integer
Dim nCeldaV As Integer
Dim nUltFila, nUltColum As Integer

nUltColum = grilla.cols - 1 '- 1
nUltFila = grilla.rows - 1
    For i = 1 To grilla.rows - 1
        
        For j = 1 To grilla.cols - 1
        
            nCeldaV = grilla.cols - 1
                If nCeldaV = j And nUltFila = i Then
                    verificaGrilla = False
                    Exit Function
                End If
                   
           If SSTabs.Tab = 3 Or SSTabs.Tab = 4 Or SSTabs.Tab = 5 Then
            'grilla.TextMatrix(i, 1) = 0
               If j <> 2 Then
                If grilla.TextMatrix(i, j) = "" And j < nUltColum Then
                    verificaGrilla = True
                    Exit Function
                End If
               End If
           Else
    
                   If grilla.TextMatrix(i, j) = "" And j < nUltColum Then
                     verificaGrilla = True
                     Exit Function
                   End If
           
           End If

           
        Next j
       
    Next i

 verificaGrilla = False

End Function

Private Sub cmdNuevoC_Click()
If cmdNuevoC.Caption = "&Nuevo" Then

    Call desahabilita(nPestana, False)
    
    cmdEditarC.Enabled = False
    cmdEliminarC.Enabled = False
    
    Me.fleCargos.lbEditarFlex = True
    Me.fleCargos.AdicionaFila
    Me.fleCargos.TextMatrix(Me.fleCargos.row, 1) = "0"

    fleCargos.SetFocus
    
    Me.cmdNuevoC.Caption = "&Aceptar"

ElseIf cmdNuevoC.Caption = "&Aceptar" Then

    If verificaGrilla(fleCargos) = True Then
      MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO"
      Me.fleCargos.SetFocus
      Exit Sub
    End If

    Me.fleCargos.lbEditarFlex = False

    Call habilita(True)
    cmdEditarC.Enabled = True
    cmdEliminarC.Enabled = True
    Me.cmdNuevoC.Caption = "&Nuevo"

End If

End Sub

Private Sub cmdNuevoE_Click()

If cmdNuevoE.Caption = "&Nuevo" Then

    Call desahabilita(nPestana, False)
    
    cmdGEditar.Enabled = False
    cmdGEliminar.Enabled = False
    
    Me.fleGerencias.lbEditarFlex = True
    Me.fleGerencias.AdicionaFila

    fleGerencias.SetFocus
    
    Me.cmdNuevoE.Caption = "&Aceptar"

ElseIf cmdNuevoE.Caption = "&Aceptar" Then

    If verificaGrilla(fleGerencias) = True Then
      MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO"
      Me.fleGerencias.SetFocus
      Exit Sub
    End If

    Me.fleGerencias.lbEditarFlex = False

    Call habilita(True)
    cmdGEditar.Enabled = True
    cmdGEliminar.Enabled = True
    Me.cmdNuevoE.Caption = "&Nuevo"

End If
End Sub

Private Sub cmdNuevoG_Click()

If cmdNuevoG.Caption = "&Nuevo" Then

    Call desahabilita(nPestana, False)
    
    Me.fleDirectorio.lbEditarFlex = True
    Me.fleDirectorio.AdicionaFila
    
    cmdDEditar.Enabled = False
    cmdDEliminar.Enabled = False
    
   ' FERelPersNoMoverdeFila = fleAccionistas.row - 1
    fleDirectorio.SetFocus
    
    Me.cmdNuevoG.Caption = "&Aceptar"

ElseIf cmdNuevoG.Caption = "&Aceptar" Then

    If verificaGrilla(fleDirectorio) = True Then
      MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO"
      Me.fleDirectorio.SetFocus
      Exit Sub
    End If

    Me.fleDirectorio.lbEditarFlex = False
    'guardar datos
   ' MsgBox "Datos Guardados Corectamente como '" & nPestana & "' ", vbDefaultButton3, "SICMACT"
    Call habilita(True)
    cmdDEditar.Enabled = True
    cmdDEliminar.Enabled = True
    Me.cmdNuevoG.Caption = "&Nuevo"

End If
End Sub

Private Sub cmdPatrimonio1_Click()

If cmdPatrimonio1.Caption = "&Nuevo" Then

    Call desahabilita(nPestana, False)
    
    Me.flePatrimonio.lbEditarFlex = True
    Me.flePatrimonio.AdicionaFila
    Me.flePatrimonio.TextMatrix(Me.flePatrimonio.row, 1) = "0"
    
    cmdPEditar.Enabled = False
    cmdPEliminar.Enabled = False
    
   ' FERelPersNoMoverdeFila = fleAccionistas.row - 1
    flePatrimonio.SetFocus
    
    Me.cmdPatrimonio1.Caption = "&Aceptar"

ElseIf cmdPatrimonio1.Caption = "&Aceptar" Then

    If verificaGrilla(flePatrimonio) = True Then
      MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO"
      Me.flePatrimonio.SetFocus
      Exit Sub
    End If

    Me.flePatrimonio.lbEditarFlex = False
    'guardar datos
   ' MsgBox "Datos Guardados Corectamente como '" & nPestana & "' ", vbDefaultButton3, "SICMACT"
    Call habilita(True)
    cmdPEditar.Enabled = True
    cmdPEliminar.Enabled = True
    Me.cmdPatrimonio1.Caption = "&Nuevo"

End If

End Sub

Private Sub cmdPatrimonio2_Click()
If cmdPatrimonio2.Caption = "&Nuevo" Then

    Call desahabilita(nPestana, False)
    
    Me.flePatOtrasEmpresa.lbEditarFlex = True
    Me.flePatOtrasEmpresa.AdicionaFila
    Me.flePatOtrasEmpresa.TextMatrix(Me.flePatOtrasEmpresa.row, 1) = "0"
    
    cmdPEEditar.Enabled = False
    cmdPEEliminar.Enabled = False
    
   ' FERelPersNoMoverdeFila = fleAccionistas.row - 1
    flePatOtrasEmpresa.SetFocus
    
    Me.cmdPatrimonio2.Caption = "&Aceptar"

ElseIf cmdPatrimonio2.Caption = "&Aceptar" Then

    If verificaGrilla(flePatOtrasEmpresa) = True Then
      MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO"
      Me.flePatOtrasEmpresa.SetFocus
      Exit Sub
    End If

    Me.flePatOtrasEmpresa.lbEditarFlex = False
    'guardar datos
   ' MsgBox "Datos Guardados Corectamente como '" & nPestana & "' ", vbDefaultButton3, "SICMACT"
    Call habilita(True)
    cmdPEEditar.Enabled = True
    cmdPEEliminar.Enabled = True
    Me.cmdPatrimonio2.Caption = "&Nuevo"

End If
End Sub

Private Sub cmdPEditar_Click()
If cmdPEditar.Caption = "&Editar" Then

    Call desahabilita(nPestana, False)
    cmdPEditar.Caption = "&Aceptar"
    
    Me.flePatrimonio.lbEditarFlex = True
    
    cmdPatrimonio1.Enabled = False
    cmdPEliminar.Enabled = False
    flePatrimonio.SetFocus

Else
    If cmdPEditar.Caption = "&Aceptar" Then
        
        If verificaLinea(flePatrimonio) = False And flePatrimonio.rows - 1 = 1 Then Call habilita(True): cmdPEditar.Caption = "&Editar": Me.flePatrimonio.lbEditarFlex = False: cmdPatrimonio1.Enabled = True: cmdPEliminar.Enabled = True: Exit Sub
        If verificaGrilla(flePatrimonio) = True Then MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO": Me.fleAccionistas.SetFocus: Exit Sub
        Call habilita(True)
        cmdPEditar.Caption = "&Editar"
        Me.flePatrimonio.lbEditarFlex = False
        cmdPatrimonio1.Enabled = True
        cmdPEliminar.Enabled = True
    
    End If
End If
End Sub

Private Sub cmdPEEditar_Click()
If cmdPEEditar.Caption = "&Editar" Then
    
    Call desahabilita(nPestana, False)
    cmdPEEditar.Caption = "&Aceptar"
    
    Me.flePatOtrasEmpresa.lbEditarFlex = True
    
    cmdPatrimonio2.Enabled = False
    cmdPEEliminar.Enabled = False
    flePatrimonio.SetFocus

Else
    If cmdPEEditar.Caption = "&Aceptar" Then
        If verificaLinea(flePatOtrasEmpresa) = False And flePatOtrasEmpresa.rows - 1 = 1 Then Call habilita(True): cmdPEEditar.Caption = "&Editar": Me.flePatOtrasEmpresa.lbEditarFlex = False: cmdPatrimonio2.Enabled = True: cmdPEEliminar.Enabled = True: Exit Sub
        If verificaGrilla(flePatOtrasEmpresa) = True Then MsgBox "Existen campos vacios, por favor verifique", vbInformation, "AVISO": Me.fleAccionistas.SetFocus: Exit Sub
        Call habilita(True)
        cmdPEEditar.Caption = "&Editar"
        Me.flePatOtrasEmpresa.lbEditarFlex = False
        cmdPatrimonio2.Enabled = True
        cmdPEEliminar.Enabled = True
    
    End If
End If
End Sub

Private Sub cmdPEEliminar_Click()
flePatOtrasEmpresa.EliminaFila flePatOtrasEmpresa.row
End Sub

Private Sub cmdPEliminar_Click()
flePatrimonio.EliminaFila flePatrimonio.row
End Sub

'    If nPestana = 0 Then 'Accionistas
'        nTipoPestana = "A"
'    ElseIf nPestana = 1 Then 'Directorio
'        nTipoPestana = "D"
'    ElseIf nPestana = 2 Then 'Gerencias
'        nTipoPestana = "G"
'    ElseIf nPestana = 3 Then 'Patrimonio
'        nTipoPestana = "P"
'    End if
Private Sub cmdsalir_Click()
Dim nFichasLlenas As Integer
Dim i As Integer

'Dim sTipo As TpoDetalle

'Dim odetallesjur As New UPersona_Cli
'GuardarDatos

If lblPersElejida.Caption = 1 Then Me.Hide

If lblPersElejida.Caption = 2 Then
    nFichasLlenas = VerificaFichas(SSTabs)
    
    If nFichasLlenas = -1 Then
        Me.Hide
        Exit Sub
    End If
    
    SSTabs.Tab = nFichasLlenas
    
    Dim sResult As String
    
    sResult = MsgBox("No se ha llenado información relevante a Accionistas, Directorio, Gerencias y/o Patrimonio: Desea salir de Todas Maneras?", vbYesNo + vbInformation, "AVISO")
    
    If sResult = vbYes Then
            Me.Hide
        Exit Sub
    End If
    
    'MsgBox "Se necesita mas información, verifique", vbInformation, "SICMACT"
    
End If

End Sub

Function VerificaFichas(ByVal stabs As SSTAB) As Integer
Dim i As Integer
Dim GrillaEsta As Boolean


 For i = 0 To stabs.Tabs - 1

    If i < 4 Then
    
        If i = 0 Then
    
            GrillaEsta = verificaGrilla(fleAccionistas)
            If GrillaEsta = True Then VerificaFichas = i: Exit Function
        
        End If
        If i = 1 Then
    
            GrillaEsta = verificaGrilla(fleDirectorio)
            If GrillaEsta = True Then VerificaFichas = i: Exit Function
        
        End If
        If i = 2 Then
    
            GrillaEsta = verificaGrilla(fleGerencias)
            If GrillaEsta = True Then VerificaFichas = i: Exit Function
        
        End If
        If i = 3 Then
    
            GrillaEsta = verificaGrilla(flePatrimonio)
            If GrillaEsta = True Then VerificaFichas = i: Exit Function
        End If
        If i = 5 Then
    
            GrillaEsta = verificaGrilla(fleCargos)
            If GrillaEsta = True Then VerificaFichas = i: Exit Function
        End If
    End If
 
 Next i
 VerificaFichas = -1

End Function

Sub habilita(ByVal pnEstaFicha As Boolean)
Dim i As Integer
For i = 0 To 5
        Me.SSTabs.TabEnabled(i) = pnEstaFicha
Next i
End Sub
Sub desahabilita(ByVal pnFicha As Integer, ByVal pnEstaFicha As Boolean)
Dim i As Integer
For i = 0 To 5
    If pnFicha = i Then
        Me.SSTabs.TabEnabled(pnFicha) = True
    Else
        Me.SSTabs.TabEnabled(i) = pnEstaFicha
    End If
Next i
End Sub

Private Sub fleAccionistas_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 End If
End Sub

Private Sub fleAccionistas_OnCellChange(pnRow As Long, pnCol As Long)
'ValidarEntradas
If pnCol = 5 Then fleAccionistas.TextMatrix(fleAccionistas.row, fleAccionistas.Col) = EliminarString(fleAccionistas.TextMatrix(fleAccionistas.row, fleAccionistas.Col), "-")
sResul = ValidarEntradasADG(fleAccionistas, pnRow, pnCol)

If sResul = "" Then
  Exit Sub
Else
     MsgBox sResul, vbInformation, "AVISO"
     fleAccionistas.TextMatrix(pnRow, pnCol) = ""
     fleAccionistas.SetFocus
End If

End Sub

Private Sub fleAccionistas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
  If pnCol = 1 And Len(fleAccionistas.TextMatrix(pnRow, pnCol)) > 200 Then fleAccionistas.TextMatrix(pnRow, pnCol) = Left(fleAccionistas.TextMatrix(pnRow, pnCol), 200)
  If pnCol = 4 And Len(fleAccionistas.TextMatrix(pnRow, pnCol)) > 20 Then fleAccionistas.TextMatrix(pnRow, pnCol) = Left(fleAccionistas.TextMatrix(pnRow, pnCol), 20)
  
End Sub


Private Sub fleCargos_OnCellChange(pnRow As Long, pnCol As Long)
'ValidarEntradas
sResul = ValidarEntradasPatrimonio(fleCargos, pnRow, pnCol)

If sResul = "" Then
  Exit Sub
Else
     MsgBox sResul, vbInformation, "AVISO"
     fleCargos.TextMatrix(pnRow, pnCol) = ""
     fleCargos.SetFocus
End If

End Sub

Private Sub fleCargos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
If psDataCod <> "" Then

Set oCargarPersona = Nothing
Set oCargarPersona = New UPersona_Cli
Call oCargarPersona.RecuperaPersona(psDataCod)

fleCargos.TextMatrix(pnRow, 3) = oCargarPersona.NombreCompleto
fleCargos.TextMatrix(pnRow, 2) = oCargarPersona.PersCodSbs

End If
End Sub

Private Sub fleCargos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim nCaracteres As Integer
  nCaracteres = Len(flePatrimonio.TextMatrix(pnRow - 1, pnCol))
  If pnCol = 3 And nCaracteres > 200 Then flePatrimonio.TextMatrix(pnRow - 1, pnCol) = Left(flePatrimonio.TextMatrix(pnRow - 1, pnCol), 200)
  nCaracteres = Len(flePatrimonio.TextMatrix(pnRow - 1, pnCol))
  If pnCol = 4 And nCaracteres > 150 Then flePatrimonio.TextMatrix(pnRow - 1, pnCol) = Left(flePatrimonio.TextMatrix(pnRow - 1, pnCol), 150)
End Sub

Private Sub fleDirectorio_OnCellChange(pnRow As Long, pnCol As Long)

'ValidarEntradas
sResul = ValidarEntradasADG(fleDirectorio, pnRow, pnCol)

If sResul = "" Then
  Exit Sub
Else
            MsgBox sResul, vbInformation, "AVISO"
            fleDirectorio.TextMatrix(pnRow, pnCol) = ""
            fleDirectorio.SetFocus
End If

End Sub

Function ValidarEntradasPatrimonio(ByVal grilla As FlexEdit, ByVal npnRow As Long, npnCol As Long) As String

Dim nRuc As String
Dim nNruc As Integer
Dim nVpor As Double
 
Dim nCeros As String
Dim j As Integer
nCeros = ""

nRuc = Format(grilla.TextMatrix(npnRow, npnCol))
nNruc = Len(nRuc)

 If grilla.TextMatrix(npnRow, npnCol) <> "" Then
    
'Pestaña 3
If SSTabs.Tab = 3 Or SSTabs.Tab = 4 Then
'If npnCol = 2 And SSTabs.Tab = 3 Then grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasPatrimonio = "": Exit Function
    If npnCol = 3 Then grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasPatrimonio = "": Exit Function
    If npnCol = 4 Then grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasPatrimonio = "": Exit Function
       
   'VALIDA LA CANTIDAD DE DIGITOS PARA EL RUC
        If npnCol = 5 And Len(nRuc) = 11 Then
          grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), "00000000000")
          ValidarEntradasPatrimonio = "": Exit Function
        Else
          If npnCol = 5 And Len(nRuc) > 11 Then
              'ValidarEntradasPatrimonio = "RUC No Válido: No puede ser mayor de 11 Digitos, por favor reintente"
              ValidarEntradasPatrimonio = ""
              grilla.TextMatrix(npnRow, npnCol) = Left(Format(grilla.TextMatrix(npnRow, npnCol), "00000000000"), 11)
              Exit Function
              
          ElseIf npnCol = 5 And (Len(nRuc) > 0 Or Len(nRuc) < 11) Then
            If CDbl(grilla.TextMatrix(npnRow, npnCol)) = 0 Then grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasPatrimonio = "": Exit Function
'            For j = 1 To Len(nRuc)
'             nCeros = nCeros + "0"
'            Next j
          
             ValidarEntradasPatrimonio = ""
             'grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
              grilla.TextMatrix(npnRow, npnCol) = ""
             Exit Function
          End If
        End If
        'valida el porcentaje
        If npnCol = 7 Then
          
           nVpor = grilla.TextMatrix(npnRow, npnCol)
               If (nVpor < 0 Or nVpor > 100) Then
                    ValidarEntradasPatrimonio = "El Aporte en Porcentaje no debe Superar del 100%"
                    Exit Function
               End If
        End If
        
   End If

'If SSTabs.Tab = 4 Then
'    If (npnCol = 3 Or npnCol = 4) Then grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasPatrimonio = "": Exit Function
'   'VALIDA LA CANTIDAD DE DIGITOS PARA EL RUC
'        If npnCol = 5 And Len(nRuc) = 11 Then
'          grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), "00000000000")
'          ValidarEntradasPatrimonio = "": Exit Function
'        Else
'          If npnCol = 5 And Len(nRuc) > 11 Then
'              'ValidarEntradasPatrimonio = "RUC No Válido: No puede ser mayor de 11 Digitos, por favor reintente"
'              ValidarEntradasPatrimonio = ""
'              grilla.TextMatrix(npnRow, npnCol) = Left(Format(grilla.TextMatrix(npnRow, npnCol), "00000000000"), 11)
'              Exit Function
'
'          ElseIf npnCol = 5 And (Len(nRuc) > 0 Or Len(nRuc) < 11) Then
'
'            For j = 1 To Len(nRuc)
'             nCeros = nCeros + "0"
'            Next j
'
'             ValidarEntradasPatrimonio = ""
'             grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
'
'             Exit Function
'          End If
'        End If
'
'        'valida el porcentaje
'        If npnCol = 7 Then
'
'           nVpor = grilla.TextMatrix(npnRow, npnCol)
'               If (nVpor < 0 Or nVpor > 100) Then
'                    ValidarEntradasPatrimonio = "El Aporte en Porcentaje no debe Superar del 100%"
'                    Exit Function
'               End If
'        End If
'
'
'End If
   

If SSTabs.Tab = 5 Then

    If (npnCol = 3 Or npnCol = 4 Or npnCol = 6) Then grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasPatrimonio = "": Exit Function

   'VALIDA LA CANTIDAD DE DIGITOS PARA EL RUC
        If npnCol = 5 And Len(nRuc) = 11 Then
          grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), "00000000000")
          ValidarEntradasPatrimonio = "": Exit Function
        Else
          If npnCol = 5 And Len(nRuc) > 11 Then
              'ValidarEntradasPatrimonio = "RUC No Válido: No puede ser mayor de 11 Digitos, por favor reintente"
              ValidarEntradasPatrimonio = ""
              grilla.TextMatrix(npnRow, npnCol) = Left(Format(grilla.TextMatrix(npnRow, npnCol), "00000000000"), 11)
              Exit Function
              
          ElseIf npnCol = 5 And (Len(nRuc) > 0 Or Len(nRuc) < 11) Then
             If CDbl(grilla.TextMatrix(npnRow, npnCol)) = 0 Then grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasPatrimonio = "": Exit Function
'            For j = 1 To Len(nRuc)
'             nCeros = nCeros + "0"
'            Next j

             ValidarEntradasPatrimonio = ""
             'grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
             grilla.TextMatrix(npnRow, npnCol) = ""
             'flePatrimonio.Col = npnCol: flePatrimonio.row = npnRow
             Exit Function
          End If
        End If
End If

        If npnCol = 6 And SSTabs.Caption = "Patrimonio en Otras Entidades" Then
        
         grilla.TextMatrix(npnRow, npnCol) = Format(Abs(grilla.TextMatrix(npnRow, npnCol)), "#,##.00")
         
        End If
        If npnCol = 6 And SSTabs.Caption = "Participación Patrimonial en Empresas" Then
           grilla.TextMatrix(npnRow, npnCol) = Format(Abs(grilla.TextMatrix(npnRow, npnCol)), "#,##.00")
        End If
    
 End If
ValidarEntradasPatrimonio = ""
End Function

Function ValidarEntradasADG(ByVal grilla As FlexEdit, ByVal npnRow As Long, npnCol As Long) As String

Dim nRuc As String
Dim nNruc As Integer

Dim nCeros As String
Dim j As Integer
nCeros = ""

nRuc = Format(grilla.TextMatrix(npnRow, npnCol))
nNruc = Len(nRuc)

 If grilla.TextMatrix(npnRow, npnCol) <> "" Then
    
    If npnCol = 1 Then
        
        'verificar si son letras
    
        grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol))
        ValidarEntradasADG = ""
        Exit Function
    End If

   'VALIDA LA CANTIDAD DE DIGITOS PARA EL RUC
        If npnCol = 2 And Len(nRuc) = 11 Then
          grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), "00000000000")
          ValidarEntradasADG = "": Exit Function
        Else
          If npnCol = 2 And Len(nRuc) > 11 Then
              grilla.TextMatrix(npnRow, npnCol) = Left(Format(grilla.TextMatrix(npnRow, npnCol), "00000000000"), 11)
              ValidarEntradasADG = ""
              Exit Function
              
          ElseIf npnCol = 2 And (Len(nRuc) > 0 Or Len(nRuc) < 11) Then
            If CDbl(grilla.TextMatrix(npnRow, npnCol)) = 0 Then grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasADG = "": Exit Function

             ValidarEntradasADG = ""
             'grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
             grilla.TextMatrix(npnRow, npnCol) = ""
             Exit Function
             
          End If
          End If

      'VALIDA LA CANTIDAD DE DIGITOS PARA EL DNI0=
      If npnCol = 3 Then
        Dim valorcero As String
        
        If Not IsNumeric(grilla.TextMatrix(npnRow, npnCol)) Then
            grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasADG = "": Exit Function
        End If
        
        valorcero = Mid(grilla.TextMatrix(npnRow, npnCol), 1, 1)
        If valorcero = "0" Then nRuc = valorcero & Format(grilla.TextMatrix(npnRow, npnCol))
        valorcero = Mid(grilla.TextMatrix(npnRow, npnCol), 1, 2)
        If valorcero = "00" Then nRuc = valorcero & Format(grilla.TextMatrix(npnRow, npnCol))
        valorcero = Mid(grilla.TextMatrix(npnRow, npnCol), 1, 3)
        If valorcero = "000" Then nRuc = valorcero & Format(grilla.TextMatrix(npnRow, npnCol))
        valorcero = Mid(grilla.TextMatrix(npnRow, npnCol), 1, 4)
        If valorcero = "0000" Then nRuc = valorcero & Format(grilla.TextMatrix(npnRow, npnCol))
        valorcero = Mid(grilla.TextMatrix(npnRow, npnCol), 1, 5)
        If valorcero = "00000" Then nRuc = valorcero & Format(grilla.TextMatrix(npnRow, npnCol))
        
        nNruc = Len(nRuc)
        For j = 1 To Len(nRuc)
            nCeros = nCeros + "0"
        Next j
        
            If nNruc = 8 Then grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
            If nNruc >= 9 And nNruc <= 11 Then grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasADG = "": Exit Function
            If nNruc < 8 Or nNruc > 12 Then grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasADG = "": Exit Function
            If nNruc = 12 Then grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
        End If
        
'            If npnCol = 3 And (nNruc >= 8 Or nNruc) <= 12 Then
'              grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
'            Else
'              If npnCol = 3 And Len(nRuc) > 12 Then
'                  ValidarEntradasADG = ""
'                  grilla.TextMatrix(npnRow, npnCol) = Left(Format(grilla.TextMatrix(npnRow, npnCol), "00000000"), 8)
'                  Exit Function
'
'              ElseIf npnCol = 3 And (Len(nRuc) > 0 Or Len(nRuc) < 12) Then
'
'                   If CDbl(grilla.TextMatrix(npnRow, npnCol)) = 0 Then grilla.TextMatrix(npnRow, npnCol) = "": ValidarEntradasADG = "": Exit Function
'                  grilla.TextMatrix(npnRow, npnCol) = Format(grilla.TextMatrix(npnRow, npnCol), nCeros)
'
'              End If
'            End If
       
      'columna 4
        If npnCol = 4 Then
          grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasADG = "": Exit Function
        End If
        If npnCol = 5 And SSTabs.Caption = "Directorio" Then
          grilla.TextMatrix(npnRow, npnCol) = UCase(grilla.TextMatrix(npnRow, npnCol)): ValidarEntradasADG = "": Exit Function
        End If
        'COLUMNA 5
        If npnCol = 5 And SSTabs.Caption = "Accionistas" Then
           grilla.TextMatrix(npnRow, npnCol) = Format(Abs(grilla.TextMatrix(npnRow, npnCol)), "#,##.00")
        End If
        
         ' Columna 6
        If npnCol = 6 And SSTabs.Caption = "Accionistas" Then
           If (CInt(grilla.TextMatrix(npnRow, npnCol)) < 0 Or CInt(grilla.TextMatrix(npnRow, npnCol)) > 100) Then
                ValidarEntradasADG = "El Aporte en Porcentaje no debe Superar del 100%"
                Exit Function
           End If
        End If
      
 
 End If
  ValidarEntradasADG = ""
End Function

Private Sub fleDirectorio_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
  If pnCol = 1 And Len(fleDirectorio.TextMatrix(pnRow, pnCol)) > 200 Then fleDirectorio.TextMatrix(pnRow, pnCol) = Left(fleDirectorio.TextMatrix(pnRow, pnCol), 200)
  If pnCol = 4 And Len(fleDirectorio.TextMatrix(pnRow, pnCol)) > 100 Then fleDirectorio.TextMatrix(pnRow, pnCol) = Left(fleDirectorio.TextMatrix(pnRow, pnCol), 100)
  If pnCol = 5 And Len(fleDirectorio.TextMatrix(pnRow, pnCol)) > 20 Then fleDirectorio.TextMatrix(pnRow, pnCol) = Left(fleDirectorio.TextMatrix(pnRow, pnCol), 20)
End Sub
Private Sub fleGerencias_OnCellChange(pnRow As Long, pnCol As Long)
'ValidarEntradas
sResul = ValidarEntradasADG(fleGerencias, pnRow, pnCol)

If sResul = "" Then
  Exit Sub
Else
     MsgBox sResul, vbInformation, "AVISO"
     fleGerencias.TextMatrix(pnRow, pnCol) = ""
     fleGerencias.SetFocus
End If
End Sub

Private Sub fleGerencias_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
  If pnCol = 1 And Len(fleGerencias.TextMatrix(pnRow, pnCol)) > 200 Then fleGerencias.TextMatrix(pnRow, pnCol) = Left(fleGerencias.TextMatrix(pnRow, pnCol), 200)
  If pnCol = 4 And Len(fleGerencias.TextMatrix(pnRow, pnCol)) > 100 Then fleGerencias.TextMatrix(pnRow, pnCol) = Left(fleGerencias.TextMatrix(pnRow, pnCol), 100)
End Sub

Private Sub flePatOtrasEmpresa_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 6 Then flePatOtrasEmpresa.TextMatrix(flePatOtrasEmpresa.row, flePatOtrasEmpresa.Col) = EliminarString(flePatOtrasEmpresa.TextMatrix(flePatOtrasEmpresa.row, flePatOtrasEmpresa.Col), "-")

sResul = ValidarEntradasPatrimonio(flePatOtrasEmpresa, pnRow, pnCol)

If sResul = "" Then
  Exit Sub
Else
     MsgBox sResul, vbInformation, "AVISO"
     flePatOtrasEmpresa.TextMatrix(pnRow, pnCol) = ""
     flePatOtrasEmpresa.SetFocus
End If
End Sub

Private Sub flePatOtrasEmpresa_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
If psDataCod <> "" Then

Set oCargarPersona = Nothing
Set oCargarPersona = New UPersona_Cli
Call oCargarPersona.RecuperaPersona(psDataCod)

flePatOtrasEmpresa.TextMatrix(pnRow, 3) = oCargarPersona.NombreCompleto
flePatOtrasEmpresa.TextMatrix(pnRow, 2) = oCargarPersona.PersCodSbs

End If
End Sub

Private Sub flePatOtrasEmpresa_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
  If pnCol = 3 And Len(flePatOtrasEmpresa.TextMatrix(pnRow, pnCol)) > 200 Then flePatOtrasEmpresa.TextMatrix(pnRow, pnCol) = Left(flePatOtrasEmpresa.TextMatrix(pnRow, pnCol), 200)
  If pnCol = 4 And Len(flePatOtrasEmpresa.TextMatrix(pnRow, pnCol)) > 150 Then flePatOtrasEmpresa.TextMatrix(pnRow, pnCol) = Left(flePatOtrasEmpresa.TextMatrix(pnRow, pnCol), 150)
End Sub

'Public Sub validaNum(ByVal valor As Integer, ByVal pCol As Integer, ByVal prow As Integer)
'Dim cntRuc As Integer
'cntRuc = valor
' If cntRuc > 11 And pCol = 5 Then flePatrimonio.Col = pCol: flePatrimonio.row = prow
' If cntRuc < 11 And pCol = 5 Then flePatrimonio.Col = pCol: flePatrimonio.row = prow
'End Sub

Private Sub flePatrimonio_OnCellChange(pnRow As Long, pnCol As Long)
'ValidarEntradas
If pnCol = 6 Then flePatrimonio.TextMatrix(flePatrimonio.row, flePatrimonio.Col) = EliminarString(flePatrimonio.TextMatrix(flePatrimonio.row, flePatrimonio.Col), "-")
 
sResul = ValidarEntradasPatrimonio(flePatrimonio, pnRow, pnCol)

If sResul = "" Then
  Exit Sub
Else
     MsgBox sResul, vbInformation, "AVISO"
     flePatrimonio.TextMatrix(pnRow, pnCol) = ""
     flePatrimonio.SetFocus
End If
End Sub

Private Sub flePatrimonio_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
If psDataCod <> "" Then

Set oCargarPersona = Nothing
Set oCargarPersona = New UPersona_Cli
Call oCargarPersona.RecuperaPersona(psDataCod)

flePatrimonio.TextMatrix(pnRow, 3) = oCargarPersona.NombreCompleto
flePatrimonio.TextMatrix(pnRow, 2) = oCargarPersona.PersCodSbs

End If

End Sub

Private Sub flePatrimonio_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
  If pnCol = 3 And Len(flePatrimonio.TextMatrix(pnRow, pnCol)) > 200 Then flePatrimonio.TextMatrix(pnRow, pnCol) = Left(flePatrimonio.TextMatrix(pnRow, pnCol), 200)
  If pnCol = 4 And Len(flePatrimonio.TextMatrix(pnRow, pnCol)) > 150 Then flePatrimonio.TextMatrix(pnRow, pnCol) = Left(flePatrimonio.TextMatrix(pnRow, pnCol), 150)
End Sub

Private Sub Form_Load()
Me.SSTabs.Tab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim nFichasLlenas As Integer
'GuardarDatos
Cancel = True

End Sub

Private Sub SSTabs_Click(PreviousTab As Integer)
  nPestana = Me.SSTabs.Tab
End Sub

Public Function EliminarString(Cadena As String, aEliminar As String)

    Dim nroDigi As Integer
Dim n As String
    nroDigi = Len(Cadena)
    
    n = Replace(Cadena, aEliminar, "")
    EliminarString = n

End Function
