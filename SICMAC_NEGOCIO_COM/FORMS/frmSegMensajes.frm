VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSegMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje de Seguridad"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   Icon            =   "frmSegMensajes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Mensajes"
      TabPicture(0)   =   "frmSegMensajes.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraMensajes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraMensajes 
         Caption         =   "Mensajes de Seguridad Registrados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9375
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
            Height          =   375
            Left            =   1440
            TabIndex        =   6
            Top             =   3720
            Width           =   1170
         End
         Begin VB.CommandButton cmdQuitar 
            Caption         =   "Quitar"
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
            Left            =   2760
            TabIndex        =   5
            Top             =   3720
            Width           =   1170
         End
         Begin VB.CommandButton cmdCerrar 
            Caption         =   "Cerrar"
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
            Left            =   8040
            TabIndex        =   4
            Top             =   3720
            Width           =   1170
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "Nuevo"
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
            Left            =   120
            TabIndex        =   3
            Top             =   3720
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feMensajes 
            Height          =   3255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5741
            Cols0           =   4
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Mensaje-Estado-Aux"
            EncabezadosAnchos=   "500-7000-1000-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-X"
            ListaControles  =   "0-0-4-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-L"
            FormatosEdit    =   "0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
End
Attribute VB_Name = "frmSegMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina         :   frmSegMensajes
'***     Descripcion    :   Formulario para la administracion de los Mensaje de Seguridad
'***     Creado por     :   WIOR
'***     Maquina        :   TIF-1-19
'***     Fecha-Creación :   01/09/2013 08:20:00 AM
'*****************************************************************************************
Option Explicit
Private FEMoverFila As Integer
Private fbEdicion As Boolean

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdEditar_Click()
FEMoverFila = feMensajes.Row
feMensajes.lbEditarFlex = True
feMensajes.SetFocus
SendKeys "{Enter}"
End Sub

Private Sub cmdNuevo_Click()
Me.feMensajes.AdicionaFila
FEMoverFila = feMensajes.Rows - 1
feMensajes.lbEditarFlex = True
feMensajes.SetFocus
SendKeys "{Enter}"
End Sub

Private Sub cmdQuitar_Click()
If MsgBox("Estas Seguro de eliminar el registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim IdMensaje As String
    IdMensaje = Trim(feMensajes.TextMatrix(feMensajes.Row, 3))
    
    Dim oSeg As COMDPersona.UCOMAcceso
    Set oSeg = New COMDPersona.UCOMAcceso
    
    If IdMensaje <> "" Then
        Call oSeg.OpeMensajeSeguridad(2, "", False, CInt(IdMensaje))
    End If
    feMensajes.EliminaFila feMensajes.Row
End If
End Sub

Private Sub feMensajes_OnCellChange(pnRow As Long, pnCol As Long)
If ValidaDatos(pnRow) Then
    Registro
End If
End Sub

Private Sub feMensajes_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If ValidaDatos(pnRow) Then
    Registro
End If
End Sub

Private Function ValidaDatos(Optional ByVal pnRow As Long = 0) As Boolean
ValidaDatos = True
If Trim(feMensajes.TextMatrix(pnRow, 1)) = "" Then
    MsgBox "Ingrese un Mensaje de Seguridad en la fila " & pnRow, vbCritical, "Aviso"
    ValidaDatos = False
    CargaMensajes
    Exit Function
End If
End Function
Private Sub Registro()
Dim oSeg As COMDPersona.UCOMAcceso
Dim i As Integer
Dim IdMensaje As String
Dim Mensaje As String
Dim Activo As String
Set oSeg = New COMDPersona.UCOMAcceso

For i = 1 To feMensajes.Rows - 1
    IdMensaje = Trim(feMensajes.TextMatrix(i, 3))
    Mensaje = Trim(feMensajes.TextMatrix(i, 1))
    Activo = Trim(feMensajes.TextMatrix(i, 2))
    
    If Trim(IdMensaje) = "" Then
        IdMensaje = oSeg.InsertaMensajeSeguridad(Mensaje, IIf(Activo = ".", True, False))
        feMensajes.TextMatrix(i, 3) = IdMensaje
    Else
        Call oSeg.OpeMensajeSeguridad(1, Mensaje, IIf(Activo = ".", True, False), CInt(IdMensaje))
    End If
Next i
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    feMensajes.lbEditarFlex = False
    CargaMensajes
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
CargaMensajes
fbEdicion = False
FEMoverFila = 1
Me.feMensajes.lbEditarFlex = False
End Sub

Private Sub CargaMensajes()
Dim oSeg As COMDPersona.UCOMAcceso
Dim rsSeg As ADODB.Recordset
Dim i As Long
Dim CantReg As Long
Set oSeg = New COMDPersona.UCOMAcceso
Set rsSeg = oSeg.ObtenerMensajeSeguridad
CantReg = 0
LimpiaFlex feMensajes
If Not (rsSeg.EOF And rsSeg.BOF) Then
    CantReg = rsSeg.RecordCount
    For i = 1 To CantReg
        feMensajes.AdicionaFila
        feMensajes.TextMatrix(i, 1) = Trim(rsSeg!cMensaje)
        feMensajes.TextMatrix(i, 2) = IIf(CBool(rsSeg!bEstado), "1", "")
        feMensajes.TextMatrix(i, 3) = Trim(rsSeg!IdMensaje)
        rsSeg.MoveNext
    Next i
End If

feMensajes.Row = 1
feMensajes.TopRow = 1

Set oSeg = Nothing
Set rsSeg = Nothing
End Sub
