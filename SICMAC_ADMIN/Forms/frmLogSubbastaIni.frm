VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogSubbastaIni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmLogSubbastaIni.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   6240
      TabIndex        =   11
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   7485
      TabIndex        =   10
      Top             =   5790
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8715
      TabIndex        =   9
      Top             =   5790
      Width           =   1155
   End
   Begin VB.Frame fraSubasta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Subasta"
      ForeColor       =   &H00800000&
      Height          =   1080
      Left            =   90
      TabIndex        =   1
      Top             =   15
      Width           =   9795
      Begin VB.TextBox txtSubasta 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   1575
         TabIndex        =   12
         Top             =   255
         Width           =   8100
      End
      Begin Sicmact.TxtBuscar txtSubastraCod 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   255
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
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
      End
   End
   Begin TabDlg.SSTab sTab 
      Height          =   4530
      Left            =   90
      TabIndex        =   0
      Top             =   1185
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   7990
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Miembros"
      TabPicture(0)   =   "frmLogSubbastaIni.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraMiembros"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes"
      TabPicture(1)   =   "frmLogSubbastaIni.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraBB"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraBB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4020
         Left            =   90
         TabIndex        =   4
         Top             =   375
         Width           =   9570
         Begin Sicmact.FlexEdit flexBB 
            Height          =   3735
            Left            =   75
            TabIndex        =   8
            Top             =   195
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   6588
            Cols0           =   13
            EncabezadosNombres=   $"frmLogSubbastaIni.frx":0342
            EncabezadosAnchos=   "400-1200-1200-1200-4000-1200-1200-1200-1200-1200-1200-1200-1200"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-R-R-L-R-R-R-C-R-R-C-R"
            FormatosEdit    =   "0-0-3-0-0-0-2-2-2-2-2-2-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraMiembros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4035
         Left            =   -74910
         TabIndex        =   3
         Top             =   375
         Width           =   9570
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   1155
            TabIndex        =   7
            Top             =   3570
            Width           =   1005
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   90
            TabIndex        =   6
            Top             =   3570
            Width           =   1005
         End
         Begin Sicmact.FlexEdit flexMiembros 
            Height          =   3285
            Left            =   75
            TabIndex        =   5
            Top             =   195
            Width           =   9420
            _ExtentX        =   16616
            _ExtentY        =   5794
            Cols0           =   4
            EncabezadosNombres=   "#-Codigo-Nombre-Relacion"
            EncabezadosAnchos=   "300-1200-4200-3000"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0-3"
            EncabezadosAlineacion=   "C-L-L-L"
            FormatosEdit    =   "0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
         End
      End
   End
End
Attribute VB_Name = "frmLogSubbastaIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String

Public Sub Ini(psOpeCod As String, psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub

Public Function Valida() As Boolean
    Dim I As Integer
    Dim lbPresidente As Boolean
    Dim lbMartllero As Boolean
    Dim lbObservador As Boolean
    
    If Me.txtSubasta.Text = "" Then
        MsgBox "Debe Ingresar un comentario.", vbInformation
        Me.txtSubasta.SetFocus
        Valida = False
        Exit Function
    End If
    
    lbPresidente = True
    lbMartllero = True
    lbObservador = True
    
    For I = 1 To Me.flexMiembros.Rows - 1
        If Trim(Right(flexMiembros.TextMatrix(I, 3), 5)) = "1" Then
            lbPresidente = False
        ElseIf Trim(Right(flexMiembros.TextMatrix(I, 3), 5)) = "5" Then
            lbMartllero = False
        ElseIf Trim(Right(flexMiembros.TextMatrix(I, 3), 5)) = "4" Then
            lbObservador = False
        End If
    Next I
    
    If lbPresidente Then
        MsgBox "Debe Ingresar al presidente de la comision de venta.", vbInformation
        Me.txtSubasta.SetFocus
        Valida = False
        Exit Function
    ElseIf lbObservador Then
        MsgBox "Debe Ingresar al observador de la comision de venta.", vbInformation
        Me.txtSubasta.SetFocus
        Valida = False
        Exit Function
    ElseIf lbMartllero Then
        MsgBox "Debe Ingresar al martillero.", vbInformation
        Me.txtSubasta.SetFocus
        Valida = False
        Exit Function
    End If
End Function

Private Sub cmdAgregar_Click()
    flexMiembros.AdicionaFila
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro " & Me.flexMiembros.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    flexMiembros.EliminaFila flexMiembros.Row
End Sub

Private Sub cmdGrabar_Click()
    If Not Valida Then Exit Sub
    
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
        
        
    
    
End Sub

Private Sub cmdNuevo_Click()
    Dim oSubasta As DSubasta
    Set oSubasta = New DSubasta
    
    Me.txtSubastraCod.Text = oSubasta.GetCodigo(Format(gdFecSis, "yyyy"))
    Me.txtSubasta.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    
    flexMiembros.CargaCombo oCon.GetConstante(5009, , , , , , True)
    
    Caption = lsCaption
    sTab.Tab = 0
End Sub


