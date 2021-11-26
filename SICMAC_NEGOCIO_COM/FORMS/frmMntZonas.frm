VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntZonas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Zonas"
   ClientHeight    =   5835
   ClientLeft      =   1905
   ClientTop       =   1995
   ClientWidth     =   8205
   Icon            =   "frmMntZonas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   825
      Left            =   75
      TabIndex        =   6
      Top             =   4980
      Width           =   8025
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   435
         Left            =   6645
         TabIndex        =   12
         Top             =   210
         Width           =   1230
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   6645
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   435
         Left            =   5400
         TabIndex        =   10
         Top             =   210
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   435
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   435
         Left            =   1395
         TabIndex        =   8
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   435
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4290
      Left            =   75
      TabIndex        =   2
      Top             =   690
      Width           =   8055
      Begin MSDataGridLib.DataGrid DGZonas 
         Height          =   3990
         Left            =   135
         TabIndex        =   3
         Top             =   210
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   7038
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   -2147483630
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "cUbiGeoCod"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cUbiGeoDescripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   1860.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5114.835
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   315
         Left            =   915
         MaxLength       =   12
         TabIndex        =   4
         Top             =   3795
         Width           =   1365
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   3795
         Width           =   5085
      End
      Begin VB.Shape Shape1 
         Height          =   510
         Left            =   135
         Top             =   3690
         Width           =   7785
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   8085
      Begin VB.CommandButton CmdRetro 
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7005
         TabIndex        =   15
         Top             =   225
         Width           =   480
      End
      Begin VB.CommandButton CmdAvanza 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7515
         TabIndex        =   14
         Top             =   225
         Width           =   480
      End
      Begin VB.Label LblZona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   2535
         TabIndex        =   13
         Top             =   210
         Width           =   4395
      End
      Begin VB.Label LblTitulo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PAIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   210
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmMntZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nNivel As Integer
Private nAccion As Integer
Private RZonas As ADODB.Recordset
Private sValor As String
Private sTitulo(5) As String

Private Sub LimpiaDatos()
    TxtCodigo.Text = ""
    TxtDescripcion.Text = ""
End Sub

Private Sub HabilitaDatos(ByVal pbHabilita As Boolean)
    If pbHabilita Then
        DGZonas.Height = 3405
    Else
        DGZonas.Height = 3990
    End If
    DGZonas.Enabled = Not pbHabilita
    TxtCodigo.Enabled = pbHabilita
    TxtDescripcion.Enabled = pbHabilita
    CmdNuevo.Visible = Not pbHabilita
    CmdEditar.Visible = Not pbHabilita
    CmdEliminar.Visible = Not pbHabilita
    CmdAceptar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
    CmdSalir.Visible = Not pbHabilita
    
End Sub

Private Sub CmdAceptar_Click()
Dim oNZonas As COMDConstantes.DCOMZonas
Dim DZonas As COMDConstantes.DCOMZonas
    If MsgBox("Se va a Grabar los Datos, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oNZonas = New COMDConstantes.DCOMZonas
        Call oNZonas.ActualizaZonas(TxtCodigo.Text, TxtDescripcion.Text, nAccion)
    Set oNZonas = Nothing
    HabilitaDatos (False)
    Set DZonas = New COMDConstantes.DCOMZonas
    Set RZonas = DZonas.DameZonas(nNivel, sValor)
    Set DZonas = Nothing
    Set DGZonas.DataSource = RZonas
End Sub

Private Sub CmdAvanza_Click()
Dim DZonas As COMDConstantes.DCOMZonas
        
    nNivel = nNivel + 1
    Select Case nNivel
        Case 0
            LblTitulo.Caption = "TODOS"
        Case 1
            LblTitulo.Caption = "PAIS"
        Case 2
            LblTitulo.Caption = "DEPARTAMENTO"
        Case 3
            LblTitulo.Caption = "PROVINCIA"
        Case 4
            LblTitulo.Caption = "DISTRITO"
    End Select
    If nNivel > 4 Then
        nNivel = 4
    Else
        If nNivel > 1 Then
            sValor = RZonas!cUbiGeoCod
            LblZona.Caption = Trim(RZonas!cUbiGeoDescripcion)
        Else
            LblZona.Caption = "PERU"
        End If
        If nNivel = 0 Then
            sTitulo(0) = "TODOS"
        Else
            sTitulo(nNivel) = LblZona.Caption
        End If
        Set DZonas = New COMDConstantes.DCOMZonas
        If nNivel > 1 Then
            Set RZonas = DZonas.DameZonas(nNivel, RZonas!cUbiGeoCod)
        Else
            Set RZonas = DZonas.DameZonas(nNivel, "")
        End If
        Set DZonas = Nothing
        Set DGZonas.DataSource = RZonas
    End If
    
End Sub

Private Sub CmdCancelar_Click()
    HabilitaDatos False
    LimpiaDatos
End Sub

Private Sub CmdEditar_Click()
    Call HabilitaDatos(True)
    TxtCodigo.Text = RZonas!cUbiGeoCod
    TxtCodigo.Enabled = False
    TxtDescripcion.Text = Trim(RZonas!cUbiGeoDescripcion)
    TxtDescripcion.SetFocus
    nAccion = 2
End Sub

Private Sub cmdeliminar_Click()
Dim oBase As COMDCredito.DCOMCredActBD
Dim DZonas As COMDConstantes.DCOMZonas
    DGZonas.Col = 1
    If MsgBox("Se va a Eliminar la Zona " & Trim(DGZonas.Text) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    DGZonas.Col = 0
    Set oBase = New COMDCredito.DCOMCredActBD
    Call oBase.dDeleteZonas(DGZonas.Text, False)
    Set oBase = Nothing
    Set DZonas = New COMDConstantes.DCOMZonas
    Set RZonas = DZonas.DameZonas(nNivel, sValor)
    Set DZonas = Nothing
    Set DGZonas.DataSource = RZonas
End Sub

Private Sub cmdNuevo_Click()
Dim oDZonas As COMDConstantes.DCOMZonas

    Call HabilitaDatos(True)
    Call LimpiaDatos
    nAccion = 1
    Set oDZonas = New COMDConstantes.DCOMZonas
    TxtCodigo.Text = oDZonas.DameMaximoValorZona(sValor, nNivel)
    Set oDZonas = Nothing
End Sub

Private Sub CmdRetro_Click()
Dim DZonas As COMDConstantes.DCOMZonas
    nNivel = nNivel - 1
    Select Case nNivel
        Case 0
            LblTitulo.Caption = "TODOS"
        Case 1
            LblTitulo.Caption = "PAIS"
        Case 2
            LblTitulo.Caption = "DEPARTAMENTO"
        Case 3
            LblTitulo.Caption = "PROVINCIA"
    End Select
    If nNivel < 0 Then
        nNivel = 0
    Else
        If RZonas.RecordCount > 0 Then
            sValor = RZonas!cUbiGeoCod
            LblZona.Caption = Trim(RZonas!cUbiGeoDescripcion)
            Set DZonas = New COMDConstantes.DCOMZonas
            Set RZonas = DZonas.DameZonas(nNivel, RZonas!cUbiGeoCod)
            Set DZonas = Nothing
            Set DGZonas.DataSource = RZonas
        Else
            Set DZonas = New COMDConstantes.DCOMZonas
            Set RZonas = DZonas.DameZonas(nNivel, sValor)
            Set DZonas = Nothing
            Set DGZonas.DataSource = RZonas
        End If
    End If
    LblZona.Caption = sTitulo(nNivel)
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    nNivel = 0
    Call CmdAvanza_Click
End Sub

Private Sub txtcodigo_GotFocus()
    fEnfoque TxtCodigo
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtDescripcion.SetFocus
    End If
End Sub

Private Sub TxtDescripcion_GotFocus()
    fEnfoque TxtDescripcion
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub
