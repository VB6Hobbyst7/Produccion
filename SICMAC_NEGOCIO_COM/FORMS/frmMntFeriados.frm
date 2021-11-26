VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntFeriados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Feriados"
   ClientHeight    =   4920
   ClientLeft      =   1755
   ClientTop       =   3090
   ClientWidth     =   7020
   Icon            =   "frmMntFeriados.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   105
      TabIndex        =   1
      Top             =   4170
      Width           =   6810
      Begin VB.CommandButton CmdMuestraAgencias 
         Caption         =   "..."
         Height          =   405
         Left            =   2340
         TabIndex        =   10
         ToolTipText     =   "Mostrar Agencias"
         Top             =   210
         Width           =   375
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   5610
         TabIndex        =   7
         Top             =   210
         Width           =   1080
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   405
         Left            =   1230
         TabIndex        =   5
         Top             =   210
         Width           =   1080
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   1080
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   5610
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   4500
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin MSDataGridLib.DataGrid DGFeriado 
      Height          =   3930
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6932
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         DataField       =   "dFeriado"
         Caption         =   "Fecha"
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
         DataField       =   "cDescrip"
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
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4185.071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtDescrip 
      Height          =   315
      Left            =   2310
      TabIndex        =   6
      Top             =   3660
      Width           =   4125
   End
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   315
      Left            =   1110
      TabIndex        =   8
      Top             =   3660
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdAgencias 
      Caption         =   "..."
      Height          =   375
      Left            =   6465
      TabIndex        =   9
      ToolTipText     =   "Mostrar Agencias"
      Top             =   3630
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   600
      Left            =   105
      Top             =   3540
      Width           =   6825
   End
End
Attribute VB_Name = "frmMntFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RFeriado As ADODB.Recordset
Dim MatAgencias() As String
Dim RFeriadoAge As ADODB.Recordset

Private Sub HabilitaDatos(ByVal pbHabilita As Boolean)
    DGFeriado.Enabled = Not pbHabilita
    CmdNuevo.Visible = Not pbHabilita
    CmdEliminar.Visible = Not pbHabilita
    CmdSalir.Visible = Not pbHabilita
    CmdAceptar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
    If pbHabilita Then
        DGFeriado.Height = 3240
    Else
        DGFeriado.Height = 3930
    End If
    
End Sub

Private Sub CargaGrid()
Dim oDFeriado As DFeriado
    Set oDFeriado = New DFeriado
    Set RFeriado = oDFeriado.RecuperaFeriado
    Set oDFeriado = Nothing
    Set DGFeriado.DataSource = RFeriado
End Sub

Private Sub CmdAceptar_Click()
Dim sMovAct As String
Dim oBase As DCredActualizaBD
Dim i As Integer

    If UBound(MatAgencias) = 0 Then
        MsgBox "Ingrese las Agencias a las cuales se va a asignar el feriado"
        Exit Sub
    End If

    If MsgBox("Se va a Ingresar la Fecha : " & TxtFecha.Text & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    RFeriado.MoveFirst
    RFeriado.Find "dFeriado = '" & Format(CDate(TxtFecha.Text), "dd/mm/yyyy") & "'"
    If Not RFeriado.EOF Then
        MsgBox "Feriado ya Existe", vbInformation, "Aviso"
        Exit Sub
    End If
        
    sMovAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, Mid(gsCodAge, 1, 3), gsCodAge)
    Set oBase = New DCredActualizaBD
    Call oBase.dInsertFeriado(CDate(TxtFecha.Text), Trim(TxtDescrip.Text), sMovAct, False)
    
    oBase.coConex.ConexionActiva.Execute "delete FeriadoAge where dferiado='" & Format(CDate(TxtFecha.Text), "mm/dd/yyyy") & "')"
    
    For i = 0 To UBound(MatAgencias) - 1
       
        If MatAgencias(2) = "1" Then
            Call oBase.dInsertFeriadoAge(CDate(TxtFecha.Text), MatAgencias(0))
        End If
    Next i
    
    Set oBase = Nothing
    Call CargaGrid
    RFeriado.Find "dFeriado = '" & Format(CDate(TxtFecha.Text), "mm/dd/yyyy") & "'"
    If RFeriado.EOF Then
        RFeriado.MoveFirst
    End If
    HabilitaDatos False
End Sub

Private Sub CmdAgencias_Click()
Dim oDFeriado As DFeriado
Dim i As Integer
    Set oDFeriado = New DFeriado
    Set RFeriadoAge = oDFeriado.RecuperaFeriadoAgencias(CDate(TxtFecha.Text))
    Set oDFeriado = Nothing
    ReDim MatAgencias(RFeriadoAge.RecordCount, 3)
    i = 0
    Do While Not RFeriadoAge.EOF
        MatAgencias(i, 0) = RFeriadoAge!cAgeCod
        MatAgencias(i, 1) = RFeriadoAge!cAgeDescripcion
        MatAgencias(i, 2) = RFeriadoAge!Valor
        RFeriadoAge.MoveNext
    Loop
    RFeriadoAge.Close
    Set RFeriadoAge = Nothing
    
    MatAgencias = frmMntFeriadoAge.CargaFlex(MatAgencias, True)
    
End Sub

Private Sub cmdCancelar_Click()
    HabilitaDatos False
End Sub

Private Sub cmdeliminar_Click()
Dim oBase As DCredActualizaBD

    DGFeriado.Col = 0
    If MsgBox("Se va a Eliminar el Dia : " & DGFeriado.Text & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oBase = New DCredActualizaBD
    Call oBase.dDeleteFeriado(CDate(DGFeriado.Text), False)
    Set oBase = Nothing
    Call CargaGrid
End Sub

Private Sub CmdMuestraAgencias_Click()
Dim oDFeriado As DFeriado
Dim i As Integer
    Set oDFeriado = New DFeriado
    Set RFeriadoAge = oDFeriado.RecuperaFeriadoAgencias(RFeriado!DFeriado)
    Set oDFeriado = Nothing
    ReDim MatAgencias(RFeriadoAge.RecordCount, 3)
    i = 0
    Do While Not RFeriadoAge.EOF
        MatAgencias(i, 0) = RFeriadoAge!cAgeCod
        MatAgencias(i, 1) = RFeriadoAge!cAgeDescripcion
        MatAgencias(i, 2) = RFeriadoAge!Valor
        RFeriadoAge.MoveNext
    Loop
    RFeriadoAge.Close
    Set RFeriadoAge = Nothing
    
    MatAgencias = frmMntFeriadoAge.CargaFlex(MatAgencias, False)
    
End Sub

Private Sub cmdNuevo_Click()
    HabilitaDatos True
    TxtFecha.Text = "__/__/____"
    TxtDescrip.Text = ""
    TxtFecha.SetFocus
    ReDim MatAgencias(0)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaGrid
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub TxtDescrip_GotFocus()
    fEnfoque TxtDescrip
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptar.SetFocus
    End If
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque TxtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtDescrip.SetFocus
    End If
End Sub
