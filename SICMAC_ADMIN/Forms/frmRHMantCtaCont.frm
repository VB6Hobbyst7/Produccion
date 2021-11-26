VERSION 5.00
Begin VB.Form frmRHMantCtaCont 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmRHMantCtaCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   60
      TabIndex        =   7
      Top             =   4485
      Width           =   1110
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   4485
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7560
      TabIndex        =   3
      Top             =   4485
      Width           =   1110
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   1230
      TabIndex        =   2
      Top             =   4485
      Width           =   1110
   End
   Begin Sicmact.TxtBuscar txtPlanilla 
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      Appearance      =   0
      BackColor       =   12648447
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
      sTitulo         =   ""
   End
   Begin VB.Frame fraCon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   60
      TabIndex        =   5
      Top             =   390
      Width           =   8595
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   7455
         TabIndex        =   9
         Top             =   3645
         Width           =   1080
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   6345
         TabIndex        =   8
         Top             =   3645
         Width           =   1080
      End
      Begin Sicmact.FlexEdit Flex 
         Height          =   3375
         Left            =   75
         TabIndex        =   6
         Top             =   225
         Width           =   8460
         _ExtentX        =   15002
         _ExtentY        =   5953
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cod Concepto-Concepto Remunerativo-Debe-Haber"
         EncabezadosAnchos=   "400-1200-4000-1200-1200"
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
         ColumnasAEditar =   "X-1-X-3-4"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-R-R"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Label lblPlanilla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1665
      TabIndex        =   0
      Top             =   105
      Width           =   7005
   End
End
Attribute VB_Name = "frmRHMantCtaCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String

Public Sub Ini(psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Activa False
    txtPlanilla.Text = ""
    txtPlanilla_EmiteDatos
End Sub

Private Sub cmdEditar_Click()
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro numero : " & Me.Flex.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Flex.EliminaFila Flex.Row
End Sub

Private Sub cmdGrabar_Click()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    If Not Valida Then Exit Sub
    
    oPla.SetConceptoCta txtPlanilla.Text, Flex.GetRsNew
    
    cmdCancelar_Click
End Sub

Private Sub cmdNuevo_Click()
    Flex.AdicionaFila
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    If pnCol = 3 Then
        If Flex.TextMatrix(pnRow, 4) <> "" And Flex.TextMatrix(pnRow, 3) <> "" Then
            MsgBox "Solo puede ingresar por registro una cuenta contable al debe o al haber. Si desea agregar otra cuenta ingres un nuvo registro.", vbInformation, "Aviso"
            Cancel = False
        End If
    Else
        If Flex.TextMatrix(pnRow, 3) <> "" And Flex.TextMatrix(pnRow, 4) <> "" Then
            MsgBox "Solo puede ingresar por registro una cuenta contable al debe o al haber. Si desea agregar otra cuenta ingres un nuvo registro.", vbInformation, "Aviso"
            Cancel = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    Caption = lsCaption
    
    txtPlanilla.rs = oPla.GetPlanillasOpeCon
    Me.Flex.rsTextBuscar = oPla.GetConceptoTablaArbol
    
    Set oPla = Nothing
    
    Activa False
End Sub

Private Sub txtPlanilla_EmiteDatos()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    Me.lblPlanilla.Caption = txtPlanilla.psDescripcion
    
    If Me.lblPlanilla.Caption = "" Then
        Flex.Clear
        Flex.Rows = 2
        Flex.FormaCabecera
    Else
        Flex.rsFlex = oPla.GetConceptoCta(Me.txtPlanilla.Text)
    End If
End Sub

Private Sub Activa(pbEditar As Boolean)
    Me.txtPlanilla.Enabled = Not pbEditar
    Me.fraCon.Enabled = pbEditar
    Me.cmdCancelar.Visible = pbEditar
    Me.cmdEditar.Visible = Not pbEditar
    Me.cmdGrabar.Enabled = pbEditar
End Sub

Private Function Valida() As Boolean
    Dim lnI As Integer
    
    If Me.txtPlanilla.Text = "" Then
        Valida = False
        MsgBox "Debe hacer referencia a una planilla valida.", vbInformation, "Aviso"
        Exit Function
    End If
    
    For lnI = 1 To Me.Flex.Rows - 1
        Flex.Row = lnI
        If Flex.TextMatrix(lnI, 1) = "" Then
            Flex.Col = 1
            MsgBox "Debe ingresar un codigo valido.", vbInformation, "Aviso"
            Flex.SetFocus
            Exit Function
        ElseIf Flex.TextMatrix(lnI, 3) = "" And Flex.TextMatrix(lnI, 4) = "" Then
            Flex.Col = 3
            MsgBox "Debe ingresar una cuenta contable valida, para el debe o el haber.", vbInformation, "Aviso"
            Flex.SetFocus
            Exit Function
        End If
    Next lnI
    
    Valida = True
End Function
