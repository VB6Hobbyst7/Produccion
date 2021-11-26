VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHAsignacionPlan 
   Caption         =   "Asignacion de Conceptos Plan EPS"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   Icon            =   "frmRHAsignacionPlan.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   11835
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Listado"
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   435
      Width           =   1335
   End
   Begin MSMask.MaskEdBox meEmpresa 
      Height          =   285
      Left            =   10080
      TabIndex        =   28
      Top             =   4560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   4
      Format          =   "0.##"
      Mask            =   "0.##"
      PromptChar      =   "0"
   End
   Begin VB.TextBox txtMes 
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtano 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtDHabientes 
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtPromedioEPS 
      Height          =   285
      Left            =   7320
      TabIndex        =   19
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txt25EPS 
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   9240
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Fdetalle 
      Caption         =   "Planes Asignadosal Trabajador"
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
      Height          =   3015
      Left            =   6120
      TabIndex        =   9
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   5295
      End
      Begin Sicmact.FlexEdit flexPerPlan 
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4471
         Cols0           =   3
         AllowUserResizing=   1
         EncabezadosNombres=   "#-CodPlanEPS-Descripcion"
         EncabezadosAnchos=   "500-1200-3000"
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
         ColumnasAEditar =   "X-1-X"
         ListaControles  =   "0-1-0"
         EncabezadosAlineacion=   "C-C-L"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
      End
   End
   Begin VB.TextBox txtCantidadPer 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "0"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtMontoTotalEPS 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Text            =   "0"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHLista 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin MSMask.MaskEdBox meQuincena 
      Height          =   285
      Left            =   10080
      TabIndex        =   29
      Top             =   5040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   4
      Format          =   "0.##"
      Mask            =   "0.##"
      PromptChar      =   "0"
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "% Empresa"
      Height          =   195
      Left            =   9000
      TabIndex        =   27
      Top             =   4560
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "%Quincena"
      Height          =   195
      Left            =   9000
      TabIndex        =   26
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Der Habientes"
      Height          =   195
      Left            =   6240
      TabIndex        =   23
      Top             =   4560
      Width           =   1020
   End
   Begin VB.Label lblmes 
      AutoSize        =   -1  'True
      Caption         =   "Mes"
      Height          =   195
      Left            =   1320
      TabIndex        =   21
      Top             =   480
      Width           =   300
   End
   Begin VB.Label lblaño 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Promedio EPS"
      Height          =   195
      Left            =   6240
      TabIndex        =   18
      Top             =   5565
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "2.25%EPS"
      Height          =   195
      Left            =   6240
      TabIndex        =   16
      Top             =   5085
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
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
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   6330
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Monto  Total"
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
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      Top             =   6330
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Personal Afiliado a EPS"
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
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2970
   End
End
Attribute VB_Name = "frmRHAsignacionPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oAsisMedica As DActualizaAsistMedicaPrivada
Dim lbEdita As Boolean
Dim rs As New ADODB.Recordset

Private Sub CmdCalcular_Click()
frmRHCalculoEPS.Show
End Sub

Private Sub cmdCancelar_Click()
 Activa False
 MSHLista.Enabled = True
 flexPerPlan.Enabled = False
 Fdetalle.Enabled = False
 
 
 MSHLista_Click
 
End Sub

Private Sub cmdEditar_Click()
lbEdita = True
Activa True
MSHLista.Enabled = False
flexPerPlan.Enabled = True
Fdetalle.Enabled = True
End Sub


Private Sub cmdEliminar_Click()
If MsgBox("¿ Desea Eliminar el Registro ?", vbQuestion + vbYesNo, "Se Eliminar el Registro") = vbNo Then Exit Sub
 flexPerPlan.EliminaFila flexPerPlan.Row
End Sub


Private Sub cmdGrabar_Click()

Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sCodPersona As String
Set rs = Me.flexPerPlan.GetRsNew
sCodPersona = Me.MSHLista.TextMatrix(Me.MSHLista.Row, 0)
If sCodPersona = "" Then
   MsgBox "Debe Seleccionar una Persona a la cual Modificar o Agregar un plan EPS", vbInformation, "Aviso"
   Exit Sub
End If
If lbEdita Then
        'Modifica0
         If flexPerPlan.Rows = 2 And flexPerPlan.TextMatrix(1, 0) = "" Then
                oAsisMedica.EliminaRHPersonaPlan sCodPersona
         Else
                oAsisMedica.AgregaRHPersonaPlan rs, sCodPersona, GetMovNro(gsCodUser, gsCodAge)
         End If
    Else
        'Nuevo
        oAsisMedica.AgregaRHPersonaPlan rs, sCodPersona, GetMovNro(gsCodUser, gsCodAge)
        
End If
    Set oAsistencia = Nothing
    lbEdita = False
    Activa False
    MSHLista.Enabled = True
    flexPerPlan.Enabled = False
    Fdetalle.Enabled = False
End Sub

Private Sub cmdMostrar_Click()
Listado txtano + txtMes
End Sub

Private Sub CmdNuevo_Click()
    lbEdita = False
    Activa True
    Me.flexPerPlan.AdicionaFila
    Me.flexPerPlan.Enabled = True
    Me.flexPerPlan.SetFocus
    MSHLista.Enabled = False
    flexPerPlan.Enabled = True
    
    
End Sub

Private Sub Activa(pbvalor As Boolean)
    
    'Me.Flex.Enabled = Not pbValor
    'Me.cmdNuevo.Visible = Not pbValor
    'Me.cmdEditar.Visible = Not pbValor
    'Me.cmdGrabar.Visible = pbValor
    'Me.cmdCancelar.Visible = pbValor
    'Me.Flex.Enabled = Not pbValor
    'Me.cmdNuevo.Enabled = Not pbValor
    Me.cmdEditar.Enabled = Not pbvalor
    Me.cmdGrabar.Enabled = pbvalor
    Me.cmdCancelar.Enabled = pbvalor
    
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub flexPerPlan_RowColChange()
 If Me.flexPerPlan.Col = 1 Then
    Me.flexPerPlan.rsTextBuscar = oAsisMedica.GetRHCatPlanEPS
 End If

End Sub

Private Sub Form_Load()
   Me.Width = 12000
   Me.Height = 7215
   Set oAsisMedica = New DActualizaAsistMedicaPrivada
   Set rs = New ADODB.Recordset
   MSHLista.ColWidth(0) = 1300
   MSHLista.ColWidth(1) = 2500
   MSHLista.ColWidth(2) = 1250
   MSHLista.ColWidth(3) = 500
   txtano.Text = Year(gdFecSis)
   txtMes.Text = Format(Month(gdFecSis), "00")
   Listado txtano.Text + txtMes.Text
   
   
   meQuincena.Text = "0.50"
   meEmpresa.Text = "0.50"
   
End Sub



Private Sub MSHLista_Click()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
txtNombre.Text = UCase(MSHLista.TextMatrix(MSHLista.Row, 1))

Set rs = oAsisMedica.GetRHPersonaPlan(MSHLista.TextMatrix(MSHLista.Row, 0))
If rs.EOF Then
    
    flexPerPlan.Clear
    flexPerPlan.Rows = 2
    flexPerPlan.FormaCabecera
Else
    Set flexPerPlan.Recordset = rs
    flexPerPlan.FormaCabecera
End If


End Sub

Private Sub MSHLista_KeyDown(KeyCode As Integer, Shift As Integer)
MSHLista_Click
End Sub

Sub Listado(psPeriodo As String)
 Set MSHLista.DataSource = oAsisMedica.GetRHListadoEPS(1, psPeriodo)
 If MSHLista.Rows = 1 Then
   Me.txtMontoTotalEPS.Text = 0
   Me.txtCantidadPer.Text = 0
   Me.txt25EPS.Text = 0
   Me.txtDHabientes.Text = 0
   txtPromedioEPS.Text = 0

 Exit Sub
 End If
 If MSHLista.TextMatrix(MSHLista.Rows - 1, 2) <> "" Then
        MSHLista.Rows = MSHLista.Rows + 1
        MSHLista.TextMatrix(MSHLista.Rows - 1, 0) = "TOTAL"
        For J = 2 To Me.MSHLista.Cols - 1
            lnAcumulador = 0
            If Left(MSHLista.TextMatrix(0, J), 2) <> "U_" And Left(MSHLista.TextMatrix(0, J), 1) <> "_" Then
                For i = 1 To Me.MSHLista.Rows - 2
                    If MSHLista.TextMatrix(i, J) <> "" Then
                        lnAcumulador = lnAcumulador + CCur(MSHLista.TextMatrix(i, J))
                        'If MSHLista.TextMatrix(i, 3) = "L" Then
                        '    MSHLista.Row = i
                        '    MSHLista.Col = J
                        '    MSHLista.CellBackColor = &HA0C000
                        'End If
                    End If
                Next i
                MSHLista.TextMatrix(MSHLista.Rows - 1, J) = Format(lnAcumulador, "#,##.00")
                MSHLista.Row = MSHLista.Rows - 1
                MSHLista.Col = J
                MSHLista.CellBackColor = &HA0C000
                'FlexPrePla.CellFontBold = True
                lnAcumulador = lnAcumulador + CCur(MSHLista.TextMatrix(i, J))
                
            End If
        Next J
    End If
MSHLista.TextMatrix(MSHLista.Rows - 1, 1) = MSHLista.Rows - 2


   Set rs = oAsisMedica.GetRHListadoEPS(3, txtano.Text + txtMes.Text)
   Me.txtMontoTotalEPS.Text = rs!Monto
   Set rs = oAsisMedica.GetRHListadoEPS(2, txtano.Text + txtMes.Text)
   Me.txtCantidadPer.Text = rs!cant
   Me.txt25EPS.Text = Round((txtMontoTotalEPS.Text * 0.0225), 2)
   Set rs = oAsisMedica.GetRHListadoEPS(4, txtano.Text + txtMes.Text)
   Me.txtDHabientes.Text = rs!cant
   txtPromedioEPS.Text = Format(Round(Me.txt25EPS.Text / Me.txtDHabientes.Text, 2), "####.##")

 
End Sub

