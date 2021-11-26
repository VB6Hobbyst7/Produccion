VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPreMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuesto: Mantenimiento"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "frmPreMantenimiento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDuplicar 
      Cancel          =   -1  'True
      Caption         =   "&Duplicar"
      Height          =   390
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3495
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3495
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   6180
      TabIndex        =   8
      Top             =   3495
      Width           =   1110
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3495
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Frame fraContenedor 
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   2340
      Width           =   7275
      Begin VB.ComboBox cboTpo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   630
         Width           =   1635
      End
      Begin VB.TextBox txtCodPre 
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   570
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox txtDesPre 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   1
         Top             =   225
         Width           =   5985
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Tipo :"
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   12
         Top             =   705
         Width           =   645
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   11
         Top             =   285
         Width           =   1005
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgIngreso 
      Height          =   2310
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   4075
      _Version        =   393216
      Cols            =   4
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3495
      Width           =   1110
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   390
      Left            =   1215
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3495
      Width           =   1110
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   390
      Left            =   2355
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3495
      Width           =   1110
   End
End
Attribute VB_Name = "frmPreMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bNuevo As Boolean

Private Sub cmdCancelar_Click()
HabilitaBotones False
End Sub

Private Sub cmdDuplicar_Click()
    Dim lsCodigo As String
    Dim lsCodigoCopia As String
    Dim lnAnio As Long
    Dim lnAnioCopia As Long
    Dim oPresup As DPresupuesto
    Set oPresup = New DPresupuesto

    frmPreTpo.Ini lsCodigo, lnAnio, lnAnioCopia

    If lsCodigo = "" Then
        MsgBox "Debe ingresar un Presupuesto como plantilla valido.", vbInformation, "Aviso"
        Me.cmdDuplicar.SetFocus
        Exit Sub
    End If

    If MsgBox("Desea Duplicar el Presupuesto indicado con el escogido anteriormente ? " & vbEnter & " Si elije duplicar, los rubros que contenga este presupuesto seran borrados y no podra recuperarlos.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub

    lsCodigoCopia = Me.fgIngreso.TextMatrix(fgIngreso.Row, 1)
    oPresup.DuplicaClase lsCodigo, Str(lnAnio), lsCodigoCopia, Str(lnAnioCopia)

    Set oPresup = Nothing
    
    MsgBox "Duplicacion Terminada.", vbInformation, "Aviso"
End Sub

Private Sub cmdEliminar_Click()
Dim oPP As DPresupuesto
If MsgBox("¿ Seguro que desea eliminar Clase de Presupuesto?", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
Set oPP = New DPresupuesto
oPP.EliminaPresupuesto (Me.txtCodPre)
Set oPP = Nothing
fgIngreso.RemoveItem fgIngreso.Row
MuestraDatos
End Sub

Private Sub CmdGrabar_Click()
    Dim tmpSql As String
    Dim sTpo As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    
    txtDesPre.Text = Replace(txtDesPre.Text, "'", "", , , vbTextCompare)
    txtDesPre.Text = UCase(txtDesPre.Text)
    sTpo = Right(Trim(cboTpo.Text), 1)
    
    If Len(Trim(txtDesPre)) = 0 Then
        MsgBox "no se ha ingresado la Descripción", vbInformation, " Aviso "
        Exit Sub
    End If
    
    If sTpo = "" Then
        MsgBox "Seleccione el Tipo de Presupuesto", vbInformation, " Aviso "
        Exit Sub
    End If
  
    If MsgBox("Esta seguro de Grabar esta Clase de Presupuesto", vbQuestion + vbYesNo, " ¡Confirmación! ") = vbYes Then
        If bNuevo = True Then
           'Insertar un nuevo Presupuesto
           oPP.AgregaPresupuesto txtCodPre.Text, txtDesPre.Text, sTpo, False
        Else
           'Actualiza Presupuesto
           oPP.ModificaPresupuesto txtCodPre.Text, txtDesPre.Text, sTpo, False
        End If
        HabilitaBotones False
        txtDesPre.Enabled = False
        cboTpo.Enabled = False
        Call CargaPresu
    End If
    bNuevo = False
End Sub

Private Sub HabilitaBotones(plActiva As Boolean)
   cmdNuevo.Visible = Not plActiva
   cmdModificar.Visible = Not plActiva
   cmdEliminar.Visible = Not plActiva
   cmdGrabar.Visible = plActiva
   cmdCancelar.Visible = plActiva
End Sub

Private Sub cmdModificar_Click()
   HabilitaBotones True
    txtDesPre.Enabled = True
    cboTpo.Enabled = False
    txtDesPre.SetFocus
End Sub

Private Sub cmdNuevo_Click()
    Dim vCod As String
    Dim oPP As DPresupuesto
    
    Set oPP = New DPresupuesto
    vCod = oPP.GetNroPresupuesto
    Set oPP = Nothing
    
    txtCodPre.Text = vCod
    txtDesPre.Text = ""
    HabilitaBotones True
    txtDesPre.Enabled = True
    cboTpo.Enabled = True
    cboTpo.ListIndex = -1
    txtDesPre.SetFocus
    bNuevo = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub MuestraDatos()
   Dim vFil As Integer, nCont As Integer
   vFil = fgIngreso.Row
   If fgIngreso.TextMatrix(vFil, 0) <> "" Then
      txtCodPre.Text = fgIngreso.TextMatrix(vFil, 1)
      txtDesPre.Text = fgIngreso.TextMatrix(vFil, 2)
      For nCont = 0 To cboTpo.ListCount - 1
          cboTpo.ListIndex = nCont
          If Right(Trim(cboTpo.Text), 1) = Right(Trim(fgIngreso.TextMatrix(vFil, 3)), 1) Then
             Exit Sub
          End If
      Next
      cboTpo.ListIndex = -1
   End If
End Sub

Private Sub fgIngreso_Click()
   MuestraDatos
End Sub

Private Sub fgIngreso_RowColChange()
    MuestraDatos
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rsPP As ADODB.Recordset
    Set rsPP = New ADODB.Recordset
    Limpiar
    txtDesPre.Enabled = False
    Call CargaPresu
    'Carga Tpos de Presupuestos
    Set rsPP = oCon.GetConstante(gPPPresupuestoTpo, , , , , gPPPresupuestoTpo)
    CargaCombo rsPP, cboTpo
End Sub

Private Sub Limpiar()
    Call MSHFlex(fgIngreso, 4, "Item-Código-Descripción-Tipo", "450-0-5000-1500", "R-L-L-L")
End Sub

Private Sub CargaPresu()
   Dim tmpReg As ADODB.Recordset
   Set tmpReg = New ADODB.Recordset
   Dim oPP As DPresupuesto
   Set oPP = New DPresupuesto
   Dim tmpSql As String
   Dim x As Integer, n As Integer
   Limpiar

   tmpSql = " SELECT p.cPresu, p.cDesPre, (t.cNomTab + space(40) + p.cTpo) cTpo " & _
            " FROM PPresu P LEFT JOIN " & gcCentralCom & "TablaCod T ON p.cTpo = t.cValor AND t.cCodTab LIKE 'P4__'" & _
            " ORDER BY cPresu "
   
   Set tmpReg = oPP.GetPresupuesto
   
   If Not (tmpReg.BOF Or tmpReg.EOF) Then
      With tmpReg
          Do While Not .EOF
              x = x + 1
              AdicionaRow fgIngreso, x
              fgIngreso.TextMatrix(x, 0) = x
              fgIngreso.TextMatrix(x, 1) = !nPresuCod
              fgIngreso.TextMatrix(x, 2) = !cPresuDescripcion
              fgIngreso.TextMatrix(x, 3) = !cTpo
              .MoveNext
          Loop
      End With
   End If
   tmpReg.Close
   Set tmpReg = Nothing
   Call fgIngreso_RowColChange
End Sub


Private Sub txtDesPre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If cboTpo.Enabled Then
        Me.cboTpo.SetFocus
      Else
         cmdGrabar.SetFocus
      End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub
