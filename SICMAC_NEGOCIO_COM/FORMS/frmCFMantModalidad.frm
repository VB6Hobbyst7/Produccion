VERSION 5.00
Begin VB.Form frmCFMantModalidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas Fianza - Modalidades"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   Icon            =   "frmCFMantModalidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraModalidad 
      Caption         =   "Nueva Modalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   8025
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6720
         TabIndex        =   10
         Top             =   720
         Width           =   1170
      End
      Begin VB.CheckBox chkEstado 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtModalidad 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   6855
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         TabIndex        =   5
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado :"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblModalidad 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Top             =   4080
      Width           =   1170
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1170
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   0
      Top             =   4080
      Width           =   1170
   End
   Begin SICMACT.FlexEdit feModalidades 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8040
      _extentx        =   14182
      _extenty        =   6376
      cols0           =   4
      highlight       =   1
      encabezadosnombres=   "#-Modalidad-Estado-Aux"
      encabezadosanchos=   "500-6000-1100-0"
      font            =   "frmCFMantModalidad.frx":030A
      font            =   "frmCFMantModalidad.frx":0332
      font            =   "frmCFMantModalidad.frx":035A
      font            =   "frmCFMantModalidad.frx":0382
      font            =   "frmCFMantModalidad.frx":03AA
      fontfixed       =   "frmCFMantModalidad.frx":03D2
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      tipobusqueda    =   3
      columnasaeditar =   "X-1-2-X"
      listacontroles  =   "0-0-4-0"
      encabezadosalineacion=   "L-L-C-C"
      formatosedit    =   "0-0-0-4"
      textarray0      =   "#"
      lbeditarflex    =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   495
      rowheight0      =   300
   End
End
Attribute VB_Name = "frmCFMantModalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim i As Integer
'Dim fsMensaje As String
'Dim fnTipo As Integer
'Dim fnCodigo As Integer
'
'Private Sub cmdAgregar_Click()
'fnTipo = 1
'txtModalidad.Text = ""
'chkEstado.value = 0
'cmdEditar.Enabled = False
'cmdCerrar.Enabled = False
'fnCodigo = 0
'Call OperacionesModalidad
'End Sub
'
'Private Sub cmdCancelar_Click()
'fraModalidad.Top = 4800
'fraModalidad.Left = 120
'cmdAgregar.Enabled = True
'cmdCerrar.Enabled = True
'cmdEditar.Enabled = True
'txtModalidad.Text = ""
'chkEstado.value = 0
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub CmdEditar_Click()
'fnTipo = 2
'cmdAgregar.Enabled = False
'cmdCerrar.Enabled = False
'txtModalidad.Text = feModalidades.TextMatrix(feModalidades.row, 1)
'chkEstado.value = IIf(feModalidades.TextMatrix(feModalidades.row, 2) = ".", 1, 0)
'fnCodigo = CInt(feModalidades.TextMatrix(feModalidades.row, 3))
'Call OperacionesModalidad
'End Sub
'
'Private Sub cmdGrabar_Click()
'Dim oDCartaFianza As COMDCartaFianza.DCOMCartaFianza
'Set oDCartaFianza = New COMDCartaFianza.DCOMCartaFianza
'
'Call oDCartaFianza.ObtenerModalidadesCF("0,1")
'If validaDatos Then
'    If MsgBox("Estas seguro de " & fsMensaje & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'
'        Call oDCartaFianza.GrabarModalidadesCF(fnTipo, Trim(txtModalidad.Text), chkEstado.value, fnCodigo)
'
'        MsgBox "Modalidad Grabada Satisfactoriamente.", vbInformation, "Aviso"
'
'        fraModalidad.Top = 4800
'        fraModalidad.Left = 120
'        cmdAgregar.Enabled = True
'        cmdCerrar.Enabled = True
'        cmdEditar.Enabled = True
'        txtModalidad.Text = ""
'        chkEstado.value = 0
'        Call CargarGrid
'        cmdAgregar.SetFocus
'    End If
'End If
'End Sub
'
'Private Function validaDatos() As Boolean
'Dim oDCartaFianza As COMDCartaFianza.DCOMCartaFianza
'Set oDCartaFianza = New COMDCartaFianza.DCOMCartaFianza
'Dim nCantidad As Long
'If Trim(txtModalidad.Text) = "" Then
'    MsgBox "Ingrese la Modalidad de Carta Fianza.", vbInformation, "Aviso"
'    validaDatos = False
'    Exit Function
'End If
'
'For i = 0 To feModalidades.Rows - 2
'    If feModalidades.TextMatrix(i + 1, 3) <> fnCodigo Then
'        If Trim(feModalidades.TextMatrix(i + 1, 1)) = Trim(Me.txtModalidad.Text) Then
'            MsgBox "Modalidad de Carta Fianza ya Existe en la Fila " & (i + 1) & ".", vbInformation, "Aviso"
'            validaDatos = False
'            Exit Function
'        End If
'    End If
'Next i
'
'nCantidad = oDCartaFianza.ExiteModalidaENCF(fnCodigo)
'If nCantidad > 0 Then
'    If MsgBox("Existen " & nCantidad & " Cartas Fianza con esta Modalidad, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
'        validaDatos = False
'        Exit Function
'    End If
'End If
'
'validaDatos = True
'End Function
'Private Sub Form_Load()
'Call CargarGrid
'Me.Height = 5040
'End Sub
'
'Private Sub CargarGrid()
'Dim R As ADODB.Recordset
'Dim oDCartaFianza As COMDCartaFianza.DCOMCartaFianza
'
'Set oDCartaFianza = New COMDCartaFianza.DCOMCartaFianza
'Set R = oDCartaFianza.ObtenerModalidadesCF("0,1")
'feModalidades.lbEditarFlex = False
'Call LimpiaFlex(feModalidades)
'
'If R.RecordCount > 1 Then
'    If Not (R.EOF And R.BOF) Then
'        For i = 0 To R.RecordCount - 1
'            feModalidades.AdicionaFila
'            feModalidades.TextMatrix(i + 1, 0) = i + 1
'            feModalidades.TextMatrix(i + 1, 1) = Trim(R!Descripcion)
'            feModalidades.TextMatrix(i + 1, 2) = IIf(R!Estado, 1, 0)
'            feModalidades.TextMatrix(i + 1, 3) = (R!cod)
'            R.MoveNext
'        Next i
'    End If
'End If
'
'End Sub
'
'Private Sub OperacionesModalidad()
'If fnTipo = 1 Then
'    fraModalidad.Caption = "Nueva Modalidad"
'    fsMensaje = "Grabar la Nueva Modalidad"
'ElseIf fnTipo = 2 Then
'    fraModalidad.Caption = "Editar Modalidad"
'    fsMensaje = "Editar la Modalidad"
'End If
'fraModalidad.Top = 3300
'fraModalidad.Left = 120
'
'i = Len(Trim(txtModalidad.Text))
'txtModalidad.SelStart = i
'txtModalidad.SetFocus
'End Sub
'
'Private Sub txtModalidad_Change()
'If txtModalidad.SelStart > 0 Then
'    i = Len(Mid(txtModalidad.Text, 1, txtModalidad.SelStart))
'End If
'txtModalidad.Text = UCase(txtModalidad.Text)
'txtModalidad.SelStart = i
'End Sub
'
'Private Sub txtModalidad_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdGrabar.SetFocus
'End If
'End Sub
