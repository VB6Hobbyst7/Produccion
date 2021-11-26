VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmChequeMantenimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Cheques"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   Icon            =   "frmChequeMantenimiento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   340
      Left            =   10320
      TabIndex        =   14
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton btnExportar 
      Caption         =   "&Exportar"
      Height          =   340
      Left            =   2535
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton btnAnular 
      Caption         =   "&Anular"
      Height          =   340
      Left            =   1305
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton btnEliminar 
      Caption         =   "&Eliminar"
      Height          =   340
      Left            =   75
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Cheques"
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
      Height          =   4350
      Left            =   80
      TabIndex        =   15
      Top             =   80
      Width           =   11460
      Begin VB.CheckBox chkTodosSel 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1230
         Width           =   855
      End
      Begin VB.CheckBox chkFechaReg 
         Caption         =   "Todos"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   8280
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkNroCheque 
         Caption         =   "Todos"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   6240
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkIFI 
         Caption         =   "Todos"
         CausesValidation=   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "&Limpiar"
         Height          =   340
         Left            =   10200
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         Height          =   340
         Left            =   10200
         TabIndex        =   9
         Top             =   350
         Width           =   1095
      End
      Begin VB.Frame FraFechaReg 
         Height          =   975
         Left            =   8160
         TabIndex        =   22
         Top             =   240
         Width           =   1995
         Begin MSMask.MaskEdBox txtFechaReg 
            Height          =   300
            Left            =   480
            TabIndex        =   8
            Top             =   480
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Registro"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraNroCheque 
         Height          =   975
         Left            =   6120
         TabIndex        =   19
         Top             =   240
         Width           =   1995
         Begin VB.TextBox txtHasta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            MaxLength       =   8
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtDesde 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            MaxLength       =   8
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   260
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame fraIFI 
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5950
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmChequeMantenimiento.frx":030A
            Left            =   840
            List            =   "frmChequeMantenimiento.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   1690
         End
         Begin Sicmact.TxtBuscar txtIFICod 
            Height          =   270
            Left            =   3360
            TabIndex        =   2
            Top             =   270
            Width           =   2500
            _ExtentX        =   4419
            _ExtentY        =   476
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   280
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cuenta:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblIFINroCuenta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   840
            TabIndex        =   3
            Top             =   620
            Width           =   5010
         End
         Begin VB.Label Label1 
            Caption         =   "I.Financ:"
            Height          =   255
            Left            =   2640
            TabIndex        =   17
            Top             =   280
            Width           =   615
         End
      End
      Begin Sicmact.FlexEdit feCheque 
         Height          =   2730
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4815
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-cIFTpo-cPersCod-cCtaIFCod-nMovNro-nMovFlag-Sel-Institución Financiera-Cuenta-F.Reg-Número-Estado-Fecha Uso-Glosa"
         EncabezadosAnchos=   "400-0-0-0-0-0-400-3200-2200-1000-1000-700-1000-4500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-6-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-4-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-L-L-C-C-C-C-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmChequeMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'** Nombre : frmChequeMantenimiento
'** Descripción : Clase de Cheques  creado según RFC117-2012
'** Creación : EJVG 20121201 09:00:00 AM
'***********************************************************************
Option Explicit
Dim fsopecod As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub Form_Load()
    CentraForm Me
    limpiarCampos
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub
Private Sub btnLimpiar_Click()
    Call limpiarCampos
End Sub
Private Sub cboMoneda_Click()
    Dim oOpe As New DOperacion
    Dim lsMonedaCod As String
    lsMonedaCod = Trim(Right(cboMoneda.Text, 3))
    If lsMonedaCod = "" Then Exit Sub
    If CInt(lsMonedaCod) = 1 Then
        fsopecod = OpeChqRegistroTalonarioMN
    Else
        fsopecod = OpeChqRegistroTalonarioME
    End If
    txtIFICod.rs = oOpe.GetOpeObj(fsopecod, "2")
    Set oOpe = Nothing
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtIFICod.SetFocus
    End If
End Sub
Private Sub chkIFI_Click()
    If Me.chkIFI.value = 1 Then
        fraIFI.Enabled = False
        cboMoneda.ListIndex = -1
        txtIFICod.Text = ""
        lblIFINroCuenta.Caption = ""
    Else
        fraIFI.Enabled = True
    End If
End Sub
Private Sub chkIFI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkIFI.value = 1 Then
            chkNroCheque.SetFocus
        Else
            cboMoneda.SetFocus
        End If
    End If
End Sub
Private Sub chkNroCheque_Click()
    If chkNroCheque.value = 1 Then
        fraNroCheque.Enabled = False
        txtDesde.Text = ""
        txthasta.Text = ""
    Else
        fraNroCheque.Enabled = True
    End If
End Sub
Private Sub chkNroCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkNroCheque.value = 1 Then
            chkFechaReg.SetFocus
        Else
            txtDesde.SetFocus
        End If
    End If
End Sub
Private Sub chkFechaReg_Click()
    If chkFechaReg.value = 1 Then
        FraFechaReg.Enabled = False
        Me.txtFechaReg.Text = "__/__/____"
    Else
        FraFechaReg.Enabled = True
    End If
End Sub
Private Sub chkFechaReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkFechaReg.value = 1 Then
            btnBuscar.SetFocus
        Else
            txtFechaReg.SetFocus
        End If
    End If
End Sub
Private Sub txtIFICod_EmiteDatos()
    Dim oCtaIf As New NCajaCtaIF
    lblIFINroCuenta.Caption = ""
    If txtIFICod.Text <> "" Then
        lblIFINroCuenta.Caption = oCtaIf.NombreIF(Mid(txtIFICod.Text, 4, 13)) & " - " & oCtaIf.EmiteTipoCuentaIF(Mid(txtIFICod.Text, 18, Len(txtIFICod.Text))) & " " & txtIFICod.psDescripcion
    End If
    Set oCtaIf = Nothing
End Sub
Private Sub txtIFICod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkNroCheque.SetFocus
    End If
End Sub
Private Sub txtIFICod_GotFocus()
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la moneda de la operación", vbInformation, "Aviso"
        cboMoneda.SetFocus
    End If
End Sub
Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txthasta.SetFocus
    End If
End Sub
Private Sub txtDesde_LostFocus()
    If Len(txtDesde.Text) > 0 Then
        txtDesde.Text = Format(txtDesde.Text, "00000000")
    End If
End Sub
Private Sub txtFechaReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnBuscar.SetFocus
    End If
End Sub
Private Sub txtFechaReg_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFechaReg.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtFechaReg.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txthasta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        chkFechaReg.SetFocus
    End If
End Sub
Private Sub txtHasta_LostFocus()
    If Len(txthasta.Text) > 0 Then
        txthasta.Text = Format(txthasta.Text, "00000000")
    End If
End Sub
Private Sub limpiarCampos()
    chkIFI.value = 0
    chkNroCheque.value = 0
    chkFechaReg.value = 0
    cboMoneda.ListIndex = -1
    txtIFICod.Text = ""
    lblIFINroCuenta.Caption = ""
    txtDesde.Text = ""
    txthasta.Text = ""
    txtFechaReg.Text = "__/__/____"
    chkTodosSel.value = 0
    Call FormateaFlex(feCheque)
End Sub
Private Sub chkTodosSel_Click()
    Dim i As Long
    Dim lnCheck As Integer
    
    If feCheque.Rows = 2 And feCheque.TextMatrix(1, 0) = "" Then 'Esta Vacío
        Exit Sub
    End If
    lnCheck = chkTodosSel.value

    For i = 1 To feCheque.Rows - 1
        feCheque.TextMatrix(i, 6) = lnCheck
    Next
End Sub
Private Sub btnBuscar_Click()
    Dim oDoc As New DDocRec
    Dim rs As New ADODB.Recordset
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String
    Dim lnNroChequeIni As Long, lnNroChequeFin As Long
    Dim ldFechaReg As Date
    Dim i As Long
    
    If validaCamposBuscar = False Then Exit Sub
    
    If chkIFI.value = 0 Then
        lsIFTpo = Mid(txtIFICod.Text, 1, 2)
        lsPersCod = Mid(txtIFICod.Text, 4, 13)
        lsCtaIFCod = Mid(txtIFICod.Text, 18, 10)
    End If
    If chkNroCheque.value = 0 Then
        lnNroChequeIni = CLng(txtDesde.Text)
        lnNroChequeFin = CLng(txthasta.Text)
    End If
    If chkFechaReg.value = 0 Then
        ldFechaReg = CDate(txtFechaReg.Text)
    Else
        ldFechaReg = gdFecSis
    End If
    
    Call FormateaFlex(feCheque)
    chkTodosSel.value = 0
    Set rs = oDoc.RecuperaChequesMantenimiento(chkIFI.value, lsIFTpo, lsPersCod, lsCtaIFCod, chkNroCheque.value, lnNroChequeIni, lnNroChequeFin, chkFechaReg.value, ldFechaReg)
    
    If rs.RecordCount = 0 Then
        MsgBox "No se encontraron datos para la Busqueda realizada", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Do While Not rs.EOF
        feCheque.AdicionaFila
        i = feCheque.row
        feCheque.TextMatrix(i, 1) = rs!cIFTpo
        feCheque.TextMatrix(i, 2) = rs!cPersCod
        feCheque.TextMatrix(i, 3) = rs!cCtaIfCod
        feCheque.TextMatrix(i, 4) = rs!nMovNro
        feCheque.TextMatrix(i, 5) = rs!nMovFlag
        feCheque.TextMatrix(i, 7) = rs!cPersNombre
        feCheque.TextMatrix(i, 8) = rs!cCtaIFDesc
        feCheque.TextMatrix(i, 9) = Format(rs!dFechaCrea, "dd/mm/yyyy")
        feCheque.TextMatrix(i, 10) = rs!cNroCheque
        feCheque.TextMatrix(i, 11) = rs!cEstado
        feCheque.TextMatrix(i, 12) = IIf(Len(rs!cMovNro) >= 8, Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4), "")
        feCheque.TextMatrix(i, 13) = rs!cMovDesc
        rs.MoveNext
    Loop
    feCheque.TopRow = 1
    feCheque.row = 1
End Sub
Private Sub btnEliminar_Click()
    Dim oDocRec As DDocRec
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String, lsNroCheque As String
    Dim i As Long
    Dim lbExiste As Boolean
    Dim lsComentario As String
    Dim ldFecha As Date
    
    'Verifica se hayan seleccionado
    lbExiste = False
    For i = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(i, 6) = "." Then
            lbExiste = True
            Exit For
        End If
    Next
    If Not lbExiste Then
        MsgBox "Ud. debe seleccionar los registros a Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    'Verifica sea solo registrados
    For i = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(i, 6) = "." Then
            If feCheque.TextMatrix(i, 11) <> "R" Then
                MsgBox "Cheque Nro. " & feCheque.TextMatrix(i, 10) & " ya fue utilizado, no se puede eliminar", vbInformation, "Aviso"
                Me.feCheque.TopRow = i
                Me.feCheque.row = i
                Me.feCheque.col = 7
                Exit Sub
            End If
        End If
    Next
    'Verifica ingreso comentario
    lsComentario = frmChequeComentarioGral.Inicio(1)
    If lsComentario = "" Then
        Exit Sub
    Else
        lsComentario = "ELIMINACIÓN: " & lsComentario
    End If
    
    If MsgBox("Si elimina los cheques NO podrán ser utilizados en ninguna operación" & Chr(10) & "¿Esta seguro de eliminar los cheques seleccionados?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    ldFecha = CDate(gdFecSis & " " & Format(Time, "hh:mm:ss"))
    Set oDocRec = New DDocRec
    For i = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(i, 6) = "." Then
            lsIFTpo = Trim(feCheque.TextMatrix(i, 1))
            lsPersCod = Trim(feCheque.TextMatrix(i, 2))
            lsCtaIFCod = Trim(feCheque.TextMatrix(i, 3))
            lsNroCheque = Trim(feCheque.TextMatrix(i, 10))
            Call oDocRec.ActualizaCheque(lsIFTpo, lsPersCod, lsCtaIFCod, lsNroCheque, 1, lsComentario, ldFecha)
        End If
    Next
    
    MsgBox "Se ha grabado satisfactoriamente los cambios", vbInformation, "Aviso"
    btnLimpiar_Click
    Set oDocRec = Nothing
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Se Elimino Cheque "
                Set objPista = Nothing
                '****
End Sub
 Private Sub btnAnular_Click()
    Dim oDocRec As DDocRec
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String, lsNroCheque As String
    Dim i As Long
    Dim lsComentario As String
    Dim ldFecha As Date
    Dim lnNroRegistros As Long
    
    'Verifica se hayan seleccionado
    lnNroRegistros = 0
    For i = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(i, 6) = "." Then
            lnNroRegistros = lnNroRegistros + 1
        End If
    Next
    If lnNroRegistros = 0 Then
        MsgBox "Ud. debe seleccionar los registros a Anular", vbInformation, "Aviso"
        Exit Sub
    ElseIf lnNroRegistros > 1 Then
        MsgBox "Solo se puede anular un registro por operación", vbInformation, "Aviso"
        Exit Sub
    End If
    'Verifica sea solo usado y el moviento del mismo este extornado, para que el cheque quede inutilizable
    For i = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(i, 6) = "." Then
    'MIOL 20130605, SEGUN RQ13327 ******************************************
            If feCheque.TextMatrix(i, 11) <> "U" And feCheque.TextMatrix(i, 11) <> "R" Then
'            If feCheque.TextMatrix(i, 11) <> "U" Then
                MsgBox "Cheque Nro. " & feCheque.TextMatrix(i, 10) & " no esta en estado Usado, no se puede Anular", vbInformation, "Aviso"
                Me.feCheque.TopRow = i
                Me.feCheque.row = i
                Me.feCheque.col = 7
                Exit Sub
            Else
                If feCheque.TextMatrix(i, 11) = "U" And CLng(feCheque.TextMatrix(i, 5)) = 0 Then
'                If CLng(feCheque.TextMatrix(i, 5)) = 0 Then
                    MsgBox "Cheque Nro. " & feCheque.TextMatrix(i, 10) & " tiene movimiento vigente, no se puede Anular", vbInformation, "Aviso"
                    Me.feCheque.TopRow = i
                    Me.feCheque.row = i
                    Me.feCheque.col = 7
                    Exit Sub
                End If
            End If
    'END MIOL **************************************************************
        End If
    Next
    'Verifica ingreso comentario
    lsComentario = frmChequeComentarioGral.Inicio(2)
    If lsComentario = "" Then
        Exit Sub
    Else
        lsComentario = "ANULACIÓN: " & lsComentario
    End If
    
    If MsgBox("¿Esta seguro de Anular los cheques seleccionados?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    ldFecha = CDate(gdFecSis & " " & Format(Time, "hh:mm:ss"))
    Set oDocRec = New DDocRec
    For i = 1 To feCheque.Rows - 1
        If feCheque.TextMatrix(i, 6) = "." Then
            lsIFTpo = Trim(feCheque.TextMatrix(i, 1))
            lsPersCod = Trim(feCheque.TextMatrix(i, 2))
            lsCtaIFCod = Trim(feCheque.TextMatrix(i, 3))
            lsNroCheque = Trim(feCheque.TextMatrix(i, 10))
            Call oDocRec.ActualizaCheque(lsIFTpo, lsPersCod, lsCtaIFCod, lsNroCheque, 2, , , lsComentario, ldFecha)
        End If
    Next
    
    MsgBox "Se ha grabado satisfactoriamente los cambios", vbInformation, "Aviso"
    btnLimpiar_Click
    Set oDocRec = Nothing
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Anulo Cheque "
                Set objPista = Nothing
                '****
End Sub
Private Sub btnExportar_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lnPosActual As Long, lnPosAnterior As Long, i As Long
    Dim lsArchivo As String
    
On Error GoTo ErrExportar
    
    If feCheque.Rows = 2 And feCheque.TextMatrix(1, 0) = "" Then 'Esta Vacío
        MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lsArchivo = "\spooler\RptTalonarioCheque" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Reporte Talonario Cheques"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    xlsHoja.Columns("A:A").ColumnWidth = 4
    xlsHoja.Columns("B:B").ColumnWidth = 50
    xlsHoja.Columns("C:C").ColumnWidth = 20
    xlsHoja.Columns("D:D").ColumnWidth = 10
    xlsHoja.Columns("E:E").ColumnWidth = 10
    xlsHoja.Columns("F:F").ColumnWidth = 10
    xlsHoja.Columns("G:G").ColumnWidth = 10
    xlsHoja.Columns("H:H").ColumnWidth = 80
    
    xlsHoja.Range("B2") = "Institución Financiera"
    xlsHoja.Range("C2") = "Cuenta"
    xlsHoja.Range("D2") = "F.Reg"
    xlsHoja.Range("E2") = "Número"
    xlsHoja.Range("F2") = "Estado"
    xlsHoja.Range("G2") = "Fecha Uso"
    xlsHoja.Range("H2") = "Glosa"
    
    xlsHoja.Range("E:E").NumberFormat = "_(@_)"
    xlsHoja.Range("F:F").HorizontalAlignment = xlCenter
    
    lnPosActual = 2
    lnPosAnterior = lnPosActual
    For i = 1 To feCheque.Rows - 1
        lnPosActual = lnPosActual + 1
        xlsHoja.Range("B" & lnPosActual) = feCheque.TextMatrix(i, 7)
        xlsHoja.Range("C" & lnPosActual) = feCheque.TextMatrix(i, 8)
        xlsHoja.Range("D" & lnPosActual) = feCheque.TextMatrix(i, 9)
        xlsHoja.Range("E" & lnPosActual) = feCheque.TextMatrix(i, 10)
        xlsHoja.Range("F" & lnPosActual) = feCheque.TextMatrix(i, 11)
        xlsHoja.Range("G" & lnPosActual) = feCheque.TextMatrix(i, 12)
        xlsHoja.Range("H" & lnPosActual) = feCheque.TextMatrix(i, 13)
    Next
    
    xlsHoja.Range("B" & lnPosAnterior & ":H" & lnPosAnterior).Font.Bold = True
    xlsHoja.Range("B" & lnPosAnterior & ":H" & lnPosAnterior).HorizontalAlignment = xlCenter
    xlsHoja.Range("B" & lnPosAnterior & ":H" & lnPosAnterior).Interior.Color = RGB(191, 191, 191)
    xlsHoja.Range("B" & lnPosAnterior, "H" & lnPosActual).Borders.Weight = xlThin
    
    MsgBox "Se ha exportado satisfactoriamente la información", vbInformation, "Aviso"
    
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
    
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " Se Genero Reporte"
                Set objPista = Nothing
                '****
    Exit Sub
ErrExportar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function validaCamposBuscar() As Boolean
    validaCamposBuscar = True
    If chkIFI.value = 0 Then
        If Len(txtIFICod.Text) <> 24 Then
            MsgBox "Ud. debe seleccionar la Cuenta de la Institución Financiera", vbInformation, "Aviso"
            txtIFICod.SetFocus
            validaCamposBuscar = False
            Exit Function
        End If
    End If
    If chkNroCheque.value = 0 Then
        If Not IsNumeric(txtDesde.Text) Then
            MsgBox "Ud. debe especificar el Nro de Cheque de Inicio", vbInformation, "Aviso"
            txtDesde.SetFocus
            validaCamposBuscar = False
            Exit Function
        Else
            If CLng(txtDesde.Text) <= 0 Then
                MsgBox "El Nro de Cheque de Inicio debe ser mayor a cero", vbInformation, "Aviso"
                txtDesde.SetFocus
                validaCamposBuscar = False
                Exit Function
            End If
        End If
        If Not IsNumeric(txthasta.Text) Then
            MsgBox "Ud. debe especificar el Nro de Cheque Final", vbInformation, "Aviso"
            txthasta.SetFocus
            validaCamposBuscar = False
            Exit Function
        Else
            If CLng(txthasta.Text) <= 0 Then
                MsgBox "El Nro de Cheque Fin debe ser mayor a cero", vbInformation, "Aviso"
                txthasta.SetFocus
                validaCamposBuscar = False
                Exit Function
            End If
        End If
        If CLng(txtDesde.Text) > CLng(txthasta.Text) Then
            MsgBox "El Nro de Cheque de Inicio no puede ser mayor que el Nro Cheque Final", vbInformation, "Aviso"
            txtDesde.SetFocus
            validaCamposBuscar = False
            Exit Function
        End If
    End If
    If chkFechaReg.value = 0 Then
        If Not IsDate(txtFechaReg.Text) Then
            MsgBox "La Fecha de Registro no es correcta", vbInformation, "Aviso"
            txtFechaReg.SetFocus
            validaCamposBuscar = False
            Exit Function
        End If
    End If
End Function
Private Sub FormateaFlex(ByVal pflex As FlexEdit)
    pflex.Clear
    pflex.FormaCabecera
    pflex.Rows = 2
End Sub
