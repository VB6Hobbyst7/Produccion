VERSION 5.00
Begin VB.Form frmChequeTalonario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Talonario de Cheque"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "frmChequeTalonario.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRegistrar 
      Caption         =   "Registrar"
      Height          =   345
      Left            =   5880
      TabIndex        =   4
      Top             =   1980
      Width           =   1050
   End
   Begin VB.CommandButton btnCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6960
      TabIndex        =   5
      Top             =   1980
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Talonario de Cheque"
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
      Height          =   1900
      Left            =   40
      TabIndex        =   6
      Top             =   40
      Width           =   7980
      Begin VB.TextBox txtDesde 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   2295
         MaxLength       =   8
         TabIndex        =   2
         Tag             =   "0"
         Top             =   1500
         Width           =   1680
      End
      Begin VB.TextBox txtHasta 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   6135
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1500
         Width           =   1680
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmChequeTalonario.frx":030A
         Left            =   960
         List            =   "frmChequeTalonario.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1740
      End
      Begin Sicmact.TxtBuscar txtIFICod 
         Height          =   350
         Left            =   960
         TabIndex        =   1
         Top             =   645
         Width           =   2580
         _extentx        =   4551
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmChequeTalonario.frx":041B
         appearance      =   1
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   13
         Top             =   1545
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   960
         Top             =   1485
         Width           =   3045
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   12
         Top             =   1545
         Width           =   735
      End
      Begin VB.Label lblIFINroCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   960
         TabIndex        =   11
         Top             =   1080
         Width           =   6930
      End
      Begin VB.Label lblIFINombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3600
         TabIndex        =   10
         Top             =   650
         Width           =   4290
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   705
         Width           =   615
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   4800
         Top             =   1485
         Width           =   3045
      End
   End
End
Attribute VB_Name = "frmChequeTalonario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'** Nombre : frmChequeTalonario
'** Descripción : Clase de Cheques  creado según RFC117-2012
'** Creación : EJVG 20121124 09:00:00 AM
'***********************************************************************
Option Explicit
Dim fsOpeCod As String

Private Sub Form_Load()
    CentraForm Me
    limpiarCampos
End Sub
Private Sub btnRegistrar_Click()
    Dim oDocRec As New NDocRec
    Dim rsRepetidosEnTalonario As ADODB.Recordset, rsRepetidoEnMOV As ADODB.Recordset
    Dim lsIFTpo As String, lsPersCod As String, lsCtaIFCod As String
    Dim lsMsgError As String
    Dim ldFecha As Date
    Dim lnNroChequeIni As Long, lnNroChequeFin As Long
    Dim i As Long, iMat As Long
    Dim MatRepetidosMov() As Variant
    Dim lsCadenaChequesMov As String
    
On Error GoTo ErrRegistrar
    If validaCampos = False Then Exit Sub

    lsIFTpo = Mid(txtIFICod.Text, 1, 2)
    lsPersCod = Mid(txtIFICod.Text, 4, 13)
    lsCtaIFCod = Mid(txtIFICod.Text, 18, Len(txtIFICod.Text))
    ldFecha = CDate(gdFecSis & " " & Format(Time, "hh:mm:ss"))
    lnNroChequeIni = CLng(txtDesde.Text)
    lnNroChequeFin = CLng(txtHasta.Text)
    
    'Repetidos en Talonario
    Set rsRepetidosEnTalonario = New ADODB.Recordset
    Set rsRepetidosEnTalonario = oDocRec.RecuperaChequesTalonario(lsIFTpo, lsPersCod, lsCtaIFCod, lnNroChequeIni, lnNroChequeFin)
    If Not RSVacio(rsRepetidosEnTalonario) Then
        MsgBox "Se han encontrado Nro de Cheques existentes en el rango especificado, verifique..", vbExclamation, "Aviso"
        Set oDocRec = Nothing
        Set rsRepetidosEnTalonario = Nothing
        Exit Sub
    End If
    
    'Repetidos en MovDoc
    ReDim MatRepetidosMov(1 To 2, 0 To 0)
    For i = lnNroChequeIni To lnNroChequeFin
        Set rsRepetidoEnMOV = New ADODB.Recordset
        Set rsRepetidoEnMOV = oDocRec.RecuperaChequeExistenteEnMovDoc(lsIFTpo, lsPersCod, lsCtaIFCod, Format(i, "00000000"))
        If Not RSVacio(rsRepetidoEnMOV) Then
            iMat = UBound(MatRepetidosMov, 2) + 1
            ReDim Preserve MatRepetidosMov(1 To 2, 0 To iMat)
            MatRepetidosMov(1, iMat) = Format(i, "00000000")
            MatRepetidosMov(2, iMat) = rsRepetidoEnMOV!cMovNro
        End If
    Next
    lsCadenaChequesMov = RecuperaCadenaChequesEnMovDoc(MatRepetidosMov)
    If Len(lsCadenaChequesMov) > 0 Then
        EnviaPrevio lsCadenaChequesMov, "CHEQUES UTILIZADOS CON ANTERIORIDAD", gnLinPage, True
    End If
    
    If MsgBox("¿Esta seguro de grabar el Talonario de Cheques?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
        
    btnRegistrar.Enabled = False
    lsMsgError = oDocRec.grabarTalonarioCheque(lsIFTpo, lsPersCod, lsCtaIFCod, lnNroChequeIni, lnNroChequeFin, gsCodUser, ldFecha)
    btnRegistrar.Enabled = True
    
    If lsMsgError <> "" Then
        MsgBox lsMsgError, vbCritical, "Aviso"
        Exit Sub
    End If
    
    MsgBox "Se ha registrado satisfactoriamente el Talonario de Cheques", vbInformation, "Aviso"
    Call limpiarCampos
    btnRegistrar.Enabled = True
    
    Set rsRepetidosEnTalonario = Nothing
    Set oDocRec = Nothing
    Exit Sub
ErrRegistrar:
    MsgBox "Ocurrió un error al registrar el Talonario de Cheques", vbCritical, "Aviso"
End Sub
Private Sub btnCancelar_Click()
    Unload Me
End Sub
Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtHasta.SetFocus
    End If
End Sub
Private Sub txtDesde_LostFocus()
    If Len(txtDesde.Text) > 0 Then
        txtDesde.Text = Format(txtDesde.Text, "00000000")
    End If
End Sub
Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        btnRegistrar.SetFocus
    End If
End Sub
Private Sub txtHasta_LostFocus()
    If Len(txtHasta.Text) > 0 Then
        txtHasta.Text = Format(txtHasta.Text, "00000000")
    End If
End Sub
Private Sub txtIFICod_EmiteDatos()
    Dim oCtaIf As New NCajaCtaIF
    lblIFINombre.Caption = ""
    lblIFINroCuenta.Caption = ""
    If txtIFICod.Text <> "" Then
        lblIFINombre.Caption = oCtaIf.NombreIF(Mid(txtIFICod.Text, 4, 13))
        lblIFINroCuenta.Caption = oCtaIf.EmiteTipoCuentaIF(Mid(txtIFICod.Text, 18, Len(txtIFICod.Text))) & " " & txtIFICod.psDescripcion
    End If
    Set oCtaIf = Nothing
End Sub
Private Sub txtIFICod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDesde.SetFocus
    End If
End Sub
Private Sub txtIFICod_GotFocus()
    If cboMoneda.ListIndex = -1 Then
        MsgBox "Ud. debe seleccionar la moneda de la operación", vbInformation, "Aviso"
        cboMoneda.SetFocus
    End If
End Sub
Private Sub cboMoneda_Click()
    Dim oOpe As New DOperacion
    If CInt(Trim(Right(cboMoneda.Text, 3))) = 1 Then
        fsOpeCod = OpeChqRegistroTalonarioMN
    Else
        fsOpeCod = OpeChqRegistroTalonarioME
    End If
    txtIFICod.rs = oOpe.GetOpeObj(fsOpeCod, "2")
    Set oOpe = Nothing
End Sub
Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtIFICod.SetFocus
    End If
End Sub
Private Function validaCampos() As Boolean
    validaCampos = True
    If Len(txtIFICod.Text) <> 24 Then
        MsgBox "Ud. debe seleccionar la Cuenta del Banco", vbInformation
        txtIFICod.SetFocus
        validaCampos = False
        Exit Function
    End If
    If Not IsNumeric(txtDesde.Text) Then
        MsgBox "Ud. debe especificar el Nro de Cheque de Inicio", vbInformation
        txtDesde.SetFocus
        validaCampos = False
        Exit Function
    Else
        If CLng(txtDesde.Text) <= 0 Then
            MsgBox "El Nro de Cheque de Inicio debe ser mayor a cero", vbInformation
            txtDesde.SetFocus
            validaCampos = False
            Exit Function
        End If
    End If
    If Not IsNumeric(txtHasta.Text) Then
        MsgBox "Ud. debe especificar el Nro de Cheque Final", vbInformation
        txtHasta.SetFocus
        validaCampos = False
        Exit Function
    Else
        If CLng(txtHasta.Text) <= 0 Then
            MsgBox "El Nro de Cheque Fin debe ser mayor a cero", vbInformation
            txtHasta.SetFocus
            validaCampos = False
            Exit Function
        End If
    End If
    If CLng(txtDesde.Text) > CLng(txtHasta.Text) Then
        MsgBox "El Nro Cheque Inicial no puede ser mayor que el Final", vbInformation
        txtDesde.SetFocus
        validaCampos = False
        Exit Function
    End If
End Function
Private Sub limpiarCampos()
    txtIFICod.Text = ""
    lblIFINombre.Caption = ""
    lblIFINroCuenta.Caption = ""
    txtDesde.Text = ""
    txtHasta.Text = ""
End Sub
Private Function RecuperaCadenaChequesEnMovDoc(ByRef pMatCheques As Variant) As String
    Dim lsCadena As String
    Dim lsItem As String * 4
    Dim lsNroCheque As String * 8
    Dim lsMovNro As String * 25
    Dim i As Long
    Dim lnLinea As Long, lnNroPagina As Long

    lnLinea = 57
    lnNroPagina = 1
    For i = 1 To UBound(pMatCheques, 2)
        If lnLinea > 56 Then
            If i > 1 Then
                lsCadena = lsCadena & Chr(12)
                lnNroPagina = lnNroPagina + 1
            End If
            lsCadena = lsCadena & Cabecera("CHEQUES ANTERIORMENTE FUERON REGISTRADOS EN OPERACIONES", lnNroPagina - 1, "", gnColPage, "", Format(gdFecSis, "dd/mm/yyyy"), "")
            lsCadena = lsCadena & String(40, "-") & Chr(10)
            lsCadena = lsCadena & "ITEM" & Space(1) & "NRO CHEQUE" & Space(1) & "      NRO MOVIMIENTO     " & Chr(10)
            lsCadena = lsCadena & String(40, "-") & Chr(10)
            lnLinea = 5
        End If

        lsItem = Centra(CStr(i), 4)
        lsNroCheque = Centra(CStr(pMatCheques(1, i)), 8)
        lsMovNro = pMatCheques(2, i)
    
        lsCadena = lsCadena & lsItem & Space(1) & lsNroCheque & Space(1) & lsMovNro & Chr(10)
        lnLinea = lnLinea + 1
    Next
    RecuperaCadenaChequesEnMovDoc = lsCadena
End Function
