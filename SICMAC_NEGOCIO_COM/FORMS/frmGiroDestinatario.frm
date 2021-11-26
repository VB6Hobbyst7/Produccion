VERSION 5.00
Begin VB.Form frmGiroDestinatario 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   5235
   ClientTop       =   3900
   ClientWidth     =   5475
   Icon            =   "frmGiroDestinatario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   2640
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Frame fraDatos 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   5235
      Begin VB.Frame fraCliente 
         Height          =   1515
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5055
         Begin VB.Label lblCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2940
            TabIndex        =   19
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Left            =   2220
            TabIndex        =   18
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1020
            TabIndex        =   17
            Top             =   1020
            Width           =   3735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dirección :"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1095
            Width           =   765
         End
         Begin VB.Label lblNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1020
            TabIndex        =   15
            Top             =   660
            Width           =   3735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nombre :"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblDocID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1020
            TabIndex        =   13
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Doc ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   375
            Width           =   555
         End
      End
      Begin VB.Frame fraNoCliente 
         Height          =   1515
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   4995
         Begin VB.ComboBox cboTipoDoi 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   190
            Width           =   1685
         End
         Begin VB.TextBox txtnumdoc 
            Height          =   375
            Left            =   3185
            TabIndex        =   2
            Top             =   180
            Width           =   1690
         End
         Begin VB.TextBox txtReferencia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   4
            Top             =   1020
            Width           =   3915
         End
         Begin VB.TextBox txtNombre 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   3
            Top             =   600
            Width           =   3915
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo DOI:"
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "DOI :"
            Height          =   195
            Left            =   2700
            TabIndex        =   20
            Top             =   255
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Referencia :"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre :"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   660
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkCliCMACT 
         Caption         =   "Cliente CMAC Maynas"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmGiroDestinatario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearScreen()
txtNombre = ""
txtReferencia = ""
txtnumdoc = ""
chkCliCMACT.value = 0
lblNombre = ""
lblDireccion = ""
lblCodigo = ""
lblDocID = ""
cmdBuscar.Visible = False
fraCliente.Visible = False
txtNombre = ""
txtReferencia = ""
End Sub

Private Sub chkCliCMACT_Click()
If chkCliCMACT.value = 1 Then
    fraCliente.Visible = True
    cmdBuscar.Visible = True
    fraNoCliente.Visible = False
    cmdBuscar.SetFocus
Else
    fraCliente.Visible = False
    cmdBuscar.Visible = False
    fraNoCliente.Visible = True
    txtnumdoc.SetFocus
End If
End Sub

Private Sub CmdAceptar_Click()
If chkCliCMACT.value = 1 Then
    If lblCodigo = "" Then
        MsgBox "No ha registrado ningun destinatario.", vbInformation, "Aviso"
        cmdBuscar.SetFocus
        Exit Sub
    End If
Else
   '20180611 NRLO INC1804170006
    Dim sDoi As String
    Dim cTipo As String
    sDoi = txtnumdoc.Text
    cTipo = CInt(Trim(Right(cboTipoDoi.Text, 2)))
    'END 20180611 NRLO INC1804170006
    
    If Trim(txtNombre) = "" Then
        MsgBox "No ha digitado el nombre del destinatario.", vbInformation, "Aviso"
        txtNombre.SetFocus
        Exit Sub
    End If
    
    Dim j As Integer

    If Len(txtnumdoc.Text) = 0 Then
            MsgBox "Falta Ingresar el Numero de Documento", vbInformation, "Aviso"
            txtnumdoc.SetFocus
            Exit Sub
    End If
    
    'If Len(txtnumdoc.Text) <> gnNroDigitosDNI Then   '20180611 NRLO INC1804170006
    '      MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
    '      txtnumdoc.SetFocus
    '      Exit Sub
    'End If
    
    '20180611 NRLO INC1804170006
    If cTipo = 1 And Len(sDoi) <> gnNroDigitosDNI Then
        MsgBox "DOI No es de " & gnNroDigitosDNI & " dígitos", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        Exit Sub
    ElseIf cTipo = 2 And Len(sDoi) <> 11 Then
        MsgBox "DOI No es de 11 digitos", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        Exit Sub
    ElseIf cTipo = 4 And (Len(sDoi) > 12 Or Len(sDoi) < 5) Then
        MsgBox "DOI debe tener entre 5 y 12 dígitos", vbInformation, "Aviso"
        txtnumdoc.SetFocus
        Exit Sub
    End If
    'END 20180611 NRLO INC1804170006
    
    'NRLO COMENTÓ ESTO 20180611 INC1804170006
    'Else
    '    If Len(sDoi) < 8 Then
    '      MsgBox "DOI No es de al menos 8 dígitos", vbInformation, "Aviso"
    '      txtnumdoc.SetFocus
    '      Exit Sub
    '    End If
    'END 20180611 NRLO INC1804170006
    
    If cTipo = 1 Then
        For j = 1 To Len(Trim(txtnumdoc.Text))
            If (Mid(txtnumdoc.Text, j, 1) < "0" Or Mid(txtnumdoc.Text, j, 1) > "9") Then
               MsgBox "Uno de los Digitos del DNI no es un Numero", vbInformation, "Aviso"
               txtnumdoc.SetFocus
               Exit Sub
            End If
        Next j
    End If
    
    If BuscaNumDoc(cTipo, Trim(txtnumdoc.Text)) Then
        MsgBox "El Numero de DOI pertenece a un Cliente CMAC Maynas", vbInformation, "Aviso"
        Exit Sub
    End If
    
End If

Dim nFila As Long
With frmGiroApertura
    .grdDest.AdicionaFila , , True
    nFila = .grdDest.rows - 1
    If chkCliCMACT.value = 1 Then
        .grdDest.TextMatrix(nFila, 1) = lblDocID
        .grdDest.TextMatrix(nFila, 2) = lblNombre
        .grdDest.TextMatrix(nFila, 3) = lblDireccion
        .grdDest.TextMatrix(nFila, 4) = lblCodigo
    Else
        .grdDest.TextMatrix(nFila, 1) = txtnumdoc 'comentado x MADM 20101011 ""
        .grdDest.TextMatrix(nFila, 2) = txtNombre
        .grdDest.TextMatrix(nFila, 3) = Trim(txtReferencia)
        .grdDest.TextMatrix(nFila, 4) = ""
    End If
End With
Unload Me
End Sub

Private Function BuscaNumDoc(ByVal pnTipoDoc As Integer, ByVal psNumDoc As String) As Boolean
    Dim ObjP As COMDPersona.DCOMPersonas
    Dim rs As ADODB.Recordset
       
    BuscaNumDoc = False ''no encontrado
         
         Set ObjP = New COMDPersona.DCOMPersonas
         Set rs = New ADODB.Recordset
         Set rs = ObjP.BusquedaNumDocPersona(pnTipoDoc, psNumDoc)

        If Not (rs.EOF And rs.BOF) Then
            BuscaNumDoc = True ''encontrado
            Set rs = Nothing
            Set ObjP = Nothing
            Exit Function
        End If
        
End Function

Private Sub CmdBuscar_Click()
    Dim clsPers As COMDPersona.UCOMPersona
    Set clsPers = frmBuscaPersona.Inicio(False)
    If Not clsPers Is Nothing Then
        If clsPers.sPerscod <> "" Then
            lblDocID = clsPers.sPersIdnroDNI
            '20181214 NRLO INC1804170006
            If Len(lblDocID) = 0 Then
                lblDocID = clsPers.sPersIdnroOtro
            End If
            'END 20181214 NRLO INC1804170006
            lblNombre = PstaNombre(Trim(clsPers.sPersNombre), False)
            lblDireccion = Trim(clsPers.sPersDireccDomicilio)
            lblCodigo = clsPers.sPerscod
            cmdAceptar.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Me.Caption = "Giro - Registro Destinatario"
    ClearScreen
    Call CargarDOIs   '20180611 NRLO INC1804170006
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtReferencia.SetFocus
    End If
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub
'''''''''''''' MADM 20101110
Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
    Dim cTipo As String
    cTipo = CInt(Trim(Right(cboTipoDoi.Text, 2)))
    
    If KeyAscii = 13 Then
        txtNombre.SetFocus
    End If
    
    If cTipo = 1 Then
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
    If cTipo = 4 Then
        KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    End If
End Sub

Private Sub txtnumdoc_LostFocus()
'Dim j As Integer
'
'If Len(txtNumDoc.Text) = 0 Then
'        MsgBox "Falta Ingresar el Numero de Documento", vbInformation, "Aviso"
'        txtNumDoc.SetFocus
'        Exit Sub
'End If
'
'If Len(txtNumDoc.Text) <> gnNroDigitosDNI Then
'      MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
'      txtNumDoc.SetFocus
'      Exit Sub
'End If
'
'For j = 1 To Len(Trim(txtNumDoc.Text))
'    If (Mid(txtNumDoc.Text, j, 1) < "0" Or Mid(txtNumDoc.Text, j, 1) > "9") Then
'       MsgBox "Uno de los Digitos del DNI no es un Numero", vbInformation, "Aviso"
'       txtNumDoc.SetFocus
'       Exit Sub
'    End If
'Next j

End Sub
''''''''''''''end MADM 20101110

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub

Private Sub CargarDOIs() '20181214 NRLO INC1804170006
    Dim lsDOIs As ADODB.Recordset
    Dim lsDOIFinal As ADODB.Recordset
    Dim oConstante As COMDConstantes.DCOMConstantes
    Set oConstante = New COMDConstantes.DCOMConstantes
    Set lsDOIs = oConstante.RecuperaConstantes(1003)
    
    Set lsDOIFinal = New ADODB.Recordset
    lsDOIFinal.Fields.Append "nConsValor", adInteger
    lsDOIFinal.Fields.Append "cConsDescripcion", adBSTR, 10, adFldUpdatable
    lsDOIFinal.Open
    
    Do While Not lsDOIs.EOF
        If (lsDOIs!nConsValor = 1 Or lsDOIs!nConsValor = 4) Then
            lsDOIFinal.AddNew
            lsDOIFinal.Fields!cConsDescripcion = lsDOIs!cConsDescripcion
            lsDOIFinal.Fields!nConsValor = lsDOIs!nConsValor
            lsDOIFinal.Update
            lsDOIFinal.MoveFirst
        End If
        lsDOIs.MoveNext
    Loop
    lsDOIs.Close

    cboTipoDoi.Clear
    Call Llenar_Combo_con_Recordset(lsDOIFinal, Me.cboTipoDoi)
    cboTipoDoi.ListIndex = 0
    
    txtnumdoc.MaxLength = 8
    txtNombre.MaxLength = 70
    txtReferencia.MaxLength = 70
End Sub
'END 20181214 NRLO INC1804170006

Private Sub cboTipoDoi_Click()
    Dim cTipo As String
    cTipo = CInt(Trim(Right(cboTipoDoi.Text, 2)))
    
    If cTipo = 1 Then
        txtnumdoc.Text = "" 'Left(txtnumdoc.Text, 8)
        txtnumdoc.MaxLength = 8
    ElseIf cTipo = 2 Then
        txtnumdoc.Text = Left(txtnumdoc.Text, 11)
        txtnumdoc.MaxLength = 11
    ElseIf cTipo = 4 Then
        txtnumdoc.Text = "" 'Left(txtnumdoc.Text, 12)
        txtnumdoc.MaxLength = 12
        MsgBox "Seleccionó CARNET DE EXTRANJERÍA, se permitirá el ingreso de caracteres alfanumericos, " & _
            "podrá digitar entre 5 a 12 caracteres." & Chr(13) & _
            "Al finalizar verifique el DOI insertado con detenimiento para evitar inconvenientes futuros.", vbInformation, "Aviso"
    Else
        txtnumdoc.MaxLength = 20
    End If
End Sub
