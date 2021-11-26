VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenMovEfectivo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   1785
   ClientTop       =   2355
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenMovEfectivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDepositar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4755
      TabIndex        =   24
      Top             =   4095
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6135
      TabIndex        =   6
      Top             =   4095
      Width           =   1380
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   630
      Left            =   135
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Width           =   7395
   End
   Begin VB.TextBox txtmonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   4102
      Width           =   1545
   End
   Begin VB.Frame FraDestino 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1020
      Left            =   120
      TabIndex        =   11
      Top             =   1050
      Width           =   7335
      Begin Sicmact.TxtBuscar txtCtaDest 
         Height          =   345
         Left            =   765
         TabIndex        =   2
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
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
      Begin Sicmact.TxtBuscar txtObjDest 
         Height          =   345
         Left            =   750
         TabIndex        =   3
         Top             =   600
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
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
      Begin VB.Label lblDescCtaDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   4995
      End
      Begin VB.Label lblDescObjDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2655
         TabIndex        =   14
         Top             =   600
         Width           =   4515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   165
         Left            =   150
         TabIndex        =   13
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Objeto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   165
         Left            =   135
         TabIndex        =   12
         Top             =   675
         Width           =   540
      End
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   345
      Left            =   6405
      TabIndex        =   0
      Top             =   15
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483635
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame FraOrigen 
      Caption         =   "Origen:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   105
      TabIndex        =   7
      Top             =   360
      Width           =   7335
      Begin Sicmact.TxtBuscar txtCtaOrig 
         Height          =   345
         Left            =   780
         TabIndex        =   1
         Top             =   225
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   135
         TabIndex        =   9
         Top             =   300
         Width           =   570
      End
      Begin VB.Label lblDescCtaOrig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2145
         TabIndex        =   8
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.Frame FraDeposito 
      Caption         =   "Efectivo"
      Height          =   990
      Left            =   150
      TabIndex        =   17
      Top             =   2850
      Width           =   7350
      Begin VB.Frame FraDocumento 
         Caption         =   "Documento"
         Height          =   690
         Left            =   2820
         TabIndex        =   19
         Top             =   195
         Width           =   4425
         Begin VB.ComboBox cboDocumento 
            Height          =   330
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   210
            Width           =   1890
         End
         Begin VB.TextBox txtNroDoc 
            Height          =   315
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   1830
         End
         Begin VB.Label Label7 
            Caption         =   "N° :"
            Height          =   195
            Left            =   2100
            TabIndex        =   22
            Top             =   285
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdefectivo 
         Caption         =   "&Efectivo"
         Height          =   375
         Left            =   615
         TabIndex        =   18
         Top             =   315
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdRetirar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4755
      TabIndex        =   25
      Top             =   4095
      Width           =   1380
   End
   Begin VB.Frame FraRetiro 
      Caption         =   "Emisión de Documentos"
      Height          =   990
      Left            =   150
      TabIndex        =   23
      Top             =   2850
      Width           =   7350
      Begin VB.OptionButton OptDoc 
         Caption         =   "Che&que"
         Height          =   375
         Index           =   1
         Left            =   5865
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   375
         Width           =   1305
      End
      Begin VB.OptionButton OptDoc 
         Caption         =   "&Carta"
         Height          =   375
         Index           =   0
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   375
         Width           =   1305
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   60
      X2              =   7515
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   7560
      Y1              =   3945
      Y2              =   3945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   330
      TabIndex        =   16
      Top             =   4185
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   5715
      TabIndex        =   10
      Top             =   60
      Width           =   570
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   390
      Left            =   225
      Top             =   4080
      Width           =   2910
   End
End
Attribute VB_Name = "frmCajaGenMovEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe As DOperacion
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset
Dim oCtaIf As NCajaCtaIF
Dim lnTipoObj As TpoObjetos
Dim lsNroVoucher As String
Dim lsNombreBanco As String
Dim lsCuenta As String
Dim objPista As COMManejador.Pista

Dim lsNroDoc As String
Dim lsDocumento As String
Dim lnTipoDoc As TpoDoc
'***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D
Dim ldFecCie As Date
'***Fin Modificado por ELRO**********************************

Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNroDoc.SetFocus
End If
End Sub

Private Sub cmdDepositar_Click()
Dim oCajero As nCajaGeneral
Dim oCont As NContFunciones
Set oCont = New NContFunciones
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsCtafiltro As String
Dim lsMovNro As String

If Valida = False Then Exit Sub

'***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D y Acta 311-2011/TI-D
If CDate(txtFecha) <= ldFecCie Then
    MsgBox "Mes ya Cerrado. Imposible realizar la operación...!", vbInformation, "Aviso"
    Exit Sub
Else
    If CDate(txtFecha) < gdFecSis Then
        MsgBox "Día ya Cerrado. Imposible realizar la operación...!", vbInformation, "Aviso"
        Exit Sub
    End If
End If
'***Fin Modificado por ELRO*******************************************************

Set oCajero = New nCajaGeneral
If MsgBox("Desea Grabar el movimiento respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    lsCtafiltro = oCont.GetFiltroObjetos(lnTipoObj, txtCtaDest, txtObjDest, False)
    '***Modificado por ELRO al 20110924, según Acta 263-2011/TI-D
    If lsCtafiltro = "" Then
        MsgBox "Esta cuenta no esta registrado en CtaIFFiltro, comunicarse con TI", vbInformation, "Aviso"
        Exit Sub
    End If
     If verificarUltimoNivelCta(txtCtaDest + lsCtafiltro) = False Then
        MsgBox "Esta cuenta no es de Ultimo Nivel, comunicarse con TI", vbInformation, "Aviso"
        Exit Sub
     End If
    '*** Fin Modificado por ELRO*********************************
    
    oCajero.GrabaMovEfectivo lsMovNro, gsOpeCod, txtMovDesc, _
                rsBill, rsMon, txtCtaDest + lsCtafiltro, txtCtaOrig, txtmonto, lnTipoObj, _
                txtObjDest, Val(Right(cboDocumento, 2)), txtNroDoc, gdFecSis

    ImprimeAsientoContable lsMovNro
    Set frmCajaGenEfectivo = Nothing
    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Deposito de Bancos en Efectivo"
    If MsgBox("Desea realizar otra operación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        '***Modificado por ELRO al 20110923, según Acta 263-2011/TI-D
        'txtCtaDest = ""
        'lblDescCtaDest = ""
        If gsOpeCod = gOpeCGOpeCMACDepEfeMN Or gsOpeCod = gOpeCGOpeCMACDepEfeME Then
                
        Else
            txtCtaDest = ""
            lblDescCtaDest = ""
        End If
        '*** Fin Modificado por ELRO*********************************
        lblDescObjDest = ""
        txtObjDest = ""
        txtMovDesc = ""
        txtmonto = "0.00"
        cboDocumento.ListIndex = -1
        txtNroDoc = ""
        Set rsBill = Nothing
        Set rsMon = Nothing
        
    Else
        Unload Me
        
    End If
End If
End Sub
Function Valida() As Boolean
Valida = True
If ValFecha(txtFecha) = False Then
    Valida = False
    Exit Function
End If
If Len(Trim(txtCtaOrig)) = 0 Then
    MsgBox "Cuenta de " & FraOrigen.Caption & " no seleccionada", vbInformation, "Aviso"
    Valida = False
    If txtCtaOrig.Enabled Then txtCtaOrig.SetFocus
    Exit Function
End If
If Trim(Len(txtCtaDest)) = 0 Then
    MsgBox "Cuenta de " & FraDestino.Caption & " no seleccionada", vbInformation, "Aviso"
    Valida = False
    If txtCtaDest.Enabled Then txtCtaDest.SetFocus
    Exit Function
End If
If Trim(Len(txtObjDest)) = 0 And txtObjDest.Enabled Then
    MsgBox "Objeto " & FraDestino.Caption & " no seleccionada", vbInformation, "Aviso"
    Valida = False
    If txtObjDest.Enabled Then txtObjDest.SetFocus
    Exit Function
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de movimiento no ingresado", vbInformation, "Aviso"
    Valida = False
    txtMovDesc.SetFocus
    Exit Function
End If
If FraDeposito.Visible Then
    If rsBill Is Nothing And rsMon Is Nothing Then
        MsgBox "Billetaje no ha sido ingresado", vbInformation, "Aviso"
        Valida = False
        cmdDepositar.SetFocus
        Exit Function
    End If
    If Len(Trim(cboDocumento)) <> 0 Then
        If Len(Trim(txtNroDoc)) = 0 Then
            MsgBox "Nro de Documento no Ingresado", vbInformation, "Aviso"
            Valida = False
            txtNroDoc.SetFocus
            Exit Function
        End If
    Else
        If MsgBox("Documento no ha sido ingresado. Desea continuar??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            Valida = False
            cboDocumento.SetFocus
            Exit Function
        End If
    End If
End If
If Val(txtmonto) = 0 Then
    MsgBox "Importe de movimiento no ingresado", vbInformation, "Aviso"
    Valida = False
    txtmonto.SetFocus
    Exit Function
End If



End Function
Private Sub cmdCancelar_Click()
Unload Me
End Sub
Private Sub cmdefectivo_Click()
    Set rsBill = New ADODB.Recordset
    Set rsMon = New ADODB.Recordset
    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, 0, Mid(gsOpeCod, 3, 1), False
    If frmCajaGenEfectivo.lbOk Then
        Set rsBill = frmCajaGenEfectivo.rsBilletes
        Set rsMon = frmCajaGenEfectivo.rsMonedas
        txtmonto = frmCajaGenEfectivo.lblTotal
        If FraDocumento.Visible Then
           cboDocumento.SetFocus
        End If
    Else
        Set rsBill = Nothing
        Set rsMon = Nothing
    End If
End Sub
Private Sub cmdRetirar_Click()
Dim oCajero As nCajaGeneral
Dim oCont As NContFunciones
Set oCont = New NContFunciones
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsCtafiltro As String
Dim lsMovNro As String

If Valida = False Then Exit Sub

'***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D y Acta 311-2011/TI-D
If CDate(txtFecha) <= ldFecCie Then
    MsgBox "Mes ya Cerrado. Imposible realizar operación...!", vbInformation, "Aviso"
    Exit Sub
Else
    If CDate(txtFecha) < gdFecSis Then
        MsgBox "Día ya Cerrado. Imposible realizar la operación...!", vbInformation, "Aviso"
        Exit Sub
    End If
End If
'***Fin Modificado por ELRO*******************************************************

If lsNroDoc = "" Or lsDocumento = "" Then
    MsgBox "No ha seleccionado el documento Utilizado en Operación", vbInformation, "aviso"
    OptDoc(0).SetFocus
    Exit Sub
End If
'***Modificado por ELRO al 20110924, según Acta 263-2011/TI-D
lsCtafiltro = oCont.GetFiltroObjetos(lnTipoObj, txtCtaDest, txtObjDest, False)
If lsCtafiltro = "" Then
    MsgBox "Esta cuenta no esta registrado en CtaIFFiltro, comunicarse con TI", vbInformation, "Aviso"
    Exit Sub
End If
If verificarUltimoNivelCta(txtCtaDest + lsCtafiltro) = False Then
    MsgBox "Esta cuenta no es de Ultimo Nivel, comunicarse con TI", vbInformation, "Aviso"
    Exit Sub
End If
'*** Fin Modificado por ELRO

Set oCajero = New nCajaGeneral
If MsgBox("Desea Grabar el movimiento respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(CDate(txtFecha), gsCodAge, gsCodUser)
    
    oCajero.GrabaMovGeneral lsMovNro, gsOpeCod, txtMovDesc, txtCtaOrig, _
                txtCtaDest, CCur(txtmonto), -1, "", lnTipoObj, txtObjDest, lnTipoDoc, lsNroDoc, gdFecSis, lsNroVoucher
              
             
    ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTipoDoc, lsDocumento
    Set frmCajaGenEfectivo = Nothing
    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Retiro de Bancos en Efectivo"
    If MsgBox("Desea realizar otra operación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
        'txtCtaDest = ""
        'lblDescCtaDest = ""
        If gsOpeCod = gOpeCGOpeCMACRetEfeMN Or gsOpeCod = gOpeCGOpeCMACRetEfeME Then
                
        Else
            txtCtaDest = ""
            lblDescCtaDest = ""
        End If
        '***Fin Modificado por ELRO**********************************
        lblDescObjDest = ""
        txtObjDest = ""
        txtMovDesc = ""
        txtmonto = "0.00"
        cboDocumento.ListIndex = -1
        txtNroDoc = ""
        OptDoc(0).value = False
        OptDoc(1).value = False
        lsNroDoc = ""
        lsDocumento = ""
        Set rsBill = Nothing
        Set rsMon = Nothing
        
    Else
        '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
        If gsOpeCod = gOpeCGOpeCMACRetEfeMN Or gsOpeCod = gOpeCGOpeCMACRetEfeME Then
        
        Else
            txtCtaDest = ""
            lblDescCtaDest = ""
        End If

        lblDescObjDest = ""
        txtObjDest = ""
        txtMovDesc = ""
        txtmonto = "0.00"
        cboDocumento.ListIndex = -1
        txtNroDoc = ""
        OptDoc(0).value = False
        OptDoc(1).value = False
        lsNroDoc = ""
        lsDocumento = ""
        Set rsBill = Nothing
        Set rsMon = Nothing
        '***Fin Modificado por ELRO**********************************
        
        Unload Me
        
    End If
End If
End Sub
Private Sub Form_Load()
Dim oGen As New NConstSistemas
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF
Set objPista = New COMManejador.Pista

CentraForm Me
Me.Caption = gsOpeDesc
txtFecha = gdFecSis
Set rs = oOpe.CargaOpeDoc(gsOpeCod)
Do While Not rs.EOF
   cboDocumento.AddItem Mid(rs!cDocDesc & Space(100), 1, 100) & rs!nDocTpo
   rs.MoveNext
Loop
CambiaTamañoCombo cboDocumento
cmdDepositar.Visible = False
cmdRetirar.Visible = False
FraDeposito.Visible = False
FraRetiro.Visible = False
txtmonto.Locked = False

'***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D
ldFecCie = CDate(oGen.LeeConstSistema(gConstSistCierreMensualCont))
Set oGen = Nothing
'***Fin Modificado por ELRO**********************************

Select Case gsOpeCod
    Case gOpeCGOpeBancosDepEfecMN, gOpeCGOpeBancosDepEfecME
        txtCtaOrig.psRaiz = "Cuentas Contables"
        txtCtaOrig.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
        txtCtaDest.psRaiz = "Cuentas Contables"
        txtCtaDest.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
        cmdDepositar.Visible = True
        FraDeposito.Visible = True
        FraOrigen.Caption = "Origen"
        FraDestino.Caption = "Destino"
        txtmonto.Locked = True
    Case gOpeCGOpeBancosRetEfecMN, gOpeCGOpeBancosRetEfecME
        FraRetiro.Visible = True
        cmdRetirar.Visible = True
        FraOrigen.Caption = "Destino"
        FraDestino.Caption = "Origen"
        txtCtaOrig.psRaiz = "Cuentas Contables"
        txtCtaOrig.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
        txtCtaDest.psRaiz = "Cuentas Contables"
        txtCtaDest.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
    '***Modificada por ELRO el 20110923, según Acta 263-2011/TI-D
    Case gOpeCGOpeCMACDepEfeMN, gOpeCGOpeCMACDepEfeME
        txtCtaOrig.psRaiz = "Cuentas Contables"
        txtCtaOrig.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
        txtCtaDest.psRaiz = "Cuentas Contables"
        txtCtaDest.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
        cmdDepositar.Visible = True
        FraDeposito.Visible = True
        FraOrigen.Caption = "Origen"
        FraDestino.Caption = "Destino"
        txtmonto.Locked = True
    Case gOpeCGOpeCMACRetEfeMN, gOpeCGOpeCMACRetEfeME
        FraRetiro.Visible = True
        cmdRetirar.Visible = True
        FraOrigen.Caption = "Destino"
        FraDestino.Caption = "Origen"
        txtCtaOrig.psRaiz = "Cuentas Contables"
        txtCtaOrig.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
        txtCtaDest.psRaiz = "Cuentas Contables"
        txtCtaDest.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")
    '***Fin Modificada por ELRO**********************************
End Select

'***Modificada por ELRO el 20110923, según Acta 263-2011/TI-D
If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
       txtmonto.BackColor = &HC0FFC0
End If
Me.Label1 = Label1 & gsSimbolo & " :"
'***Fin Modificada por ELRO**********************************

End Sub


Private Sub OptDoc_Click(Index As Integer)
Dim oDocPago As clsDocPago
Dim oNContFunc As NContFunciones
Dim oCtasIF As NCajaCtaIF

Set oDocPago = New clsDocPago

Dim lsSubCuentaIF As String

lsNroDoc = ""
lsDocumento = ""
lsNroVoucher = ""

If Valida = False Then
    OptDoc(Index).value = False
    Exit Sub
End If
Select Case Index
    Case 0
        oDocPago.InicioCarta "", Mid(txtObjDest, 4, 13), gsOpeCod, _
                            gsOpeDesc, txtMovDesc, "", CCur(txtmonto), gdFecSis, _
                            lsNombreBanco, lsCuenta, gsNomCmac, "", ""
        If oDocPago.vbOk Then
            lnTipoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsDocumento = oDocPago.vsFormaDoc
            cmdRetirar.SetFocus
        Else
            lsNroDoc = ""
            lsDocumento = ""
            lsNroVoucher = ""
            OptDoc(Index).value = False
            Set oDocPago = Nothing
            Exit Sub
        End If
    Case 1
        Set oNContFunc = New NContFunciones
        Set oCtasIF = New NCajaCtaIF

        lsNroVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1), gsCodAge)
        lsSubCuentaIF = oCtasIF.SubCuentaIF(Mid(txtObjDest, 4, 13))

        oDocPago.InicioCheque "", True, Mid(txtObjDest, 4, 13), gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, CCur(txtmonto), _
                        gdFecSis, "", lsSubCuentaIF, lsNombreBanco, lsCuenta, lsNroVoucher, True, gsCodAge, Mid(txtObjDest, 18, 10), , Mid(txtObjDest, 1, 2), Mid(txtObjDest, 4, 13), Mid(txtObjDest, 18, 10) 'EJVG20121130
        Set oNContFunc = Nothing
        If oDocPago.vbOk Then
            lsNroDoc = ""
            lsDocumento = ""
            lsNroVoucher = ""
            lnTipoDoc = Val(oDocPago.vsTpoDoc)
            lsNroDoc = oDocPago.vsNroDoc
            lsDocumento = oDocPago.vsFormaDoc
            lsNroVoucher = oDocPago.vsNroVoucher
            cmdRetirar.SetFocus
        Else
            OptDoc(Index).value = False
            Set oDocPago = Nothing
            Exit Sub
        End If
End Select
Set oDocPago = Nothing

End Sub

Private Sub txtCtaDest_EmiteDatos()
Dim lsRaiz As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
txtObjDest = ""
lblDescObjDest = ""
lblDescCtaDest = txtCtaDest.psDescripcion
If txtCtaDest <> "" Then
    Set rs = AsignaCtaObj(txtCtaDest, lsRaiz)
    txtObjDest.psRaiz = lsRaiz
    txtObjDest.rs = rs
    'txtObjDest.SetFocus
End If
End Sub
Private Sub txtCtaOrig_EmiteDatos()
lblDescCtaOrig = txtCtaOrig.psDescripcion
End Sub

Private Sub txtmonto_GotFocus()
fEnfoque txtmonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtmonto, KeyAscii)
If KeyAscii = 13 Then
    If cmdDepositar.Visible Then cmdDepositar.SetFocus
    If cmdRetirar.Visible Then cmdRetirar.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
If Trim(Len(txtmonto)) = 0 Then txtmonto = 0
txtmonto = Format(txtmonto, "#,#0.00")
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If FraDeposito.Visible Then
        cmdefectivo.SetFocus
    End If
    If FraRetiro.Visible Then
        txtmonto.SetFocus
    End If
    
End If
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtmonto.SetFocus
End If
End Sub
Private Function AsignaCtaObj(ByVal psCtaContCod As String, ByRef lsRaiz As String) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
'Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    lsRaiz = ""
    lsFiltro = ""
    lnTipoObj = Val(rs1!cObjetoCod)
    Select Case Val(rs1!cObjetoCod)
        Case ObjCMACAgencias
            Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
        Case ObjCMACAgenciaArea
            lsRaiz = "Unidades Organizacionales"
            Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
        Case ObjCMACArea
            Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
        Case ObjEntidadesFinancieras
'            lsRaiz = "Cuentas de Entidades Financieras"
             lsRaiz = ""
            'Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
            Set rs = oCtaIf.CargaCtasIF(Mid(psCtaContCod, 3, 1), rs1!cCtaObjFiltro)
        Case ObjDescomEfectivo
            lsRaiz = "Denominación"
            Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
        Case ObjPersona
            Set rs = Nothing
        Case Else
            Set rs = GetObjetos(Val(rs1!cObjetoCod))
    End Select
End If
rs1.Close
Set rs1 = Nothing
Set AsignaCtaObj = rs

Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Function
Private Sub txtObjDest_EmiteDatos()
lsNombreBanco = ""
lsCuenta = ""
lblDescObjDest = txtObjDest.psDescripcion
If lblDescObjDest <> "" Then
    If Len(txtObjDest) > 15 Then
         lsNombreBanco = oCtaIf.NombreIF(Mid(txtObjDest, 4, 13))
         lsCuenta = oCtaIf.EmiteTipoCuentaIF(Mid(txtObjDest, 18, 10)) + " " + txtObjDest.psDescripcion
         lblDescObjDest = lsNombreBanco + " " + lsCuenta
    End If
    txtMovDesc.SetFocus
End If
End Sub
Private Sub txtObjDest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

