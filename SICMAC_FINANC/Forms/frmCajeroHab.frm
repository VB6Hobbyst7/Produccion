VERSION 5.00
Begin VB.Form frmCajeroHab 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   2640
   ClientTop       =   2610
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroHab.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin Sicmact.Usuario User 
      Left            =   150
      Top             =   3705
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4725
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtMovdesc 
      Height          =   615
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2685
      Width           =   5970
   End
   Begin Sicmact.EditMoney txtImporte 
      Height          =   315
      Left            =   4290
      TabIndex        =   5
      Top             =   3345
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "0"
      Enabled         =   -1  'True
      BorderStyle     =   0
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
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   75
      TabIndex        =   10
      Top             =   1575
      Width           =   5970
      Begin Sicmact.TxtBuscar txtBuscarCtaDest 
         Height          =   345
         Left            =   240
         TabIndex        =   2
         Top             =   225
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   16512
      End
      Begin Sicmact.TxtBuscar txtBuscarObjDest 
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   585
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   16512
      End
      Begin VB.Label lblDescCtaDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1905
         TabIndex        =   12
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label lblDescObjDest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1905
         TabIndex        =   11
         Top             =   600
         Width           =   3930
      End
   End
   Begin VB.Frame FraOrigen 
      Caption         =   "Origen"
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
      Height          =   1110
      Left            =   75
      TabIndex        =   7
      Top             =   390
      Width           =   5970
      Begin Sicmact.TxtBuscar txtBuscarCtaOrig 
         Height          =   345
         Left            =   285
         TabIndex        =   0
         Top             =   278
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   8388608
      End
      Begin Sicmact.TxtBuscar txtBuscarObjOrig 
         Height          =   345
         Left            =   285
         TabIndex        =   1
         Top             =   668
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   8388608
      End
      Begin VB.Label lblDescObjOrig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1950
         TabIndex        =   9
         Top             =   675
         Width           =   3930
      End
      Begin VB.Label lblDescCtaOrig 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1950
         TabIndex        =   8
         Top             =   285
         Width           =   3930
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3375
      TabIndex        =   18
      Top             =   3720
      Width           =   1365
   End
   Begin VB.CommandButton cmddevolucion 
      Caption         =   "&Devolución"
      Height          =   360
      Left            =   3375
      TabIndex        =   17
      Top             =   3720
      Width           =   1365
   End
   Begin VB.CommandButton cmdTransCajero 
      Caption         =   "&Transferir"
      Height          =   360
      Left            =   3375
      TabIndex        =   16
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "HABILITACION A CAJEROS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1102
      TabIndex        =   15
      Top             =   30
      Width           =   3930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Glosa :"
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
      Left            =   90
      TabIndex        =   14
      Top             =   2295
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3480
      TabIndex        =   13
      Top             =   3375
      Width           =   615
   End
   Begin VB.Shape ShapeS 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   3300
      Top             =   3330
      Width           =   2760
   End
End
Attribute VB_Name = "frmCajeroHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe As DOperacion
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Private Sub cmdAceptar_Click()
Dim oCon As NContFunciones
Dim lsPersCodUser As String
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsMovNro As String
Dim lnImporte As Currency
Dim oCajero As nCajero


If Valida = False Then Exit Sub
Set oCajero = New nCajero
Set oCon = New NContFunciones
lnImporte = txtImporte.Text
If MsgBox("Desea Realizar la Habilitacion A : " & vbCrLf & "Usuario :[" & txtBuscarObjDest & "] " & Trim(lblDescObjDest), vbYesNo + vbInformation, "Aviso") = vbYes Then
    lsCtaContDebe = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaDebe, txtBuscarObjOrig)
    lsCtaContHaber = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaHaber, txtBuscarObjOrig)
    lsPersCodUser = User.PersCod
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaHabilitaAgencia(gsFormatoFecha, lsMovNro, gsOpeCod, txtMovdesc, lnImporte, _
                                lsCtaContHaber, Mid(Me.txtBuscarObjOrig, 1, 3), Mid(Me.txtBuscarObjOrig, 4, 2), _
                                lsCtaContDebe, lsPersCodUser, Trim(txtBuscarObjDest)) = 0 Then
        
        'ImprimeAsientoContable lsMovNro, , , , , , , lsPersCodUser, lnImporte
        Dim oContImp As NContImprimir
        Dim lbOk As Boolean
        Set oContImp = New NContImprimir
            lbOk = True
            User.Inicio gsCodUser
            Do While lbOk
                oContImp.ImprimeBoletahabilitacion lblTitulo, "HABILITACION EN EFECTIVO", _
                         txtBuscarObjOrig, lblDescObjOrig, txtBuscarObjDest, lblDescObjDest, Mid(gsOpeCod, 3, 1), gsOpeCod, _
                        lnImporte, gsNomAge, lsMovNro, "LPT1"
                        
                If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbOk = False
                End If
            Loop
        Set oContImp = Nothing
        Set oCajero = Nothing
        Set oCon = Nothing
            Select Case gsOpeCod
                 Case gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME
                        txtBuscarObjDest = ""
                        lblDescObjDest = ""
                        txtImporte = 0
                        txtMovdesc = ""
                        User.Inicio gsCodUser
                        txtBuscarObjDest.SetFocus
            End Select
    End If
End If

End Sub
Function Valida() As Boolean
Valida = True
If Len(Trim(txtBuscarCtaOrig)) = 0 Then
    MsgBox "Cuenta Contable de Origen no se definido", vbInformation, "Aviso"
    If txtBuscarCtaOrig.Enabled Then txtBuscarCtaOrig.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtBuscarObjOrig)) = 0 Then
    MsgBox "Objeto Origen no se encuentra definido", vbInformation, "Aviso"
    If txtBuscarObjOrig.Enabled Then txtBuscarObjOrig.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtBuscarCtaDest)) = 0 Then
    MsgBox "Cuenta Contable Destino no se encuenta definida", vbInformation, "Aviso"
    If txtBuscarCtaDest.Enabled Then txtBuscarCtaDest.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtBuscarObjDest)) = 0 Then
    MsgBox "Objeto de destino no se encuentra definido", vbInformation, "Aviso"
    If txtBuscarObjDest.Enabled Then txtBuscarObjDest.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtMovdesc)) = 0 Then
    MsgBox "Descripcion de Operacion no se ha Ingresado", vbInformation, "Aviso"
    If txtMovdesc.Enabled Then txtMovdesc.SetFocus
    Valida = False
    Exit Function
End If
If Val(txtImporte.Text) = 0 Then
    MsgBox "Importe de Operacion no se Ingresado", vbInformation, "Aviso"
    If txtImporte.Enabled Then txtImporte.SetFocus
    Valida = False
    Exit Function
End If



End Function

Private Sub cmddevolucion_Click()
Dim oCon As NContFunciones
Dim lsPersCodUser As String
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsMovNro As String
Dim lsMovNroReg As String
Dim lnMovNroReg As Long
Dim lnImporte As Currency
Dim oCajero As nCajero
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

Set oCajero = New nCajero
Set oCon = New NContFunciones
lnImporte = txtImporte.Text

frmCajaGenEfectivo.RegistroEfectivo False
If frmCajaGenEfectivo.lbOk Then
    Set rsBill = frmCajaGenEfectivo.rsBilletes
    Set rsMon = frmCajaGenEfectivo.rsMonedas
    lsMovNroReg = frmCajaGenEfectivo.MovNro
    lnMovNroReg = frmCajaGenEfectivo.nMovNro
    txtImporte = Format(frmCajaGenEfectivo.Total, "#,#0.00")
    Set frmCajaGenEfectivo = Nothing
Else
    Set frmCajaGenEfectivo = Nothing
    Exit Sub
End If
If Valida = False Then Exit Sub

If MsgBox("Desea Realizar la Devolución A : " & vbCrLf & "la Boveda de :[" & txtBuscarObjDest & "] " & Trim(lblDescObjDest), vbYesNo + vbInformation, "Aviso") = vbYes Then
    User.Inicio gsCodUser
    lsCtaContDebe = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaDebe, txtBuscarObjDest)
    lsCtaContHaber = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaHaber, txtBuscarObjDest)
    lsPersCodUser = User.PersCod
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaDevolucionABoveda(gsFormatoFecha, lsMovNro, gsOpeCod, txtMovdesc, lsCtaContDebe, _
                                lsCtaContHaber, rsBill, rsMon, Mid(txtBuscarObjDest, 1, 3), Mid(txtBuscarObjDest, 4, 2), _
                                lsPersCodUser, CCur(txtImporte), lnMovNroReg) = 0 Then
        
        'ImprimeAsientoContable lsMovNro, , , , , , , lsPersCodUser, lnImporte
        Dim oContImp As NContImprimir
        Dim lbOk As Boolean
        Set oContImp = New NContImprimir
            lbOk = True
            User.Inicio gsCodUser
            Do While lbOk
                oContImp.ImprimeBoletahabilitacion "DEVOLUCIONES", "DEVOLUCION EN EFECTIVO", _
                         txtBuscarObjOrig, lblDescObjOrig, txtBuscarObjDest, lblDescObjDest, Mid(gsOpeCod, 3, 1), gsOpeCod, _
                        lnImporte, gsNomAge, lsMovNro, "LPT1"
                        
                If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbOk = False
                End If
            Loop
        Set oContImp = Nothing
        
        Set oCajero = Nothing
        Set oCon = Nothing
        If MsgBox("Desea realizar otra habilitación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            txtImporte = 0
            txtMovdesc = ""
            User.Inicio gsCodUser
        Else
            Unload Me
        End If
    End If
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTransCajero_Click()
Dim oCon As NContFunciones
Dim lsPersCodOrig As String
Dim lsPersCodDest As String
Dim lsCtaContDebe As String
Dim lsCtaContHaber As String
Dim lsMovNro As String
Dim lnImporte As Currency
Dim oCajero As nCajero


If Valida = False Then Exit Sub
Set oCajero = New nCajero
Set oCon = New NContFunciones
lnImporte = txtImporte.Text
If MsgBox("Desea Realizar la Transferencia A : " & vbCrLf & "Usuario :[" & txtBuscarObjDest & "] " & Trim(lblDescObjDest), vbYesNo + vbInformation, "Aviso") = vbYes Then
    User.Inicio txtBuscarObjOrig
    lsPersCodOrig = User.PersCod
    lsCtaContDebe = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaDebe, User.cAreaCodAct & User.CodAgeAct)
    lsCtaContHaber = oCon.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaHaber, User.cAreaCodAct & User.CodAgeAct)
    
    User.Inicio txtBuscarObjDest
    lsPersCodDest = User.PersCod
    
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaTransferenciaCajero(gsFormatoFecha, lsMovNro, gsOpeCod, txtMovdesc, lnImporte, _
                                lsCtaContHaber, lsPersCodDest, _
                                lsCtaContDebe, lsPersCodOrig, Trim(txtBuscarObjOrig), Trim(txtBuscarObjDest)) = 0 Then
        
        Dim oContImp As NContImprimir
        Dim lbOk As Boolean
        Set oContImp = New NContImprimir
            lbOk = True
            User.Inicio gsCodUser
            Do While lbOk
                oContImp.ImprimeBoletahabilitacion lblTitulo, "TRANSFERENCIA EN EFECTIVO", _
                         txtBuscarObjOrig, lblDescObjOrig, txtBuscarObjDest, lblDescObjDest, Mid(gsOpeCod, 3, 1), gsOpeCod, _
                        lnImporte, gsNomAge, lsMovNro, "LPT1"
                        
                If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbOk = False
                End If
            Loop
        Set oContImp = Nothing
        Set oCajero = Nothing
        Set oCon = Nothing
        If MsgBox("Desea realizar otra transferencia??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            txtBuscarObjDest = ""
            lblDescObjDest = ""
            txtImporte = 0
            txtMovdesc = ""
            User.Inicio gsCodUser
            txtBuscarObjDest.SetFocus
        Else
            Unload Me
        End If
    End If
End If

End Sub

Private Sub Form_Load()
Dim oGen As dgeneral
Set oOpe = New DOperacion
Set oGen = New dgeneral
CentraForm Me
User.Inicio gsCodUser
Me.Caption = gsOpeDesc
txtImporte.psSoles IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, True, False)
cmdAceptar.Visible = False
cmddevolucion.Visible = False
cmdTransCajero.Visible = False
Select Case gsOpeCod
    Case gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME
        cmdAceptar.Visible = True
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
        FraOrigen.Caption = "Boveda Agencia Origen"
        FraDestino.Caption = "Usuario Destino"
        txtBuscarCtaOrig.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtBuscarCtaDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        txtBuscarObjOrig.rs = GetObjetosOpeCta(gsOpeCod, "0", txtBuscarCtaOrig, "", , gsCodAge)
        txtBuscarObjDest.psRaiz = "USUARIOS"
        txtBuscarObjDest.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge)
    Case gOpeHabCajDevABoveMN, gOpeHabCajDevABoveME
        cmddevolucion.Visible = True
        lblTitulo = "DEVOLUCION A BOVEDA"
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
        FraOrigen.Caption = "Usuario Origen"
        FraDestino.Caption = "Boveda Agencia Destino"
        txtBuscarCtaOrig.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtBuscarCtaDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        
        txtBuscarObjOrig.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge, gsCodUser, False)
        txtBuscarObjDest.rs = GetObjetosOpeCta(gsOpeCod, "0", txtBuscarCtaOrig, "", , gsCodAge)
        
    Case gOpeHabCajTransfEfectCajerosMN, gOpeHabCajTransfEfectCajerosME
        cmdTransCajero.Visible = True
        lblTitulo = "TRANSFERENCIA A CAJEROS"
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
        FraOrigen.Caption = "Usuario Origen"
        FraDestino.Caption = "Usuario Destino"
        txtBuscarCtaOrig.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "H")
        txtBuscarCtaDest.rs = oOpe.EmiteOpeCtasNivel(gsOpeCod, "D")
        
        txtBuscarObjOrig.psRaiz = "USUARIOS"
        txtBuscarObjOrig.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge, gsCodUser, False)
        txtBuscarObjDest.psRaiz = "USUARIOS"
        txtBuscarObjDest.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge, gsCodUser)
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oOpe = New DOperacion
End Sub
Private Sub txtBuscarCtaDest_EmiteDatos()
    lblDescCtaDest = txtBuscarCtaDest.psDescripcion
End Sub
Private Sub txtBuscarCtaOrig_EmiteDatos()
lblDescCtaOrig = txtBuscarCtaOrig.psDescripcion
End Sub
Private Sub txtBuscarObjDest_EmiteDatos()
    lblDescObjDest = txtBuscarObjDest.psDescripcion
    User.Inicio txtBuscarObjDest
    If lblDescObjDest <> "" Then
        If txtMovdesc.Visible Then txtMovdesc.SetFocus
    End If
End Sub

Private Sub txtBuscarObjOrig_EmiteDatos()
    lblDescObjOrig = txtBuscarObjOrig.psDescripcion
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmdAceptar.Visible Then cmdAceptar.SetFocus
    If cmddevolucion.Visible Then cmddevolucion.SetFocus
    If cmdTransCajero.Visible Then cmdTransCajero.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtImporte.SetFocus
End If
End Sub
