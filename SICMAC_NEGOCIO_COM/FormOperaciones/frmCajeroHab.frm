VERSION 5.00
Begin VB.Form frmCajeroHab 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
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
      Height          =   1095
      Left            =   2880
      TabIndex        =   22
      Top             =   2640
      Width           =   3135
      Begin VB.TextBox txtMovdesc 
         Height          =   735
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   2970
      End
   End
   Begin VB.Frame fraMoneda 
      Caption         =   "Moneda"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   2595
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda Extranjera"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   2115
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda Nacional"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   2115
      End
   End
   Begin SICMACT.Usuario User 
      Left            =   900
      Top             =   3810
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4725
      TabIndex        =   9
      Top             =   3862
      Width           =   1335
   End
   Begin SICMACT.EditMoney txtImporte 
      Height          =   315
      Left            =   1050
      TabIndex        =   7
      Top             =   3885
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
      TabIndex        =   13
      Top             =   1575
      Width           =   5970
      Begin SICMACT.TxtBuscar txtBuscarCtaDest 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   16512
      End
      Begin SICMACT.TxtBuscar txtBuscarObjDest 
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
         TabIndex        =   15
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
         TabIndex        =   14
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
      TabIndex        =   10
      Top             =   390
      Width           =   5970
      Begin SICMACT.TxtBuscar txtBuscarCtaOrig 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   8388608
      End
      Begin SICMACT.TxtBuscar txtBuscarObjOrig 
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
         TabIndex        =   12
         Top             =   675
         Width           =   3930
      End
      Begin VB.Label lblDescCtaOrig 
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
         TabIndex        =   11
         Top             =   285
         Width           =   3930
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   3375
      TabIndex        =   8
      Top             =   3862
      Width           =   1365
   End
   Begin VB.CommandButton cmdDevolucion 
      Caption         =   "&Devolución"
      Height          =   360
      Left            =   3375
      TabIndex        =   20
      Top             =   3862
      Width           =   1365
   End
   Begin VB.CommandButton cmdTransCajero 
      Caption         =   "&Transferir"
      Height          =   360
      Left            =   3375
      TabIndex        =   19
      Top             =   3862
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      Left            =   240
      TabIndex        =   16
      Top             =   3922
      Width           =   615
   End
   Begin VB.Shape ShapeS 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   60
      Top             =   3870
      Width           =   2760
   End
End
Attribute VB_Name = "frmCajeroHab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim nmoneda As COMDConstantes.Moneda
'MIOL 20120914, SEGUN RQ12270 ***********************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'END MIOL *******************************************

'RIRO 20170420 ***
Private Sub activaControl(Optional pbActiva As Boolean = True)
    cmdAceptar.Enabled = pbActiva
    cmdSalir.Enabled = pbActiva
End Sub
'RIRO END  *******

Private Sub cmdAceptar_Click()
    Dim oCon As COMNContabilidad.NCOMContFunciones  'NContFunciones
    Dim sUsuarioOrigen As String, sUsuarioDestino As String
    Dim lsCtaContDebe As String
    Dim lsCtaContHaber As String
    Dim lsMovNro As String
    Dim lnImporte As Double
    Dim oCajero As COMNCajaGeneral.NCOMCajero  'nCajero

    Dim lsCadImp As String

    If Valida = False Then Exit Sub
        Set oCajero = New COMNCajaGeneral.NCOMCajero  'nCajero
        Set oCon = New COMNContabilidad.NCOMContFunciones  'NContFunciones
        lnImporte = CDbl(txtImporte.Text)

        If MsgBox("Desea Realizar la Habilitacion A : " & vbCrLf & "Usuario :[" & txtBuscarObjDest & "] " & Trim(lblDescObjDest), vbYesNo + vbInformation, "Aviso") = vbYes Then
           
        'MIOL 20120914, SEGUN RQ12270 **************************************
        If gsOpeCod = "901013" Then
            Set loVistoElectronico = New frmVistoElectronico
            lbVistoVal = loVistoElectronico.Inicio(3, gsOpeCod)
            If lbVistoVal = False Then
                Unload Me
                Exit Sub
            End If
        End If
        'END MIOL **********************************************************
           
        Dim sAgeCod As String, sAreaCod As String
        sUsuarioOrigen = txtBuscarObjOrig.Text
        sUsuarioDestino = txtBuscarObjDest.Text
        lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        sAgeCod = gsCodAge
        sAreaCod = User.AreaCod
        
        activaControl False 'RIRO 20170420 ***
          If sUsuarioDestino <> "" Then 'pti1 02/10/2018
                If oCajero.GrabaHabilitaAgencia(lsMovNro, gsOpeCod, txtMovdesc, lnImporte, _
                    sAreaCod, sAgeCod, nmoneda, sUsuarioOrigen, sUsuarioDestino) = 0 Then
                    Dim oContImp As COMNContabilidad.NCOMContImprimir  'NContImprimir
                    Dim lbok As Boolean
                    Set oContImp = New COMNContabilidad.NCOMContImprimir
                    lbok = True
                    User.Inicio gsCodUser
                    lsCadImp = oContImp.ImprimeBoletahabilitacion(lblTitulo.Caption, "HABILITACION EN EFECTIVO", _
                                                                  txtBuscarObjOrig, lblDescObjOrig.Caption, txtBuscarObjDest.Text, lblDescObjDest.Caption, nmoneda, gsOpeCod, _
                                                                  lnImporte, gsNomAge, lsMovNro, sLpt, gsCodCMAC, gbImpTMU)
                       
                     Do While lbok
                        nFicSal = FreeFile
                        Open sLpt For Output As nFicSal
                                Print #nFicSal, lsCadImp & Chr$(12)
                                Print #nFicSal, ""
                        Close #nFicSal
                           
                        If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                           lbok = False
                        End If
                    Loop
                       
                       
                    Set oContImp = Nothing
                    Set oCajero = Nothing
                    Set oCon = Nothing
                    If gsOpeCod = gOpeBoveAgeHabCajero Then
                        User.Inicio gsCodUser
                        If txtBuscarObjDest.Enabled Then
                           txtBuscarObjDest.SetFocus
                           txtBuscarObjDest = ""
                           lblDescObjDest = ""
                        Else
                           optMoneda(0).SetFocus
                        End If
                     ElseIf gsOpeCod = gOpeHabCajDevABove Then
                        Else
                           txtBuscarObjDest = ""
                           lblDescObjDest = ""
                        End If
                   
                     txtImporte = 0
                     txtMovdesc = ""
                
                 End If
            Else
             MsgBox "La transferencia no se realizó, por favor seleccione el usuario destino y vuelva a intentarlo", vbInformation, "Aviso"
            End If
        activaControl 'RIRO 20170420 ***
    End If
End Sub

Private Function Valida() As Boolean
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
        MsgBox "Usuario destino no se encuentra definido", vbInformation, "Aviso"
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

Private Sub cmdDevolucion_Click()
    Dim oCon As COMNContabilidad.NCOMContFunciones  'NContFunciones
    Dim lsPersCodUser As String
    Dim lsCtaContDebe As String
    Dim lsCtaContHaber As String
    Dim lsMovNro As String
    Dim lsMovNroReg As String
    Dim lnMovNroReg As Long
    Dim lnImporte As Currency
    Dim oCajero As COMNCajaGeneral.NCOMCajero  'nCajero
    Dim rsBill As ADODB.Recordset
    Dim rsMon As ADODB.Recordset

    Dim lsCadImp As String

    Set oCajero = New COMNCajaGeneral.NCOMCajero   'nCajero
    Set oCon = New COMNContabilidad.NCOMContFunciones  'NContFunciones
    lnImporte = txtImporte.Text

    frmCajaGenEfectivo.RegistroEfectivo False
    If frmCajaGenEfectivo.lbok Then
        Set rsBill = frmCajaGenEfectivo.rsBilletes
        Set rsMon = frmCajaGenEfectivo.rsMonedas
        lsMovNroReg = frmCajaGenEfectivo.MovNro
        lnMovNroReg = frmCajaGenEfectivo.nMovNro
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
        
            Dim oContImp As COMNContabilidad.NCOMContImprimir   'NContImprimir
            Dim lbok As Boolean
            Set oContImp = New COMNContabilidad.NCOMContImprimir
            lbok = True
            User.Inicio gsCodUser
            
            lsCadImp = oContImp.ImprimeBoletahabilitacion("DEVOLUCIONES", "DEVOLUCION EN EFECTIVO", _
                       txtBuscarObjOrig.Text, lblDescObjOrig.Caption, txtBuscarObjDest.Text, lblDescObjDest.Caption, Mid(gsOpeCod, 3, 1), gsOpeCod, _
                       lnImporte, gsNomAge, lsMovNro, sLpt)
            
            Do While lbok
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                        Print #nFicSal, lsCadImp & Chr$(12)
                        Print #nFicSal, ""
                Close #nFicSal
                If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbok = False
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oGen As COMDConstSistema.DCOMGeneral  'DGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    User.Inicio gsCodUser
    Me.Caption = gsOpeDesc
    txtImporte.psSoles IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, True, False)
    cmdAceptar.Visible = False
    cmdDevolucion.Visible = False
    cmdTransCajero.Visible = False
    Select Case gsOpeCod
        Case gOpeBoveAgeHabCajero
            cmdAceptar.Visible = True
            lblTitulo = "HABILITACION A CAJERO"
            FraOrigen.Caption = "Boveda Agencia Origen"
            FraDestino.Caption = "Usuario Destino"
            txtBuscarCtaOrig.Text = gsCodAge
            txtBuscarCtaOrig.Enabled = False
            lblDescCtaOrig.Caption = gsNomAge
            txtBuscarObjOrig.Text = gsUsuarioBOVEDA
            txtBuscarObjOrig.Enabled = False
            lblDescObjOrig.Caption = "BOVEDA - AGENCIA"
            txtBuscarCtaDest.Text = gsCodAge
            txtBuscarCtaDest.Enabled = False
            lblDescCtaDest.Caption = gsNomAge
            txtBuscarObjDest.psRaiz = "USUARIOS"
            txtBuscarObjDest.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge)
        
        Case gOpeHabCajDevABove
            cmdAceptar.Visible = True
            lblTitulo = "DEVOLUCION A BOVEDA"
            FraOrigen.Caption = "Usuario Origen"
            FraDestino.Caption = "Boveda Agencia Destino"
            txtBuscarCtaOrig.Text = gsCodAge
            txtBuscarCtaOrig.Enabled = False
            lblDescCtaOrig.Caption = gsNomAge
            txtBuscarObjOrig.Text = gsCodUser
            txtBuscarObjOrig.Enabled = False
            lblDescObjOrig.Caption = User.UserNom
            txtBuscarCtaDest.Text = gsCodAge
            txtBuscarCtaDest.Enabled = False
            lblDescCtaDest.Caption = gsNomAge
        
            txtBuscarObjDest.Text = gsUsuarioBOVEDA
            txtBuscarObjDest.Enabled = False
            lblDescObjDest = "BOVEDA - AGENCIA"
        
        Case gOpeHabCajTransfEfectCajeros
            cmdAceptar.Visible = True
            lblTitulo = "TRANSFERENCIA A CAJEROS"
            FraOrigen.Caption = "Usuario Origen"
            FraDestino.Caption = "Usuario Destino"
            txtBuscarCtaOrig.Text = gsCodAge
            txtBuscarCtaOrig.Enabled = False
            lblDescCtaOrig.Caption = gsNomAge
            txtBuscarObjOrig.Text = gsCodUser
            txtBuscarObjOrig.Enabled = False
            lblDescObjOrig.Caption = User.UserNom
            txtBuscarCtaDest.Text = gsCodAge
            txtBuscarCtaDest.Enabled = False
            lblDescCtaDest.Caption = gsNomAge
            txtBuscarObjDest.psRaiz = "USUARIOS"
            txtBuscarObjDest.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge, gsCodUser)
    End Select
    Set oGen = Nothing
    optMoneda(0).value = True
End Sub

Private Sub OptMoneda_Click(Index As Integer)
    Select Case Index
        Case 0
            txtImporte.BackColor = &H80000005
            nmoneda = gMonedaNacional
        Case 1
            txtImporte.BackColor = &HC0FFC0
            nmoneda = gMonedaExtranjera
    End Select
End Sub

Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMovdesc.Visible Then txtMovdesc.SetFocus
    End If
End Sub

Private Sub txtBuscarCtaDest_EmiteDatos()
    lblDescCtaDest = txtBuscarCtaDest.psDescripcion
End Sub

Private Sub txtBuscarObjDest_EmiteDatos()
    lblDescObjDest = txtBuscarObjDest.psDescripcion
    User.Inicio txtBuscarObjDest
    If lblDescObjDest <> "" Then
        If optMoneda(0).value Then
            optMoneda(0).SetFocus
        ElseIf optMoneda(1).value Then
            optMoneda(1).SetFocus
        End If
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdAceptar.Visible Then cmdAceptar.SetFocus
        If cmdDevolucion.Visible Then cmdDevolucion.SetFocus
        If cmdTransCajero.Visible Then cmdTransCajero.SetFocus
    End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtImporte.SetFocus
    End If
End Sub
