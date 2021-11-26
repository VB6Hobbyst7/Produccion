VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DB786848-D4E8-474E-A2C2-DCBC1D43DA22}#2.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmSegTarjetaAfiliacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Afiliación de Seguro de Tarjeta de Débito"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmSegTarjetaAfiliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   4440
      TabIndex        =   12
      Top             =   3480
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos de Afiliación"
      TabPicture(0)   =   "frmSegTarjetaAfiliacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraPorOpe"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraPorOpc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame FraPorOpc 
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   5295
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Buscar"
            Height          =   345
            Left            =   3840
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.TextBox txtNumTarjeta 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            MaxLength       =   16
            TabIndex        =   25
            Top             =   390
            Width           =   2895
         End
         Begin VB.CommandButton CmdLecTarj 
            Caption         =   "Leer Tarjeta"
            Height          =   345
            Left            =   3840
            TabIndex        =   22
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label4 
            Caption         =   "Tarjeta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   400
            Width           =   615
         End
      End
      Begin VB.Frame FraPorOpe 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   5055
         Begin VB.Frame Frame1 
            Caption         =   " ¿El cliente desea afiliarse? "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   2295
            Begin VB.OptionButton optAfiliacion 
               Caption         =   "Si"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   17
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optAfiliacion 
               Caption         =   "No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   1320
               TabIndex        =   16
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.Label Label1 
            Caption         =   "El cliente no posee seguro de tarjeta, ¿desea afiliarlo?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   120
            Width           =   4095
         End
         Begin VB.Label Label2 
            Caption         =   "Nº Sol.Mes:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblNroSolicMes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4560
            TabIndex        =   18
            Top             =   690
            Width           =   450
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   1560
         Width           =   5055
         Begin VB.CheckBox chkPrimaAnual 
            Caption         =   "Anual"
            Height          =   255
            Left            =   2400
            TabIndex        =   28
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            TabIndex        =   11
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox TxtAge 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1770
            MaxLength       =   2
            TabIndex        =   10
            Top             =   600
            Width           =   345
         End
         Begin VB.TextBox TxtProd 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2130
            MaxLength       =   3
            TabIndex        =   9
            Text            =   "232"
            Top             =   600
            Width           =   435
         End
         Begin VB.TextBox TxtCuenta 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   8
            Top             =   600
            Width           =   1200
         End
         Begin VB.TextBox txtCMAC 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "109"
            Top             =   600
            Width           =   435
         End
         Begin VB.Label txtNumCert 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Cuenta Debitar:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   640
            Width           =   1215
         End
         Begin VB.Label Label5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblImporte 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   4
            Top             =   960
            Width           =   930
         End
         Begin VB.Label Label6 
            Caption         =   "Importe Prima S/.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Certificado:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   280
            Width           =   855
         End
      End
   End
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   360
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   810
   End
End
Attribute VB_Name = "frmSegTarjetaAfiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTarjetaAfiliacion
'** Descripción : Formulario para afiliar seguros a las tarjetas con las que es están
'**               haciendo operaciones creado segun TI-ERS029-2013
'** Creación : JUEZ, 20130515 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oNSegTar As COMNCaptaGenerales.NCOMSeguros
Dim oDSegTar As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset
Dim sNumTarj As String
Dim fsPersCod As String
Dim fnParamNroSolicMes As Integer
Dim fnMontoPrima As Double
Dim fnMontoPrimaAnual As Double 'APRI20181023 ERS071-2018
Dim sPVV As String 'JUEZ 20150112
Dim sPVVOrig As String 'JUEZ 20150112

Public Sub Inicio(ByVal psNumtarjeta As String)
FraPorOpc.Visible = False
FraPorOpe.Visible = True
Call IniciaAfiliacion(psNumtarjeta, 1)
'Me.Show 1
End Sub

'JUEZ 20150112 *******************************************
Public Sub InicioOpc()
FraPorOpc.Visible = True
FraPorOpe.Visible = False
optAfiliacion.item(1) = True
HabilitaControles False
Me.Show 1
End Sub

Private Sub IniciaAfiliacion(ByVal psNumtarjeta As String, pnTipoAfiliacion As Integer)
sNumTarj = psNumtarjeta
    
    Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
    
        Set rs = oDSegTar.RecuperaSegTarjetaParametro(100)
        fnParamNroSolicMes = rs!nParamValor
     
    If pnTipoAfiliacion = 1 Then
        Set rs = oDSegTar.RecuperaNroSolicitudAfiliacionMes(sNumTarj, CInt(Right(gdFecSis, 4)), CInt(Mid(gdFecSis, 4, 2)))
        If Not rs.EOF Then
            If rs!nNroSolicMes < fnParamNroSolicMes Then
                Call oNSegTar.InsertaNroSolicitudAfiliacionMes(sNumTarj, gdFecSis)
                lblNroSolicMes.Caption = CStr(CInt(rs!nNroSolicMes) + 1)
            ElseIf rs!nNroSolicMes = fnParamNroSolicMes Then
                Call oNSegTar.InsertaNroSolicitudAfiliacionMes(sNumTarj, gdFecSis)
                Exit Sub
            Else
                Exit Sub
            End If
        Else
            lblNroSolicMes.Caption = "1"
            Call oNSegTar.InsertaNroSolicitudAfiliacionMes(sNumTarj, gdFecSis)
        End If
    End If
    
    Set rs = oDSegTar.RecuperaSegTarjetaParametro(101)
    fnMontoPrima = rs!nParamValor
    
    'APRI20181023 ERS071-2018
    Set rs = oDSegTar.RecuperaSegTarjetaParametro(106)
    fnMontoPrimaAnual = rs!nParamValor
    'END APRI
    
    TxtAge.Text = gsCodAge
    lblImporte.Caption = Format(fnMontoPrima, "#0.00") & " "
    txtNumCert.Caption = oDSegTar.ObtenerSegTarjetaNumCertificado
    
    Set oNSegTar = Nothing
    Set oDSegTar = Nothing
    If pnTipoAfiliacion = 1 Then
        HabilitaControles False
        Me.Show 1
    End If
End Sub
'END JUEZ ************************************************

Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
'txtNumCert.Enabled = pbHabilita
TxtAge.Enabled = pbHabilita
TxtProd.Enabled = pbHabilita
TxtCuenta.Enabled = pbHabilita
cmdBuscar.Enabled = pbHabilita
Me.txtNumTarjeta.Enabled = Not pbHabilita
chkPrimaAnual.Enabled = pbHabilita 'APRI20181023 ERS071-2018
'cmdAceptar.Enabled = pbHabilita
End Sub
'APRI20181023 ERS071-2018
Private Sub chkPrimaAnual_Click()
    If chkPrimaAnual.value Then
        lblImporte.Caption = Format(fnMontoPrimaAnual, "#0.00") & " "
    Else
        lblImporte.Caption = Format(fnMontoPrima, "#0.00") & " "
    End If
End Sub
'END APRI

Private Sub CmdAceptar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNCapGen As COMNCaptaGenerales.NCOMCaptaGenerales
Dim oNContFunc As COMNContabilidad.NCOMContFunciones
Dim rsCuenta As ADODB.Recordset
Dim nMonto As Double
Dim sCuenta As String
Dim sProd As String
Dim nTpoPrograma As String
Dim sMovNro As String
Dim sCodOpe As CaptacOperacion
Dim lsmensaje As String
Dim lsBoleta As String
Dim nSaldo As Double

If optAfiliacion.item(0) Then
    Unload Me
    Exit Sub
End If

If ValidaDatos Then
    nMonto = CDbl(lblImporte.Caption)
    sCuenta = txtCMAC.Text & TxtAge.Text & TxtProd.Text & TxtCuenta.Text
    sProd = TxtProd.Text
    
    Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCuenta = oNCapGen.GetDatosCuenta(sCuenta)
    Set oNCapGen = Nothing
    nTpoPrograma = rsCuenta!nTpoPrograma

'***********************COMENTADO APRI20171026 ERS028-2017*******************************
'    If sProd = gCapAhorros Then
'        If nTpoPrograma <> 0 And nTpoPrograma <> 5 And nTpoPrograma <> 6 And nTpoPrograma <> 8 Then
'            MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    Else
'        If nTpoPrograma <> 0 And nTpoPrograma <> 1 Then
'            MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
    'APRI20171004 ERS028-2017
    Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
    If Not oNSegTar.SepelioVerificaTpoPrograma(2, sProd, nTpoPrograma) Then
        MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oNSegTar = Nothing
    'END APRI
        
    If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
        'JUEZ 20150331 ***********************************
        'Dim clsTC As comdcredito.DCOMCredito
        'Dim nTC As Double
        'Set clsTC = New comdcredito.DCOMCredito
        'nTC = clsTC.DevolverTCMoneda(gdFecSis)!nVenta
        'Set clsTC = Nothing
        'nMonto = Round(nMonto / nTC, 2)
        Dim ObjTc As COMDConstSistema.NCOMTipoCambio
        Dim nTC As Double
        Set ObjTc = New COMDConstSistema.NCOMTipoCambio
        nTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoMes)
        Set ObjTc = Nothing
        nMonto = Round(nMonto / nTC, 2)
        'END JUEZ ****************************************
    End If
    
    If sProd = gCapAhorros Then
        sCodOpe = gAhoCargoAfilSegTarjeta
    Else
        sCodOpe = gCTSCargoAfilSegTarjeta
    End If
    
    Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    If oDSegTar.ValidaExisteRegistroNroCertificado(Trim(txtNumCert.Caption)) Then
        MsgBox "El Número de Certificado ya fue registrado anteriormente", vbInformation, "Aviso"
        'txtNumCert.SetFocus
        Exit Sub
    End If
    'If Not oDSegTar.ValidaNroCertificadoRemesaAgencia(Trim(txtNumCert.Text), gsCodAge) Then
    '    MsgBox "El numero de certificado ingresado no pertenece al rango asignado para la agencia", vbInformation, "Aviso"
    '    txtNumCert.SetFocus
    '    Exit Sub
    'End If
    Set oDSegTar = Nothing
    
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If oNCapMov.ValidaSaldoCuenta(sCuenta, nMonto, gAhoCargoAfilSegTarjeta) Then
        
        If MsgBox("Se va a realizar el Cargo a la cuenta por Afiliación de Tarjeta, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        Set oNContFunc = New COMNContabilidad.NCOMContFunciones
        sMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oNContFunc = Nothing
        'JUEZ 20140710 ********************************************************
        'gnMovNro = 0
        'If sProd = gCapAhorros Then
        '    nSaldo = oNCapMov.CapCargoCuentaAho(sCuenta, nMonto, sCodOpe, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & sNumTarj, , , , , , , , , gsNomAge, sLpt, , , , , gsCodCMAC, , gsCodAge, , , , , , "", , lsmensaje, , , , , gbImpTMU, , , , , , gnMovNro)
        'Else
        '    nSaldo = oNCapMov.CapCargoCuentaCTS(sCuenta, nMonto, sCodOpe, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & sNumTarj, , , , , , , gsNomAge, sLpt, , , , , , , , , , , lsmensaje, lsBoleta, gbImpTMU, , , , , , gnMovNro)
        'End If
        
        nSaldo = oNCapMov.CapCargoCuentaSegTarjeta(sCuenta, nMonto, sCodOpe, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & sNumTarj, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsmensaje, gbImpTMU, Trim(txtNumCert.Caption), sNumTarj, gdFecSis, fsPersCod, lsBoleta, chkPrimaAnual.value) 'APRI20181023 ERS071-2018 ADD chkPrimaAnual.value
        
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation
            Exit Sub
        End If
        'If gnMovNro <> 0 Then
        '    Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
        '    Call oNSegTar.InsertaSegTarjetaAfiliacion(Trim(txtNumCert.Caption), sNumTarj, sCuenta, gdFecSis, sMovNro, gnMovNro, fsPersCod)
        
        '    lsBoleta = oNSegTar.ImprimeBoletaAfilicacionSeguroTarjeta(gnMovNro, sMovNro, gsNomAge, gbImpTMU)
        '    Set oNSegTar = Nothing
            
        Dim nFicSal As Integer
        Do
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
            Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
            Close #nFicSal
        Loop Until MsgBox("¿Desea reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbNo
        '    gnMovNro = 0
        'End If
        'END JUEZ *************************************************************
        Unload Me
    Else
        MsgBox "La Cuenta NO posee saldo suficiente", vbInformation, "Aviso"
    End If
    Set oNCapMov = Nothing
End If
End Sub

Private Sub cmdBuscar_Click()
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim loCuentas As comdpersona.UCOMProdPersona
Dim rsCuentas As ADODB.Recordset
Dim fsPersNombre As String

    Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSegTar.RecuperaPersonaTarjeta(sNumTarj)
    If rs.EOF Then
        MsgBox "No se pueden obtener los datos de la persona", vbInformation, "Aviso"
        fsPersCod = ""
        Exit Sub
    End If
    fsPersCod = rs!cPersCod
    fsPersNombre = rs!cPersNombre
    Set oDSegTar = Nothing

    If Trim(fsPersCod) <> "" Then
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsCuentas = oDCapGen.GetCuentasPersona(fsPersCod, , True, True, , , , , True, "0,2")
    Set oDCapGen = Nothing
    End If
    
    If rsCuentas.RecordCount > 0 Then
        Set loCuentas = New comdpersona.UCOMProdPersona
        Set loCuentas = frmProdPersona.Inicio(fsPersNombre, rsCuentas)
        If loCuentas.sCtaCod <> "" Then
            TxtAge.Text = Mid(loCuentas.sCtaCod, 4, 2)
            TxtProd.Text = Mid(loCuentas.sCtaCod, 6, 3)
            TxtCuenta.Text = Mid(loCuentas.sCtaCod, 9, 10)
            TxtCuenta.SetFocus
        End If
        Set loCuentas = Nothing
    Else
        MsgBox "El cliente no tiene cuentas de ahorro activas", vbInformation, "Aviso"
        fsPersCod = ""
    End If
End Sub

'JUEZ 20150112 **********************************************
Private Sub cmdCargar_Click()
If Len(Trim(txtNumTarjeta.Text)) <> 16 Then
    MsgBox "Favor de ingresar correctamente el número de la tarjeta", vbInformation, "Aviso"
    txtNumTarjeta.SetFocus
    Exit Sub
End If

If Not frmATMCargaCuentas.ExisteTarjeta_Selec(txtNumTarjeta.Text) Then
    MsgBox "La Tarjeta N° " & txtNumTarjeta.Text & " no Existe, Intente otra vez", vbInformation, "Aviso"
    txtNumTarjeta.SetFocus
    Exit Sub
End If

If Not frmATMCargaCuentas.ValidaEstadoTarjeta_Selec(txtNumTarjeta.Text) Then
    MsgBox "La Tarjeta no esta activa", vbInformation, "Aviso"
    txtNumTarjeta.SetFocus
    Exit Sub
End If

Call CargaTarjeta(txtNumTarjeta.Text)
End Sub

Private Sub CmdLecTarj_Click()
Dim lsTarjeta As String

CmdLecTarj.Enabled = False
Me.Caption = "Afiliación de Seguro - PASE LA TARJETA"

lsTarjeta = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto, gnTimeOutAg), 2, 16)
txtNumTarjeta.Text = lsTarjeta
Me.Caption = "Afiliación de Seguro de Tarjeta de Débito"

If lsTarjeta = "" Then
    lsTarjeta = ""
    MsgBox "No hay conexion con el PINPAD o no paso la tarjeta, Intente otra vez", vbInformation, "Aviso"
    CmdLecTarj.Enabled = True
    Exit Sub
End If

If Not frmATMCargaCuentas.ExisteTarjeta_Selec(lsTarjeta) Then
    MsgBox "La Tarjeta N° " & lsTarjeta & " no Existe, Intente otra vez", vbInformation, "Aviso"
    lsTarjeta = ""
    CmdLecTarj.Enabled = True
    Exit Sub
End If

If Not frmATMCargaCuentas.ValidaEstadoTarjeta_Selec(lsTarjeta) Then
    lsTarjeta = ""
    MsgBox "La Tarjeta no esta activa", vbInformation, "Aviso"
    CmdLecTarj.Enabled = True
    Exit Sub
End If

If Left(lsTarjeta, 3) <> "ERR" Then
    CargaTarjeta (lsTarjeta)
End If
End Sub

Private Sub CargaTarjeta(ByVal psNumTarj As String)
    Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    If Not oDSegTar.VerificaSegTarjetaAfiliacion(psNumTarj) Then
        CmdLecTarj.Enabled = False
        cmdCargar.Enabled = False
        HabilitaControles True
        Call IniciaAfiliacion(psNumTarj, 2)
    Else
        MsgBox "La tarjeta ya está afiliada al Seguro", vbInformation, "Aviso"
        CmdLecTarj.Enabled = True
        cmdCargar.Enabled = True
        HabilitaControles False
    End If
End Sub
'END JUEZ ***************************************************

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub optAfiliacion_Click(Index As Integer)
    If Index = 1 Then
        HabilitaControles True
        'txtNumCert.SetFocus
    Else
        HabilitaControles False
        'Me.txtNumCert.Text = ""
        TxtAge.Text = gsCodAge
        TxtCuenta.Text = ""
    End If
End Sub

Private Sub TxtAge_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        TxtProd.SetFocus
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

'Private Sub txtNumCert_KeyPress(KeyAscii As Integer)
'    KeyAscii = SoloNumeros(KeyAscii)
'    If KeyAscii = 13 Then
'        TxtCuenta.SetFocus
'    End If
'End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False

'If Trim(txtNumCert.Text) = "" Then
'    MsgBox "Debe ingresar el numero de certificado", vbInformation, "Aviso"
'    txtNumCert.SetFocus
'    Exit Function
'End If
If Trim(TxtAge.Text) = "" Or Trim(TxtProd.Text) = "" Or Trim(TxtCuenta.Text) = "" Then
    MsgBox "Debe ingresar correctamente el numero de cuenta", vbInformation, "Aviso"
    If TxtCuenta.Enabled Then TxtCuenta.SetFocus
    Exit Function
End If
If Trim(TxtProd.Text) <> gCapAhorros And Trim(TxtProd.Text) <> gCapCTS Then
    MsgBox "La Cuenta debe ser Ahorros o CTS", vbInformation, "Aviso"
    If TxtCuenta.Enabled Then TxtCuenta.SetFocus
    Exit Function
End If
If Trim(fsPersCod) = "" Then
    MsgBox "No hay datos de la persona", vbInformation, "Aviso"
    Exit Function
End If
'APRI20190405 MEJORA
If chkPrimaAnual.value And CDbl(lblImporte.Caption) = 0 Then
    MsgBox "Producto no aplica a prima anual.", vbInformation, "Alerta"
    Exit Function
End If
'END APRI

ValidaDatos = True
End Function

Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

'JUEZ 20150112 ****************************************
Private Sub txtNumTarjeta_Change()
    If Trim(txtNumTarjeta.Text) <> "" Then
        CmdLecTarj.Visible = False
        cmdCargar.Visible = True
    Else
        CmdLecTarj.Visible = True
        cmdCargar.Visible = False
    End If
End Sub

Private Sub txtNumTarjeta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        If cmdCargar.Enabled Then cmdCargar.SetFocus
    End If
End Sub
'END JUEZ *********************************************

Private Sub txtprod_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        TxtCuenta.SetFocus
    End If
End Sub
