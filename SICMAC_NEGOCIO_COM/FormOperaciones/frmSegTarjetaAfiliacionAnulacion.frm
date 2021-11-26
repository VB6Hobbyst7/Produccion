VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DB786848-D4E8-474E-A2C2-DCBC1D43DA22}#2.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmSegTarjetaAfiliacionAnulacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anulación de Seguro de Tarjeta de Débito"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frmSegTarjetaAfiliacionAnulacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Top             =   3600
      Width           =   1170
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   3600
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
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
      TabPicture(0)   =   "frmSegTarjetaAfiliacionAnulacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNumTarjeta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdLecTarj"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame2 
         Height          =   2295
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   5055
         Begin VB.Label lblFecAfiliacion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3600
            TabIndex        =   20
            Top             =   675
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Fec. Afilia :"
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
            Left            =   2760
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblCuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   10
            Top             =   1395
            Width           =   2055
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
            TabIndex        =   15
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblImporte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   14
            Top             =   1755
            Width           =   930
         End
         Begin VB.Label Label6 
            Caption         =   "Importe Prima:"
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
            TabIndex        =   13
            Top             =   1785
            Width           =   1095
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
            TabIndex        =   12
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblNumCert 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   11
            Top             =   1035
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "DOI :"
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
            TabIndex        =   9
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Titular:"
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
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblTitular 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   7
            Top             =   315
            Width           =   3495
         End
         Begin VB.Label lblDOI 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   675
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   3960
         TabIndex        =   4
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Tarjeta :"
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
         Left            =   240
         TabIndex        =   17
         Top             =   540
         Width           =   615
      End
      Begin VB.Label lblNumTarjeta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   2850
      End
   End
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmSegTarjetaAfiliacionAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTarjetaAfiliacionAnulacion
'** Descripción : Formulario para anular afiliaciones seguros a las tarjetas de débito creado
'**               segun TI-ERS029-2013
'** Creación : JUEZ, 20130523 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim oNSegTar As COMNCaptaGenerales.NCOMSeguros
Dim oDSegTar As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset
Dim sOpeCod As String
Dim nMovNro As Long
Dim dFecAfiliacion As Date
Dim nProd As String
Dim sOpeCodExt As String
Dim sNumTarj As String
Dim sPVV As String
Dim sPVVOrig As String
Dim sPersCod As String 'JUEZ 20150510

Private Sub cmdAnular_Click()
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNContFunc As COMNContabilidad.NCOMContFunciones
'APRI20171018 ERS028-2017
Dim loVistoElectronico As New frmVistoElectronico
Dim clsCap As New COMDCaptaGenerales.DCOMCaptaMovimiento
Dim lnMovNro As Long
Dim sMsj As String
'END APRI
Dim sOpeCod As String
Dim sCuenta As String
Dim nMonto As Double
Dim nSaldo As Double
Dim sMovNro As String
Dim lsBoleta As String
Dim lsMsjAnula As String 'JUEZ 20150331
    
    
    'APRI20171018 ERS028-2017
    If Not loVistoElectronico.Inicio(22, "401596") Then
        Exit Sub
    End If
    
    Dim DerArrep As Integer
    Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSegTar.RecuperaSegTarjetaParametro(103)
    DerArrep = rs!nParamValor
    Set oDSegTar = Nothing
    'End APRI
    
    sCuenta = Trim(lblCuenta.Caption)
    nMonto = CDbl(Trim(lblImporte.Caption))
    
    Set oNContFunc = New COMNContabilidad.NCOMContFunciones
    sMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oNContFunc = Nothing
    
    'JUEZ 20150510 *****************************************************************************************
    If DateDiff("d", dFecAfiliacion, CDate(Format(gdFecSis, "yyyy-MM-dd"))) <= DerArrep Then 'APRI20171027 ADD DerArrep ERS028-2017
        If Year(dFecAfiliacion) = Year(gdFecSis) And Month(dFecAfiliacion) = Month(gdFecSis) Then
            If MsgBox("Se va a realizar la devolución del monto por afiliación, Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            If Mid(sCuenta, 6, 3) = gCapAhorros Then
                sOpeCod = gAhoDepAfilSegTarjeta
            Else
                sOpeCod = gCTSDepAfilSegTarjeta
            End If
        Else
            If MsgBox("Se anulará una afiliación del mes anterior, Se va a realizar la devolución del monto por afiliación, Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            If Mid(sCuenta, 6, 3) = gCapAhorros Then
                sOpeCod = "200274"
            Else
                sOpeCod = "220214"
            End If
        End If
            
        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
            If Mid(sCuenta, 6, 3) = gCapAhorros Then
                nSaldo = oNCapMov.CapAbonoCuentaAho(sCuenta, nMonto, sOpeCod, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & sNumTarj, , , , , , , , , gsNomAge, sLpt, , , , , , gsCodCMAC, , , , , , , , , , , , , , , , , , gnMovNro)
            Else
                nSaldo = oNCapMov.CapAbonoCuentaCTS(sCuenta, nMonto, sOpeCod, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & sNumTarj, , , , , , , , , gsNomAge, sLpt, , , , gsCodCMAC, , , , , , , , , , gnMovNro)
            End If
        Set oNCapMov = Nothing
        If gnMovNro <> 0 Then
            If sOpeCod = "200274" Or sOpeCod = "220214" Then
                Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
                    oDSegTar.InsertaSegTarjetaAnulaDevPendiente lblNumCert.Caption, sPersCod, lblCuenta.Caption, gdFecSis, CDbl(lblImporte.Caption)
                Set oDSegTar = Nothing
            End If
            Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
                'Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 504)
                Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(lblNumCert.Caption, gnMovNro, 502) 'APRI20171027 ERS028-2017
            Set oNSegTar = Nothing
        Else
            MsgBox "Hubo un error en el proceso", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
        
        If Year(dFecAfiliacion) = Year(gdFecSis) And Month(dFecAfiliacion) = Month(gdFecSis) Then
            'If MsgBox("¿Se anulará la afiliación, pero se mostrará en la trama del mes, Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            'Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 505)
            sMsj = "¿Se anulará la afiliación, pero se mostrará en la trama del mes, Desea Continuar?" 'APRI20171027 ERS028-2017
        Else
            'If MsgBox("¿Se anulará la afiliación, Desea Continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            'Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 504)
            sMsj = "¿Se anulará la afiliación, Desea Continuar?" 'APRI20171027 ERS028-2017
        End If
        
        'APRI20171027 ERS028-2017
        If MsgBox(sMsj, vbQuestion + vbYesNo, "Aviso") = vbYes Then
            clsCap.AgregaMov sMovNro, "300154", "ANULACIÓN SEGURO TARJETA - N° CERTIFICADO: " & Trim(lblNumCert.Caption), gMovEstContabNoContable, gMovFlagVigente
            lnMovNro = clsCap.GetnMovNro(sMovNro)
            Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(lblNumCert.Caption, lnMovNro, 502)
        End If
        'END APRI
        
        Set oNSegTar = Nothing
    End If
    
'    If Year(dFecAfiliacion) = Year(gdFecSis) And Month(dFecAfiliacion) = Month(gdFecSis) Then
'        If DateDiff("d", dFecAfiliacion, CDate(Format(gdFecSis, "yyyy-MM-dd"))) <= 15 Then
'
'            If MsgBox("Se va a realizar la devolución del monto por afiliación,Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'            Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
'                If Mid(sCuenta, 6, 3) = gCapAhorros Then
'                    nSaldo = oNCapMov.CapAbonoCuentaAho(sCuenta, nMonto, sOpeCod, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & lblNumTarjeta.Caption, , , , , , , , , gsNomAge, sLpt, , , , , , gsCodCMAC, , , , , , , , , , , , , , , , , , gnMovNro)
'                Else
'                    nSaldo = oNCapMov.CapAbonoCuentaCTS(sCuenta, nMonto, sOpeCod, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & lblNumTarjeta.Caption, , , , , , , , , gsNomAge, sLpt, , , , gsCodCMAC, , , , , , , , , , gnMovNro)
'                End If
'            Set oNCapMov = Nothing
'            If gnMovNro <> 0 Then
'                Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
'                    Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 504)
'                Set oNSegTar = Nothing
'            Else
'                MsgBox "Hubo un error en el proceso", vbInformation, "Aviso"
'                Exit Sub
'            End If
'        Else
'            If MsgBox("¿Desea anular la afiliación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'            Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
'                'JUEZ 20150331 *******************************************
'                'Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 504)
'                lsMsjAnula = oNSegTar.GrabarAnulacionTarjetaSinDevolucion(nMovNro, sMovNro)
'                If lsMsjAnula <> "" Then
'                    MsgBox lsMsjAnula, vbInformation, "Aviso"
'                    Exit Sub
'                End If
'                'END JUEZ ************************************************
'            Set oNSegTar = Nothing
'        End If
'    Else
'        If MsgBox("¿Desea anular la afiliación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'        Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
'            Call oNSegTar.ActualizaEstadoSegTarjetaAfiliacion(nMovNro, 504)
'        Set oNSegTar = Nothing
'    End If
    'END JUEZ **********************************************************************************************
    
    MsgBox "Se anuló la afiliación de la tarjeta", vbInformation, "Aviso"
    Limpiar
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub CmdLecTarj_Click()
Me.Caption = "Anulación de Seguro - PASE LA TARJETA"

sNumTarj = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto, gnTimeOutAg), 2, 16)
lblNumTarjeta.Caption = Left(sNumTarj, 6) & " - - - - - - " & Right(sNumTarj, 4)
Me.Caption = "Anulación de Seguro de Tarjeta de Débito"

If sNumTarj = "" Then
    lblNumTarjeta.Caption = ""
    MsgBox "No hay conexion con el PINPAD o no paso la tarjeta, Intente otra vez", vbInformation, "Aviso"
    Limpiar
    Exit Sub
End If

If Not frmATMCargaCuentas.ExisteTarjeta_Selec(sNumTarj) Then
    lblNumTarjeta.Caption = ""
    MsgBox "La Tarjeta N° " & sNumTarj & " no Existe, Intente otra vez", vbInformation, "Aviso"
    Limpiar
    Exit Sub
End If

If Not frmATMCargaCuentas.ValidaEstadoTarjeta_Selec(sNumTarj) Then
    lblNumTarjeta.Caption = ""
    MsgBox "La Tarjeta no esta activa", vbInformation, "MENSAJE DEL SISTEMA"
    Limpiar
    Exit Sub
End If

If Left(sNumTarj, 3) <> "ERR" Then
    sPVV = frmATMCargaCuentas.RecuperaPVV(sNumTarj)
    sPVVOrig = frmATMCargaCuentas.RecuperaPVVOrig(sNumTarj)
    Call CargaDatos
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Limpiar()
sNumTarj = ""
lblNumTarjeta.Caption = ""
lbltitular.Caption = ""
sPersCod = ""
lblDOI.Caption = ""
lblNumCert.Caption = ""
lblCuenta.Caption = ""
lblImporte.Caption = ""
lblFecAfiliacion.Caption = ""
cmdAnular.Enabled = False
End Sub

Private Sub CargaDatos()
Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSegTar.RecuperaSegTarjetaAfiliacion(, sNumTarj)
    If Not (rs.EOF And rs.BOF) Then
        nMovNro = rs!nMovNroReg
        dFecAfiliacion = CDate(rs!dFecAfiliacion)
        lblFecAfiliacion.Caption = Format(rs!dFecAfiliacion, "dd/MM/yyyy") 'JUEZ 20150510
        lbltitular.Caption = rs!cPersNombre
        sPersCod = rs!cPersCod 'JUEZ 20150510
        lblDOI.Caption = rs!cPersIDnro
        lblNumCert.Caption = rs!cNumCertificado
        lblCuenta.Caption = rs!cCtaCodDebito
        lblImporte.Caption = Format(rs!nMontoPrima, "#,##0.00")
        cmdAnular.Enabled = True
    Else
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
    End If
Set oDSegTar = Nothing
End Sub


