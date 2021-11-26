VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenHabilitacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   1635
   ClientTop       =   2640
   ClientWidth     =   8925
   Icon            =   "frmCajaGenHabilitacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8685
      Begin VB.TextBox txtMonto 
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
         Left            =   5880
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   3615
         Width           =   1740
      End
      Begin VB.Frame FraGlosa 
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
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   2400
         TabIndex        =   29
         Top             =   2520
         Width           =   2535
         Begin VB.TextBox txtMovDesc 
            Height          =   630
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Frame FraDocumento 
         Caption         =   "Documento"
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
         Height          =   1050
         Left            =   5040
         TabIndex        =   25
         Top             =   2520
         Visible         =   0   'False
         Width           =   2625
         Begin VB.TextBox txtNroDoc 
            Height          =   315
            Left            =   360
            TabIndex        =   27
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboDocumento 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   210
            Width           =   2370
         End
         Begin VB.Label Label7 
            Caption         =   "N° :"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame FraIF 
         Caption         =   "Entidad Financiera"
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
         Height          =   1140
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   8415
         Begin SICMACT.TxtBuscar txtIF 
            Height          =   345
            Left            =   810
            TabIndex        =   19
            Top             =   225
            Width           =   1470
            _extentx        =   2593
            _extenty        =   609
            appearance      =   1
            appearance      =   1
            font            =   "frmCajaGenHabilitacion.frx":030A
            appearance      =   1
            stitulo         =   ""
         End
         Begin SICMACT.TxtBuscar txtCtaIF 
            Height          =   345
            Left            =   780
            TabIndex        =   20
            Top             =   600
            Width           =   2595
            _extentx        =   4577
            _extenty        =   609
            appearance      =   1
            appearance      =   1
            font            =   "frmCajaGenHabilitacion.frx":0336
            appearance      =   1
            stitulo         =   ""
         End
         Begin VB.Label lblIF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   5955
         End
         Begin VB.Label lblCtaIF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   3375
            TabIndex        =   23
            Top             =   600
            Width           =   4875
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Entidad :"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta :"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   675
            Width           =   600
         End
      End
      Begin VB.Frame frameDestino 
         Caption         =   "Destino"
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
         Height          =   750
         Left            =   120
         TabIndex        =   11
         Top             =   1335
         Visible         =   0   'False
         Width           =   7380
         Begin SICMACT.TxtBuscar TxtAreaAgeDest 
            Height          =   330
            Left            =   1380
            TabIndex        =   12
            Top             =   225
            Width           =   1485
            _extentx        =   2619
            _extenty        =   582
            appearance      =   1
            appearance      =   1
            font            =   "frmCajaGenHabilitacion.frx":0362
            appearance      =   1
            stitulo         =   ""
         End
         Begin VB.Label lblAreaAgeDest 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2895
            TabIndex        =   14
            Top             =   225
            Width           =   4275
         End
         Begin VB.Label Label6 
            Caption         =   "Area - Agencia"
            Height          =   195
            Left            =   105
            TabIndex        =   13
            Top             =   255
            Width           =   1215
         End
      End
      Begin VB.Frame frameOrigen 
         Caption         =   "Origen"
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
         Height          =   765
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   7350
         Begin SICMACT.TxtBuscar txtAreaAgeOrig 
            Height          =   330
            Left            =   1410
            TabIndex        =   8
            Top             =   285
            Width           =   1485
            _extentx        =   2619
            _extenty        =   582
            appearance      =   1
            appearance      =   1
            font            =   "frmCajaGenHabilitacion.frx":038E
            appearance      =   1
            stitulo         =   ""
         End
         Begin VB.Label lblAreaAgeOrig 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2925
            TabIndex        =   10
            Top             =   285
            Width           =   4275
         End
         Begin VB.Label Label1 
            Caption         =   "Area Agencia :"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   315
            Width           =   1215
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
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   2115
         Begin VB.OptionButton optMoneda 
            Caption         =   "Moneda Extranjera"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   540
            Width           =   1695
         End
         Begin VB.OptionButton optMoneda 
            Caption         =   "Moneda Nacional"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   240
            Width           =   1635
         End
      End
      Begin MSMask.MaskEdBox txtFechaMov 
         Height          =   330
         Left            =   6210
         TabIndex        =   15
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   128
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
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
         Left            =   5160
         TabIndex        =   32
         Top             =   3720
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   390
         Left            =   5040
         Top             =   3600
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   5595
         TabIndex        =   17
         Top             =   225
         Width           =   540
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "HABILITACION ENTRE AGENCIAS "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   360
         Left            =   315
         TabIndex        =   16
         Top             =   135
         Width           =   5190
      End
   End
   Begin VB.CommandButton cmdEfectivo 
      Caption         =   "&Efectivo"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4215
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   4215
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   4200
      Width           =   1200
   End
End
Attribute VB_Name = "frmCajaGenHabilitacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lsAreCodCaja As String, sAreaNomCaja As String
'Dim lbSalir As Boolean
'Dim nmoneda As COMDConstantes.Moneda
'
'Dim fsOpeCodConfirmacion As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
'Dim fsOpeCodDeposito As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
'Dim fsCtaDebeConHab As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
'Dim fsCtaHaberConHab As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
'Dim fsCtaHaberDep As String 'Agregado por ELRO 20111102, según Acta 277-2011/TI-D
'
''EJVG20131111 ***
''Private Sub cmdAceptar_Click()
''Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral  'nCajaGeneral
''Dim oCon As COMNContabilidad.NCOMContFunciones  'NContFunciones
''Dim lnSaldoDisp As Double
''Dim rsBill As ADODB.Recordset
''Dim rsMon As ADODB.Recordset
''Dim lnTotalServicio As Double
''Dim lsMovNro As String
''Dim lsSubCta  As String
''Dim lsCadImp As String
''Dim nFicSal As Integer
''Dim cPrevio As previo.clsprevio
''
''Dim onCajaGeneral As clases.nCajaGeneral  'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Set onCajaGeneral = New clases.nCajaGeneral  'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
'''Dim lnMovNro, lnMovNro3 As Long 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D 'Comentado por ELRO el 20120522, según Acta N° 106-2012
''Dim oDMov As DMov 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Set oDMov = New DMov 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Dim bConfirmarHabilitacion, bConfirmacionDeposito As Boolean 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Dim lsMovNro2 As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Dim lsMovNro3 As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Dim oNContFunciones As clases.NContFunciones 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Set oNContFunciones = New NContFunciones 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''Dim lsCtafiltro As String 'Agregado por ELRO 20111024, según Acta 277-2011/TI-D
''
''Set rsBill = New ADODB.Recordset
''Set rsMon = New ADODB.Recordset
''Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
''Set oCon = New COMNContabilidad.NCOMContFunciones
''
''Dim lnMovNroHabEfe As Long 'Agregado por ELRO 20120522, según Acta N° 106-2012/TI-D
''Dim lnMovNroConHabEfe As Long 'Agregado por ELRO 20120522, según Acta N° 106-2012/TI-D
''Dim lnMovNroDepEfe As Long 'Agregado por ELRO 20120522, según Acta N° 106-2012/TI-D
''Dim oDCOMMov As DCOMMov 'Agregado por ELRO 20120523, según Acta N° 106-2012/TI-D
''Dim lsMovNroExt As String 'Agregado por ELRO 20120523, según Acta N° 106-2012/TI-D
''
''
''If valida = False Then Exit Sub
''
''Select Case gsOpeCod
''    Case gOpeBoveAgeHabEntreAge
''        If TxtAreaAgeDest = txtAreaAgeOrig Then
''            MsgBox "Agencia de Destino no puede ser la misma que de Origen", vbInformation, "Aviso"
''            If TxtAreaAgeDest.Enabled Then TxtAreaAgeDest.SetFocus
''            Exit Sub
''        End If
''    Case Else
''End Select
''
''
''frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, CDbl(txtMonto), nMoneda, False
''If frmCajaGenEfectivo.lbOk Then
''     Set rsBill = frmCajaGenEfectivo.rsBilletes
''     Set rsMon = frmCajaGenEfectivo.rsMonedas
''
''Else
''    MsgBox "Debe registrar correctamente la descomposición de efectivo de la habilitación", vbInformation, "Aviso"
''    Set frmCajaGenEfectivo = Nothing
''    Exit Sub
''End If
''Set frmCajaGenEfectivo = Nothing
''If (rsBill Is Nothing And rsMon Is Nothing) Then
''    MsgBox "Error en Ingreso de Billetaje", vbInformation, "Aviso"
''    Exit Sub
''End If
''
'''***Modificado por ELRO el 20111024, según Acta 277-2011/TI-D
''lsCtafiltro = oNContFunciones.GetFiltroObjetos(1, Me.txtIF, Me.txtCtaIF, False)
''If lsCtafiltro = "" And fsOpeCodDeposito <> "" Then
''    MsgBox "Esta cuenta contable " & Me.txtIF & " no esta registrado en CtaIFFiltro, comunicarse con TI", vbInformation, "Aviso"
''    Exit Sub
''ElseIf fsOpeCodDeposito = "" Then
''    lsCtafiltro = ""
''End If
''
''If oNContFunciones.verificarUltimoNivelCta(Me.txtIF & lsCtafiltro) = False And fsOpeCodDeposito <> "" Then
''   MsgBox "La cuenta contable " & Me.txtIF + lsCtafiltro & " no es de Ultimo Nivel, comunicarse con TI", vbInformation, "Aviso"
''   Exit Sub
''ElseIf fsOpeCodDeposito = "" Then
''    Me.txtIF = ""
''End If
'''***Fin Modificado por ELRO**********************************
''
''If MsgBox("Desea Realizar la Habilitación respectiva?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
''    lsMovNro = oCon.GeneraMovNro(txtFechaMov, gsCodAge, gsCodUser)
''    '***Modificado por ELRO el 20111024, según Acta 277-2011/TI-D
''    'If oCaja.GrabaHabEfectivo(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc,
''    '            CDbl(txtMonto), Trim(txtAreaAgeOrig), Trim(TxtAreaAgeDest),
''    '            "", Format$(txtFechaMov, gsFormatoFecha),
''    '            "", Format$(txtFechaMov, gsFormatoFecha),
''    '            "", 0, nmoneda, , lsAreCodCaja) = 0 Then
''
''    'Paso 1:Registra la habilitación del Efectivo
''    If oCaja.GrabaHabEfectivo(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc, _
''                              CDbl(txtMonto), Trim(txtAreaAgeOrig), _
''                              Trim(TxtAreaAgeDest), "", Format$(txtFechaMov, _
''                              gsFormatoFecha), "", Format$(txtFechaMov, _
''                              gsFormatoFecha), "", 0, nMoneda, , lsAreCodCaja, , _
''                              fsOpeCodDeposito, lnMovNroHabEfe) = 0 Then
''
''        '***Modifcado por ELRO el 20120522, según Acta N° 106-2012/TI-D
''        'If fsOpeCodDeposito <> "" Then
''        If fsOpeCodDeposito <> "" And lnMovNroHabEfe > 0 Then
''        '***Fin Modifcado por ELRO el 20120522*************************
''            Call onCajaGeneral.registrarSeguimientoHabiltacionAgenciaCajaGeneral(lsMovNro)
''
''            '***Comentado por ELRO según Acta N° 106-2012/TI-D
''            'lnMovNro = oDMov.GetnMovNro(lsMovNro)
''            '***Fin Comentado por ELRO************************
''            lsMovNro2 = oNContFunciones.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
''            rsBill.MoveFirst
''            If nMoneda = gMonedaNacional Then
''                rsMon.MoveFirst
''            End If
''
''            bConfirmarHabilitacion = False
''            bConfirmacionDeposito = False
''
''            '***Modificado por ELRO el 20120522, según Acta N° 106-2012/TI-D
''            'If onCajaGeneral.GrabaConfHabEfectivo(lsMovNro2, rsBill, rsMon,
''            '                                      fsOpeCodConfirmacion, txtMovDesc,
''            '                                      fsCtaDebeConHab, fsCtaHaberConHab, CDbl(txtMonto),
''            '                                      Trim(txtAreaAgeOrig), Trim(TxtAreaAgeDest),
''            '                                      lnMovNro, fsOpeCodDeposito, lnMovNroConHabEfe) = 0 Then
''            '***Fin Modificado por ELRO*****************************************
''
''            'Paso2:Registra la confirmacion de la Habilitación
''            'de Agencia a Caja General
''            Call onCajaGeneral.GrabaConfHabEfectivo(lsMovNro2, rsBill, rsMon, _
''                                                    fsOpeCodConfirmacion, txtMovDesc, _
''                                                    fsCtaDebeConHab, fsCtaHaberConHab, _
''                                                    CDbl(txtMonto), Trim(txtAreaAgeOrig), _
''                                                    Trim(TxtAreaAgeDest), lnMovNroHabEfe, _
''                                                    fsOpeCodDeposito, lnMovNroConHabEfe)
''            If lnMovNroConHabEfe > 0 Then
''
''                bConfirmarHabilitacion = True
''                Call onCajaGeneral.actualizarSeguimientoHabiltacionAgenciaCajaGeneral(lsMovNro, 1, lsMovNro2, bConfirmarHabilitacion)
''
''
''                lsMovNro3 = oNContFunciones.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
''
''
''                rsBill.MoveFirst
''                If nMoneda = gMonedaNacional Then
''                    rsMon.MoveFirst
''                End If
''
''                'If onCajaGeneral.GrabaMovEfectivo(lsMovNro3, fsOpeCodDeposito, txtMovDesc,
''                '                                   rsBill, rsMon, Me.txtIF + lsCtafiltro,
''                '                                   fsCtaHaberDep, txtMonto, 1, txtCtaIF,
''                '                                   val(Right(cboDocumento, 2)), txtNroDoc,
''                '                                   gdFecSis, , lnMovNroConHabEfe) = 0 Then
''
''                'Paso3:Registra el Depósito de Efectivo de la Habilitación
''                'de Caja General
''                Call onCajaGeneral.GrabaMovEfectivo(lsMovNro3, fsOpeCodDeposito, txtMovDesc, _
''                                                    rsBill, rsMon, Me.txtIF + lsCtafiltro, _
''                                                    fsCtaHaberDep, txtMonto, 1, txtCtaIF, _
''                                                    val(Right(cboDocumento, 2)), txtNroDoc, _
''                                                    gdFecSis, , lnMovNroConHabEfe, lnMovNroDepEfe)
''                If lnMovNroDepEfe > 0 Then
''                    bConfirmacionDeposito = True
''                    Call onCajaGeneral.actualizarSeguimientoHabiltacionAgenciaCajaGeneral(lsMovNro, 2, , , lsMovNro3, bConfirmacionDeposito)
''                Else
''                    Set oDCOMMov = New DCOMMov
''                    lsMovNroExt = oNContFunciones.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
''                    Call oDCOMMov.BeginTrans
''                    'Extorna el Depósito de Efectivo de la Habilitación
''                    'de la Agencia a Caja General en caso que no se
''                    'registre correctamente el paso 3
''                    Call oDCOMMov.EliminaMov(lsMovNro3, lsMovNroExt)
''                    'Extorna la Confirmación de la Habilitación
''                    'de la Agencia a Caja General en caso que no se
''                    'registre correctamente el paso 2
''                    Call oDCOMMov.EliminaMov(lsMovNro2, lsMovNroExt)
''                    'Extorna de la Habilitación de la Agencia
''                    'a Caja General en caso que no se
''                    'registre correctamente el paso 1
''                    Call oDCOMMov.EliminaMov(lsMovNro, lsMovNroExt)
''                    Call oDCOMMov.CommitTrans
''                    Set oDCOMMov = Nothing
''                    MsgBox "Paso 3: No se pudo registrar la Habilitación a Caja General." & Chr(10) & "Por favor vuelva a ingresar los datos", vbInformation, "Aviso"
''
''                    Unload Me
''                    Exit Sub
''                End If
''            Else
''                Set oDCOMMov = New DCOMMov
''                lsMovNroExt = oNContFunciones.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
''                Call oDCOMMov.BeginTrans
''                'Extorna la Confirmación de la Habilitación
''                'de la Agencia a Caja General en caso que no se
''                'registre correctamente el paso 2
''                Call oDCOMMov.EliminaMov(lsMovNro2, lsMovNroExt)
''                'Extorna de la Habilitación de la Agencia
''                'a Caja General en caso que no se
''                'registre correctamente el paso 1
''                Call oDCOMMov.EliminaMov(lsMovNro, lsMovNroExt)
''                Call oDCOMMov.CommitTrans
''                Set oDCOMMov = Nothing
''                MsgBox "Paso 2: No se pudo registrar la Habilitación a Caja General." & Chr(10) & "Por favor vuelva a ingresar los datos", vbInformation, "Aviso"
''                Unload Me
''                Exit Sub
''            End If
''        '***Comentado por ELRO el 20130604, según SATI INC1305280013 y INC1305280007****
''        'ElseIf fsOpeCodDeposito <> "" And lnMovNroHabEfe = 0 Then
''        '    Set oDCOMMov = New DCOMMov
''        '    lsMovNroExt = oNContFunciones.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
''        '    Call oDCOMMov.BeginTrans
''        '    'Extorna de la Habilitación de la Agencia
''        '    'a Caja General en caso que no se
''        '    'registre correctamente el paso 1
''        '    Call oDCOMMov.EliminaMov(lsMovNro, lsMovNroExt)
''        '    Call oDCOMMov.CommitTrans
''        '    Set oDCOMMov = Nothing
''        '    MsgBox "Paso 1: No se pudo registrar la Habilitación a Caja General." & Chr(10) & "Por favor vuelva a ingresar los datos", vbInformation, "Aviso"
''        '    Unload Me
''        '    Exit Sub
''        '***Fin Comentado por ELRO el 20130604, según SATI INC1305280013 y INC1305280007
''        End If
''        '***Fin Modificado por ELRO**********************************
''        Dim oContImp As COMNContabilidad.NCOMContImprimir  'NContImprimir
''        Select Case gsOpeCod
''            Case gOpeBoveAgeHabAgeACG
''                Dim lbOk As Boolean
''                Set oContImp = New COMNContabilidad.NCOMContImprimir
''                lbOk = True
''                lsCadImp = oContImp.ImprimeBoletahabilitacion(lblTitulo.Caption, "HABILITACION A CAJA GENERAL", _
''                               txtAreaAgeOrig, lblAreaAgeOrig.Caption, TxtAreaAgeDest, lblAreaAgeDest.Caption, nMoneda, gsOpeCod, _
''                               CDbl(txtMonto.Text), gsNomAge, lsMovNro, sLpt, gsCodCMAC, gbImpTMU)
''                Do While lbOk
''                    nFicSal = FreeFile
''                    Open sLpt For Output As nFicSal
''                        Print #nFicSal, lsCadImp & Chr$(12)
''                        Print #nFicSal, ""
''                    Close #nFicSal
''                    If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
''                        lbOk = False
''                    End If
''                Loop
''                Set oContImp = Nothing
''                '***Modificado por ELRO el 20111024, según Acta 277-2011/TI-D
''                '***Modificado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''                'If bConfirmarHabilitacion Then
''                If lnMovNroConHabEfe > 0 Then
''                '***Fin Modificado por ELRO**********************************
''                    Call ImprimeAsientoContable(lsMovNro2, , , , True, True, txtMovDesc, , , , , , 1)
''                End If
''                '***Modificado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''                'If bConfirmacionDeposito Then
''                If lnMovNroDepEfe > 0 Then
''                '***Fin Modificado por ELRO**********************************
''                    Call ImprimeAsientoContable(lsMovNro3, , , , True, False, , , , , , , 1)
''                End If
''                '***Comentado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''                'If bConfirmarHabilitacion = False Or bConfirmacionDeposito = False Then
''                '    MsgBox "El proceso de confirmación y/o depósito del efectivo no se realizo correctamente. Solicitar a Tesorería que Extorne el Deposito de Efectivo y luego Extorne la Confirmación de la Habilitación de la Agencia a Caja General. Por último Ud. Extorne la Habilitación de Agencia a Oficina Principal", vbExclamation, "Aviso"
''                'End If
''                '***Fin Comentado por ELRO*************************************
''                '***Fin Modificado por ELRO**********************************
''            Case Else
''                Dim lsTexto As String
''                Set oContImp = New COMNContabilidad.NCOMContImprimir
''                Set cPrevio = New previo.clsprevio
''                    lsTexto = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, lsMovNro, gsNomCmac, 2)
''                    cPrevio.Show lsTexto, Me.Caption, False, , gImpresora
''                Set cPrevio = Nothing
''                Set oContImp = Nothing
''        End Select
''
''        Set oCaja = Nothing
''        If MsgBox("Desea Registrar otra Habilitación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
''            '***Modificado por ELRO el 20111114, según Acta 277-2011/TI-D
''            'txtMovDesc = ""
''            'txtMonto = "0.00"
''            fsOpeCodConfirmacion = ""
''            fsOpeCodDeposito = ""
''            fsCtaDebeConHab = ""
''            fsCtaHaberConHab = ""
''            fsCtaHaberDep = ""
''            'lnMovNro = 0 '***Comentado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''            txtMonto = ""
''            txtNroDoc = ""
''            txtMovDesc = ""
''            lblCtaIF = ""
''            txtCtaIF = ""
''            lblIF = ""
''            txtIF = ""
''            lsMovNro = ""
''            lsMovNro2 = ""
''            lsMovNro3 = ""
''            '***Modificado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''            lnMovNroHabEfe = 0
''            lnMovNroConHabEfe = 0
''            lnMovNroDepEfe = 0
''            lsMovNroExt = ""
''            '***Fin Modificado por ELRO*************************************
''            Me.FraDocumento.Visible = False
''            Set onCajaGeneral = Nothing
''            Set oDMov = Nothing
''            Set oNContFunciones = Nothing
''            Set rsBill = Nothing
''            Set rsMon = Nothing
''            optMoneda.iTem(0).value = True
''            Call optMoneda_Click(IIf(optMoneda.iTem(0).value, 0, 1))
''           '***Fin Modificado por ELRO**********************************
''            If txtFechaMov.Enabled Then txtFechaMov.SetFocus
''        Else
''            '***Agregado por ELRO el 20111114, según Acta 277-2011/TI-D
''            fsOpeCodConfirmacion = ""
''            fsOpeCodDeposito = ""
''            fsCtaDebeConHab = ""
''            fsCtaHaberConHab = ""
''            fsCtaHaberDep = ""
''            'lnMovNro = 0 '***Comentado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''            txtMonto = ""
''            txtNroDoc = ""
''            txtMovDesc = ""
''            lblCtaIF = ""
''            txtCtaIF = ""
''            lblIF = ""
''            txtIF = ""
''            lsMovNro = ""
''            lsMovNro2 = ""
''            lsMovNro3 = ""
''            '***Modificado por ELRO el 20120523, según Acta N° 106-2012/TI-D
''            lnMovNroHabEfe = 0
''            lnMovNroConHabEfe = 0
''            lnMovNroDepEfe = 0
''            lsMovNroExt = ""
''            '***Fin Modificado por ELRO*************************************
''            Me.FraDocumento.Visible = False
''            Set onCajaGeneral = Nothing
''            Set oDMov = Nothing
''            Set oNContFunciones = Nothing
''            Set rsBill = Nothing
''            Set rsMon = Nothing
''            optMoneda.iTem(0).value = True
''            Call optMoneda_Click(IIf(optMoneda.iTem(0).value, 0, 1))
''            '***Fin Agregado por ELRO**********************************
''            Unload Me
''        End If
''    Else
''        MsgBox "Paso 1: No se pudo registrar la Habilitación a Caja General." & Chr(10) & "Por favor vuelva a ingresar los datos", vbInformation, "Aviso"
''        Unload Me
''    End If
''End If
''Set oCon = Nothing
''End Sub
'Private Sub CmdAceptar_Click()
'    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
'    Dim oNContFunciones As clases.NContFunciones
'    Dim oContImp As COMNContabilidad.NCOMContImprimir
'    Dim cPrevio As previo.clsprevio
'    Dim oDMov As COMDMov.DCOMMov
'
'    Dim rsBill As ADODB.Recordset
'    Dim rsMon As ADODB.Recordset
'    Dim lsMovNro As String
'    Dim lsCadImp As String
'    Dim nFicSal As Integer
'    Dim bConfirmarHabilitacion As Boolean, bConfirmacionDeposito As Boolean
'    Dim lsMovNro2 As String, lsMovNro3 As String
'    Dim lsCtafiltro As String
'    Dim lnMovNroHabEfe As Long
'    Dim lnMovNroConHabEfe As Long
'    Dim lnMovNroDepEfe As Long
'    Dim lsMovNroExt As String
'    Dim lbok As Boolean
'    Dim bTransaccion As Boolean
'
'    On Error GoTo ErrCmdAceptar
'
'    If valida = False Then Exit Sub
'
'    Select Case gsOpeCod
'        Case gOpeBoveAgeHabEntreAge
'            If TxtAreaAgeDest = txtAreaAgeOrig Then
'                MsgBox "Agencia de Destino no puede ser la misma que de Origen", vbInformation, "Aviso"
'                If TxtAreaAgeDest.Enabled Then TxtAreaAgeDest.SetFocus
'                Exit Sub
'            End If
'        Case Else
'    End Select
'
'    Set oNContFunciones = New NContFunciones
'    Set rsBill = New ADODB.Recordset
'    Set rsMon = New ADODB.Recordset
'    frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, CDbl(txtMonto), nmoneda, False
'    If frmCajaGenEfectivo.lbok Then
'         Set rsBill = frmCajaGenEfectivo.rsBilletes
'         Set rsMon = frmCajaGenEfectivo.rsMonedas
'    Else
'        MsgBox "Debe registrar correctamente la descomposición de efectivo de la habilitación", vbInformation, "Aviso"
'        Set frmCajaGenEfectivo = Nothing
'        Exit Sub
'    End If
'    Set frmCajaGenEfectivo = Nothing
'    If (rsBill Is Nothing And rsMon Is Nothing) Then
'        MsgBox "Error en Ingreso de Billetaje", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    lsCtafiltro = oNContFunciones.GetFiltroObjetos(1, Me.txtIF, Me.txtCtaIF, False)
'    If lsCtafiltro = "" And fsOpeCodDeposito <> "" Then
'        MsgBox "Esta cuenta contable " & Me.txtIF & " no esta registrado en CtaIFFiltro, comunicarse con TI", vbInformation, "Aviso"
'        Exit Sub
'    ElseIf fsOpeCodDeposito = "" Then
'        lsCtafiltro = ""
'    End If
'
'    If oNContFunciones.verificarUltimoNivelCta(Me.txtIF & lsCtafiltro) = False And fsOpeCodDeposito <> "" Then
'       MsgBox "La cuenta contable " & Me.txtIF + lsCtafiltro & " no es de Ultimo Nivel, comunicarse con TI", vbInformation, "Aviso"
'       Exit Sub
'    ElseIf fsOpeCodDeposito = "" Then
'        Me.txtIF = ""
'    End If
'
'    If MsgBox("Desea Realizar la Habilitación respectiva?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
'
'    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
'    Set oDMov = New COMDMov.DCOMMov
'
'    oDMov.BeginTrans
'    bTransaccion = True
'
'    lsMovNro = oDMov.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
'
'    'Paso 1:Registra la habilitación del Efectivo
'    If oCaja.GrabaHabEfectivo(lsMovNro, rsBill, rsMon, gsOpeCod, txtMovDesc, _
'                              CDbl(txtMonto), Trim(txtAreaAgeOrig), _
'                              Trim(TxtAreaAgeDest), "", Format$(txtFechaMov, _
'                              gsFormatoFecha), "", Format$(txtFechaMov, _
'                              gsFormatoFecha), "", 0, nmoneda, , lsAreCodCaja, , _
'                              fsOpeCodDeposito, lnMovNroHabEfe, oDMov) = 0 Then
'
'        If fsOpeCodDeposito <> "" And lnMovNroHabEfe > 0 Then
'            oDMov.registrarSeguimientoHabiltacionAgenciaCajaGeneral lsMovNro
'            lsMovNro2 = oDMov.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
'            rsBill.MoveFirst
'            If nmoneda = gMonedaNacional Then
'                rsMon.MoveFirst
'            End If
'
'            bConfirmarHabilitacion = False
'            bConfirmacionDeposito = False
'
'            'Paso2:Registra la confirmacion de la Habilitación de Agencia a Caja General
'            Call oCaja.GrabaConfHabEfectivo_NEW(lsMovNro2, rsBill, rsMon, _
'                                                    fsOpeCodConfirmacion, txtMovDesc, _
'                                                    fsCtaDebeConHab, fsCtaHaberConHab, _
'                                                    CDbl(txtMonto), Trim(txtAreaAgeOrig), _
'                                                    Trim(TxtAreaAgeDest), lnMovNroHabEfe, _
'                                                    fsOpeCodDeposito, lnMovNroConHabEfe, oDMov)
'            If lnMovNroConHabEfe > 0 Then
'                bConfirmarHabilitacion = True
'                Call oDMov.actualizarSeguimientoHabiltacionAgenciaCajaGeneral(lsMovNro, 1, lsMovNro2, bConfirmarHabilitacion)
'
'                lsMovNro3 = oDMov.GeneraMovNro(CDate(txtFechaMov), gsCodAge, gsCodUser)
'
'                rsBill.MoveFirst
'                If nmoneda = gMonedaNacional Then
'                    rsMon.MoveFirst
'                End If
'
'                'Paso3:Registra el Depósito de Efectivo de la Habilitación de Caja General
'                Call oCaja.GrabaMovEfectivo_NEW(lsMovNro3, fsOpeCodDeposito, txtMovDesc, _
'                                                    rsBill, rsMon, Me.txtIF + lsCtafiltro, _
'                                                    fsCtaHaberDep, txtMonto, 1, txtCtaIF, _
'                                                    val(Right(cboDocumento, 2)), txtNroDoc, _
'                                                    gdFecSis, , lnMovNroConHabEfe, lnMovNroDepEfe, oDMov)
'                If lnMovNroDepEfe > 0 Then
'                    bConfirmacionDeposito = True
'                    Call oDMov.actualizarSeguimientoHabiltacionAgenciaCajaGeneral(lsMovNro, 2, , , lsMovNro3, bConfirmacionDeposito)
'                End If
'            End If
'        End If
'    Else
'        MsgBox "Paso 1: No se pudo registrar la Habilitación a Caja General." & Chr(10) & "Por favor vuelva a ingresar los datos", vbInformation, "Aviso"
'        Unload Me
'    End If
'
'    oDMov.CommitTrans
'    bTransaccion = False
'    Set oContImp = New COMNContabilidad.NCOMContImprimir
'
'    Select Case gsOpeCod
'        Case gOpeBoveAgeHabAgeACG
'            lbok = True
'            lsCadImp = oContImp.ImprimeBoletahabilitacion(lblTitulo.Caption, "HABILITACION A CAJA GENERAL", _
'                           txtAreaAgeOrig, lblAreaAgeOrig.Caption, TxtAreaAgeDest, lblAreaAgeDest.Caption, nmoneda, gsOpeCod, _
'                           CDbl(txtMonto.Text), gsNomAge, lsMovNro, sLpt, gsCodCMAC, gbImpTMU)
'            Do While lbok
'                nFicSal = FreeFile
'                Open sLpt For Output As nFicSal
'                    Print #nFicSal, lsCadImp & Chr$(12)
'                    Print #nFicSal, ""
'                Close #nFicSal
'                If MsgBox("Desea Reimprimir Boleta de Operacion??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'                    lbok = False
'                End If
'            Loop
'
'            If lnMovNroConHabEfe > 0 Then
'                Call ImprimeAsientoContable(lsMovNro2, , , , True, True, txtMovDesc, , , , , , 1)
'            End If
'            If lnMovNroDepEfe > 0 Then
'                Call ImprimeAsientoContable(lsMovNro3, , , , True, False, , , , , , , 1)
'            End If
'        Case Else
'            Dim lsTexto As String
'            Set cPrevio = New previo.clsprevio
'                lsTexto = oContImp.ImprimeDocSalidaEfectivo(gnColPage, gdFecSis, gsOpeCod, lsMovNro, gsNomCmac, 2)
'                cPrevio.Show lsTexto, Me.Caption, False, , gImpresora
'    End Select
'
'    Set oCaja = Nothing
'    Set oNContFunciones = Nothing
'    Set oContImp = Nothing
'    Set cPrevio = Nothing
'    Set oDMov = Nothing
'
'    fsOpeCodConfirmacion = ""
'    fsOpeCodDeposito = ""
'    fsCtaDebeConHab = ""
'    fsCtaHaberConHab = ""
'    fsCtaHaberDep = ""
'    txtMonto = ""
'    txtNroDoc = ""
'    txtMovDesc = ""
'    lblCtaIF = ""
'    txtCtaIF = ""
'    lblIF = ""
'    txtIF = ""
'    lsMovNro = ""
'    lsMovNro2 = ""
'    lsMovNro3 = ""
'    lnMovNroHabEfe = 0
'    lnMovNroConHabEfe = 0
'    lnMovNroDepEfe = 0
'    lsMovNroExt = ""
'    Me.FraDocumento.Visible = False
'    optMoneda.Item(0).value = True
'    Call OptMoneda_Click(IIf(optMoneda.Item(0).value, 0, 1))
'
'    If MsgBox("Desea Registrar otra Habilitación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'        If txtFechaMov.Enabled Then txtFechaMov.SetFocus
'    Else
'        Unload Me
'    End If
'    Exit Sub
'ErrCmdAceptar:
'    If bTransaccion Then
'        oDMov.RollbackTrans
'        Set oDMov = Nothing
'    End If
'    MsgBox err.Description, vbCritical, "Aviso"
'End Sub
''END EJVG *******
'Function valida() As Boolean
'valida = True
'If txtAreaAgeOrig = "" Then
'    MsgBox "Area o Agencia de Origen no Seleccionada", vbInformation, "Aviso"
'    If txtAreaAgeOrig.Enabled Then txtAreaAgeOrig.SetFocus
'    valida = False
'    Exit Function
'End If
'If Len(Trim(txtMovDesc)) = 0 Then
'    MsgBox "Descripción del Movimiento no Ingresado", vbInformation, "Aviso"
'    txtMovDesc.SetFocus
'    valida = False
'    Exit Function
'End If
'If val(txtMonto) = 0 Then
'    MsgBox "Ingrese Monto de Operación", vbInformation, "Aviso"
'    txtMonto.SetFocus
'    valida = False
'    Exit Function
'End If
''***Agregado por ELRO el 20111209, según Acta 277-2011/TI-D
'If FraDocumento.Visible Then
'    If Len(Trim(cboDocumento)) <> 0 Then
'        If Len(Trim(txtNroDoc.Text)) = 0 Then
'            MsgBox "Nro de Documento no Ingresado", vbInformation, "Aviso"
'            If txtNroDoc.Enabled And txtNroDoc.Visible Then
'                txtNroDoc.SetFocus
'            End If
'            valida = False
'            Exit Function
'        End If
'    Else
'        If MsgBox("Documento no ha sido ingresado. Desea continuar??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'            valida = False
'            cboDocumento.SetFocus
'            Exit Function
'        End If
'    End If
'End If
'
'If Len(Trim(txtIF)) = 0 Then
'    MsgBox "Debe Elegir la Intirución Financiera", vbInformation, "Aviso"
'    txtIF.SetFocus
'    valida = False
'    Exit Function
'End If
'
'If Len(Trim(txtCtaIF)) = 0 Then
'    MsgBox "Debe Elegir la Cuenta de Intirución Financiera", vbInformation, "Aviso"
'    txtCtaIF.SetFocus
'    valida = False
'    Exit Function
'End If
''***Fin Agregado por ELRO**********************************
'End Function
'
'Private Sub cmdsalir_Click()
'Unload Me
'End Sub
'
'Private Sub Form_Activate()
'If lbSalir Then
'    Unload Me
'End If
'End Sub
'
'Private Sub Form_Load()
'Dim rs As ADODB.Recordset
'Dim oOpe As COMNCajaGeneral.NCOMCajaGeneral
'Dim oGen As COMDConstSistema.DCOMGeneral  'DGeneral
'Set oGen = New COMDConstSistema.DCOMGeneral
'
'lbSalir = False
'txtFechaMov = gdFecSis
'Me.Caption = gsOpeDesc
'Me.Icon = LoadPicture(App.Path & gsRutaIcono)
'Set oOpe = New COMNCajaGeneral.NCOMCajaGeneral
'Set rs = New ADODB.Recordset
'
'Select Case gsOpeCod
'    Case gOpeBoveCGHabAge
'        lblTitulo = "HABILITACION A AGENCIAS"
'        txtAreaAgeOrig.Text = gsCodAge
'        lsAreCodCaja = ""
'    Case gOpeBoveAgeHabAgeACG
'        lblTitulo = "HABILITACION A CAJA GENERAL"
'        lblAreaAgeOrig = gsNomAge
'        Set rs = oOpe.GetOpeObj(gsOpeCod, "0")
'        lsAreCodCaja = rs("Codigo")
'        sAreaNomCaja = rs("Descripcion")
'        Set rs = oOpe.GetOpeObj(gsOpeCod, "1")
'        txtAreaAgeOrig.Text = rs("Codigo") & gsCodAge
'        rs.Close
'        Set rs = Nothing
'
'        lblAreaAgeOrig = gsNomAge
'        TxtAreaAgeDest.Text = lsAreCodCaja
'        lblAreaAgeDest = sAreaNomCaja
'        txtAreaAgeOrig.Enabled = False
'        TxtAreaAgeDest.Enabled = False
'        txtFechaMov.Enabled = False
'        '***Modificado por ELRO el 20111021, según Acta 277-2011/TI-D
'        Me.FraIF.Visible = True
'        '***Fin Modificado por ELRO**********************************
'    Case gOpeBoveAgeHabEntreAge
'        lblTitulo = "HABILITACION ENTRE AGENCIAS"
'        lsAreCodCaja = ""
'
'        lblAreaAgeOrig = gsNomAge
'        'Set rs = oOpe.GetOpeObj(gsOpeCod, "0")
'        'lsAreCodCaja = rs("Codigo")
'        'sAreaNomCaja = rs("Descripcion")
''        Set rs = oOpe.GetOpeObj(gsOpeCod, "1")
''        txtAreaAgeOrig.Text = rs("Codigo") & gsCodAge
''        rs.Close
''        Set rs = Nothing
''        lblAreaAgeOrig = gsNomAge
''        txtAreaAgeOrig.Enabled = False
''        txtFechaMov.Enabled = False
''
''        Set rs = oOpe.GetOpeObj(gsOpeCod, "1")
''        TxtAreaAgeDest.rs = rs
'
'
'End Select
'txtMonto.BackColor = IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, vbWhite, &HC0FFC0)
'Set oGen = Nothing
'Me.Caption = gsOpeCod & " - " & lblTitulo
'optMoneda(0).value = True
'End Sub
'
'Private Sub OptMoneda_Click(Index As Integer)
'Select Case Index
'Dim oDOperacion As clases.DOperacion
'    Case 0
'
'        Set oDOperacion = New clases.DOperacion
'
'        txtMonto.BackColor = &H80000005
'        nmoneda = gMonedaNacional
'        '***Modificado por ELRO el 20111025, según Acta 277-2011/TI-D
'        Me.txtIF = ""
'        Me.lblIF = ""
'        fsOpeCodConfirmacion = gOpeBoveCGConfHabAgeBove
'        fsCtaDebeConHab = "11110101"
'        fsCtaHaberConHab = "111109"
'        fsCtaHaberDep = "11110101"
'
'        'Call CargarCtaConAgencia
'        Me.txtIF.psRaiz = "Institución Financiera"
'        Me.txtIF.rs = oDOperacion.listarOperacionCtaCta("D", IIf(optMoneda(0).value, "1", "2"))
'        Set oDOperacion = Nothing
'        '***Fin Modificado por ELRO**********************************
'    Case 1
'
'        Set oDOperacion = New clases.DOperacion
'
'        txtMonto.BackColor = &HC0FFC0
'        nmoneda = gMonedaExtranjera
'        '***Modificado por ELRO el 20111025, según Acta 277-2011/TI-D
'        Me.txtIF = ""
'        Me.lblIF = ""
'        fsOpeCodConfirmacion = "402402"
'        fsCtaDebeConHab = "11210101"
'        fsCtaHaberConHab = "112109"
'        fsCtaHaberDep = "11210101"
'        'Call CargarCtaConAgencia
'        Me.txtIF.psRaiz = "Institución Financiera"
'        Me.txtIF.rs = oDOperacion.listarOperacionCtaCta("D", IIf(optMoneda(0).value, "1", "2"))
'        Set oDOperacion = Nothing
'        '***Fin Modificado por ELRO**********************************
'End Select
'End Sub
'
'Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If txtMovDesc.Enabled Then txtMovDesc.SetFocus
'End If
'End Sub
'
'Private Sub TxtAreaAgeDest_EmiteDatos()
'lblAreaAgeDest = TxtAreaAgeDest.psDescripcion
'If lblAreaAgeDest <> "" Then
'    If optMoneda(0).Visible Then optMoneda(0).SetFocus
'End If
'End Sub
'
'Private Sub txtAreaAgeOrig_EmiteDatos()
'lblAreaAgeOrig = txtAreaAgeOrig.psDescripcion
'End Sub
'
'Private Sub txtFechaMov_KeyPress(KeyAscii As Integer)
'KeyAscii = fgIntfMayusculas(KeyAscii)
'If KeyAscii = 13 Then
'    If TxtAreaAgeDest.Enabled Then
'        TxtAreaAgeDest.SetFocus
'    ElseIf Me.txtMovDesc.Enabled Then
'        txtMovDesc.SetFocus
'    ElseIf Me.txtMonto.Enabled Then
'        txtMonto.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub txtMonto_GotFocus()
'fEnfoque txtMonto
'End Sub
'
'Private Sub txtMonto_KeyPress(KeyAscii As Integer)
'KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 20, 2)
'If KeyAscii = 13 Then
'    cmdAceptar.SetFocus
'End If
'End Sub
'
'Private Sub txtMonto_LostFocus()
' If txtMonto = "" Then txtMonto = 0
'txtMonto = Format(txtMonto, "#,#0.00")
'End Sub
'
''***Agregado por ELRO el 20111022, según Acta 277-2011/TI-D
'Private Sub txtIF_EmiteDatos()
'    Dim oDOperacion As clases.DOperacion
'    Set oDOperacion = New clases.DOperacion
'    Dim sComparar As String
'
'    Dim rsCuentas As ADODB.Recordset
'    Set rsCuentas = New ADODB.Recordset
'
'
'    Me.txtCtaIF = ""
'    Me.lblCtaIF = ""
'    Me.lblIF = Me.txtIF.psDescripcion
'
'    If Me.txtIF <> "" Then
'
'        Set rsCuentas = Me.txtIF.rs
'
'        Me.txtCtaIF.psRaiz = "Cuentas de Instituciones Financieras"
'        If Me.txtIF = "111301" Or Me.txtIF = "112301" Then
'            sComparar = "_1_[12]" & IIf(optMoneda(0).value, "1", "2") & "%"
'            Me.txtCtaIF.rs = oDOperacion.listarCuentasEntidadesFinacieras(sComparar, IIf(optMoneda(0).value, "1", "2"))
'
'
'        Else
'            sComparar = "_3_2%"
'            Me.txtCtaIF.rs = oDOperacion.listarCuentasEntidadesFinacieras(sComparar, IIf(optMoneda(0).value, "1", "2"))
'
'        End If
'
'        rsCuentas.MoveFirst
'
'        Do While Not rsCuentas.EOF
'            If rsCuentas!cCtaContCod = Me.txtIF Then
'                fsOpeCodDeposito = rsCuentas!cOpecod
'            End If
'               rsCuentas.MoveNext
'        Loop
'
'    End If
'
'    Set oDOperacion = Nothing
'    Set rsCuentas = Nothing
'
'
'End Sub
'
'Private Sub txtCtaIF_EmiteDatos()
'
'    Dim lsNombreBanco As String
'    Dim lsCuenta As String
'
'    Dim oNCajaCtaIF As clases.NCajaCtaIF
'    Set oNCajaCtaIF = New clases.NCajaCtaIF
'    Dim oDOperacion As clases.DOperacion
'    Set oDOperacion = New clases.DOperacion
'    Dim rsDocumento As ADODB.Recordset
'    Set rsDocumento = New ADODB.Recordset
'
'    If Me.txtCtaIF <> "" Then
'        Me.lblCtaIF = Me.txtCtaIF.psDescripcion
'        lsNombreBanco = oNCajaCtaIF.NombreIF(Mid(Me.txtCtaIF, 4, 13))
'        lsCuenta = oDOperacion.recuperaTipoCuentaEntidadFinaciera(Mid(Me.txtCtaIF, 18, 10)) & " " & Me.txtCtaIF.psDescripcion
'        Me.lblCtaIF = lsNombreBanco & " " & lsCuenta
'
'        Me.FraDocumento.Visible = True
'        txtNroDoc = ""
'        Set rsDocumento = oDOperacion.CargaOpeDoc(fsOpeCodDeposito)
'
'        cboDocumento.Clear
'        Do While Not rsDocumento.EOF
'            cboDocumento.AddItem Mid(rsDocumento!cDocDesc & Space(100), 1, 100) & rsDocumento!nDocTpo
'            rsDocumento.MoveNext
'        Loop
'        cboDocumento.ListIndex = -1
'    End If
'
'
'
'    Set rsDocumento = Nothing
'    Set oNCajaCtaIF = Nothing
'    Set oDOperacion = Nothing
'
'End Sub
'
'Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
'    KeyAscii = NumerosEnteros(KeyAscii)
'    If KeyAscii = 13 Then
'        txtMonto.SetFocus
'    End If
'End Sub
''***Fin Agregado por ELRO**********************************



