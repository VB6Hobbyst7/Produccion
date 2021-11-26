VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroServ 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5025
   ClientLeft      =   3045
   ClientTop       =   1560
   ClientWidth     =   4470
   HelpContextID   =   400
   Icon            =   "frmOtrasOpe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame fraUsuario 
      Height          =   2175
      Left            =   660
      TabIndex        =   16
      Top             =   2730
      Width           =   4095
      Begin VB.Label lblPersIde 
         Caption         =   "Cod. Persona"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPersIdeT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblDocIdeT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1830
         Width           =   2295
      End
      Begin VB.Label lblFecNac 
         Caption         =   "Fecha Nacimiento"
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblFecNacT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label lblDocIde 
         Caption         =   "Doc. Identidad"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lbldireccionT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1250
         Width           =   3855
      End
      Begin VB.Label lbldireccion 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1000
         Width           =   3855
      End
      Begin VB.Label lblNombreT 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Frame FraServicios 
      Height          =   1395
      Left            =   225
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
      Begin MSMask.MaskEdBox txtReferencia 
         Height          =   300
         Left            =   1650
         TabIndex        =   29
         Top             =   615
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtNroRecibo 
         Height          =   315
         Left            =   1650
         TabIndex        =   27
         Top             =   255
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin Sicmact.EditMoney EditMoney2 
         Height          =   405
         Left            =   1650
         TabIndex        =   2
         Top             =   945
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label lblReferencia 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   675
         Width           =   945
      End
      Begin VB.Label lblSimbolo1 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3435
         TabIndex        =   13
         Top             =   1050
         Width           =   285
      End
      Begin VB.Label lblRecibo 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Recibo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   315
         Width           =   1305
      End
      Begin VB.Label lblmontorecibo 
         AutoSize        =   -1  'True
         Caption         =   "Neto a Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   1050
         Width           =   1140
      End
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Height          =   400
      Left            =   2970
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   400
      Left            =   1770
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraTipoCambio 
      Height          =   1095
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   3945
      Begin Sicmact.EditMoney EditMoney1 
         Height          =   405
         Left            =   1785
         TabIndex        =   0
         Top             =   210
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.TextBox TxtMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   675
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "$."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   315
         Width           =   180
      End
      Begin VB.Label lblsimbolosoles2 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   14
         Top             =   743
         Width           =   285
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto a Pagar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   713
         Width           =   1575
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Monto a Cambiar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.Label lblTipcambiaDia 
      Caption         =   "Tipo de Cambio del Día"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1605
      Width           =   3855
   End
End
Attribute VB_Name = "frmCajeroServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'Aplicacion: AHORROS
'Resumen: realiza otras operaciones en caja, tales como compra
'y venta de moneda extranjera, y pago de servicios a
'sedalib e hidrandina, solo las dos primeras operaciones
'anteriores generan recibos.
'****************************************************
'Archivo: frmOtrasOpe
'Fecha de Creación: 19/07/1999
'Ultima Modificación: 23/07/1999
'****************************************************
'OBJETOS
'Modulos Usados:
'   FunGeneral
'****************************************************
'Controles de Objetos Usados: Ninguno
'****************************************************
'Variables Globales
'Del Formulario:
'   Ninguna
'Del Sistema:
'   Ninguna
'****************************************************
'Procedimientos y funciones
'-------------------------------------------------------
'CompraMonedaExtranjera:Procedimiento que realiza la compra
'de moneda extranjera, solo se le indica el monto que el
'cliente va a comprar de la moneda extranjera
'-------------------------------------------------------
'DameValorCambio:Función que devuelve el valor de cambio
'del dia de acuerdo a la operacion que se esta realizando
'esto quiere decir que para operaciones de compra y venta
'de el tipo de cambio del dia de la compra o de la venta
'respectivamente
'-------------------------------------------------------
'PagoServicio:Procedimiento que realiza el paga de servicio,
'se le indica el servicio que va ha realizar, el monto y el
'numero del recibo que se va ha cancelar
'-------------------------------------------------------
'VentaMonedaExtranjera:Procedimiento que realiza la venta
'de moneda extranjera, solo se le indica el monto que el
'cliente va a vender a la caje municipal de moneda extranjera
'-------------------------------------------------------

Option Explicit
Dim lnTipCam As Double
Dim lsNumRecibo As String
Dim lsCodigoRec As String
Dim lnMontoRec As Double

Private Sub ClearScreen()
        Me.EditMoney1.Value = 0
        Me.TxtMontoPagar.Text = ""
        Me.lbldireccionT.Caption = ""
        Me.lblDocIdeT.Caption = ""
        Me.lblFecNacT.Caption = ""
        Me.lblNombreT.Caption = ""
        Me.lblPersIdeT.Caption = ""
        'Me.EditMoney1.SetFocus
End Sub

Private Sub ClearScreenServicio()
    Me.EditMoney2.Value = 0
    
    txtNroRecibo.Mask = ""
    txtReferencia.Mask = ""
    txtNroRecibo.Text = ""
    txtReferencia.Text = ""
    
    If frmOperacion.txtCodOpe = gsCobHid Then
        txtNroRecibo.Mask = "###-##-##-########"
        txtReferencia.Mask = "########"
        lblReferencia.Caption = "Num.Medidor"
    Else
        txtNroRecibo.Mask = "###-########-##"
        txtReferencia.Mask = "###########"
        lblReferencia.Caption = "Código"
    End If
    Me.txtNroRecibo.SetFocus
End Sub


Private Sub cmdBuscar_Click()
    frmBuscaCli.Inicia frmOtrasOpe, True, True
    If CodGrid <> "" Then
        Me.lbldireccionT.Caption = DirGrid
        Me.lblDocIdeT.Caption = DNatGrid
        Me.lblFecNacT.Caption = NacGrid
        Me.lblNombreT.Caption = PstaNombre(NomGrid, False)
        Me.lblPersIdeT.Caption = CodGrid
    End If
    Me.CmdGuardar.SetFocus
End Sub

Private Sub cmdGuardar_Click()
    Dim lnSaldo As Currency
    Dim varFecha As String
    Dim bReImp As Boolean
    Dim nNumComp As Long
    On Error GoTo ERROR
    
    gbOpeOk = False
    
    Select Case frmOperacion.txtCodOpe.Text
        Case gsCVComME
            If Me.EditMoney1.Value = 0 Then
                MsgBox "El Monto no puede ser 0 o menor que 0.", vbInformation, "Aviso"
                Me.EditMoney1.SetFocus
                Exit Sub
            ElseIf Trim(Me.TxtMontoPagar.Text) = "" Then
                MsgBox "El Monto a Pagar no puede ser Vacio", vbInformation, "Aviso"
                Exit Sub
            ElseIf Trim(Me.lblPersIdeT.Caption) = "" Then
                MsgBox "Se debe Ingresar una Persona para la compra de Dolares.", vbInformation, "Aviso"
                Me.cmdBuscar.SetFocus
                Exit Sub
            End If
            
            nNumComp = CompraMonedaExtranjera(Str(EditMoney1.Value), lnTipCam, lblPersIdeT.Caption, gsCVComME)
            If gbOpeOk Then
                bReImp = False
                Do
                    Call ImprimeBoletaCVME("COMPRA MON.EXTR.", "Monto Comprado $.", gsCVComME, Me.EditMoney1.Text, "Neto a Pagar S/.", Me.TxtMontoPagar.Text, "Tipo de Cambio Compra", Format(lnTipCam, "#0.000"), Str(nNumComp), Me.lblNombreT.Caption, Me.lbldireccionT.Caption, Me.lblDocIdeT.Caption)
                    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                        bReImp = True
                    Else
                        bReImp = False
                    End If
                Loop Until Not bReImp
                ClearScreen
            End If
        
        Case gsCVVenME
            If Me.EditMoney1.Value = 0 Then
                MsgBox "El Monto no puede ser 0 o menor que 0.", vbInformation, "Aviso"
                Me.EditMoney1.SetFocus
                Exit Sub
            ElseIf Val(TxtMontoPagar.Text) = 0 Then
                MsgBox "El Monto a Pagar no puede ser Vacio", vbInformation, "Aviso"
                Exit Sub
            ElseIf Trim(Me.lblPersIdeT.Caption) = "" Then
                MsgBox "Se debe Ingresar una Persona para la Venta de Dolares.", vbInformation, "Aviso"
                Me.cmdBuscar.SetFocus
                Exit Sub
            End If
            
            nNumComp = VentaMonedaExtranjera(Me.EditMoney1.Value, lnTipCam, lblPersIdeT.Caption, gsCVVenME)
            If gbOpeOk Then
                bReImp = False
                Do
                    Call ImprimeBoletaCVME("VENTA MON.EXTR.", "Monto Vendido $.", gsCVVenME, Me.EditMoney1.Text, "Neto a Recibir S/.", Me.TxtMontoPagar.Text, "Tipo de Cambio Venta", Format(lnTipCam, "#0.000"), Str(nNumComp), Me.lblNombreT.Caption, Me.lbldireccionT.Caption, Me.lblDocIdeT.Caption)
                    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                        bReImp = True
                    Else
                        bReImp = False
                    End If
                Loop Until Not bReImp
                ClearScreen
                                
            End If
        Case Else
            If Me.EditMoney2.Value = 0 Then
                MsgBox "El Monto no puede 0 o menor que 0", vbInformation, "Aviso"
                Me.EditMoney2.SetFocus
                Exit Sub
            ElseIf Len(Me.txtReferencia.Text) = 0 Then
                MsgBox "Debe indicar un el " & Me.lblReferencia.Caption & " del Recibo", vbInformation, "Aviso"
                txtReferencia.SetFocus
                Exit Sub
            ElseIf Not ValidaDocIngresado(txtNroRecibo.Text, txtReferencia.Text) Then
                MsgBox "El Recibo ya fue ingresado el dia de Hoy.", vbInformation, "Aviso"
                txtReferencia.SetFocus
                Exit Sub
            End If
            
            If Trim(Me.txtNroRecibo.Text) = "" Then
                MsgBox "El Recibo no puede ser Vacio", vbInformation, "Aviso"
                Me.txtNroRecibo.SetFocus
                Exit Sub
            End If
            
            If frmOperacion.txtCodOpe <> gsCobHid Then
                If ReadVarSis("AHO", "CtrValCobr") = "1" Then
                    If Not ValSedalib(Left(txtNroRecibo, 3) & Mid(txtNroRecibo.Text, 5, 8) & Right(txtNroRecibo, 2), txtReferencia.Text, EditMoney2.Value) And (lsNumRecibo <> txtNroRecibo Or lsCodigoRec <> txtReferencia.Text Or lnMontoRec <> Me.EditMoney2.Value) Then
                        lsNumRecibo = txtNroRecibo.Text
                        lsCodigoRec = txtReferencia.Text
                        lnMontoRec = Me.EditMoney2.Value
                        MsgBox "Códigos Errados.", vbInformation, "Aviso"
                        ClearScreenServicio
                        Exit Sub
                    End If
                End If
            End If
            
            PagoServicio Me.EditMoney2.Text, frmOperacion.txtCodOpe, Me.txtNroRecibo.Text, Me.txtReferencia.Text
            
            lsNumRecibo = ""
            lsCodigoRec = ""
            lnMontoRec = 0
            
            If gbOpeOk Then
                Do
                    If frmOperacion.txtCodOpe = gsCobHid Then
                        Call ImprimeBoletaCVME("PAGO.SER.HIDRAN.", "Monto Pagado", gsCVComME, Me.EditMoney2.Text, "Recibo:", txtNroRecibo.Text, "", "", "", "", "", "")
                    Else
                        Call ImprimeBoletaCVME("PAGO.SER.SEDALIB", "Monto Pagado", gsCVComME, Me.EditMoney2.Text, "Recibo:", txtNroRecibo.Text, "", "", "", "", "", "")
                    End If
                    
                    If MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                        bReImp = True
                    Else
                        bReImp = False
                    End If
                Loop Until Not bReImp
            End If
            
            ClearScreenServicio
        End Select
        
        Exit Sub
ERROR:
    MsgBox "Existe un error " & Err.Description, vbCritical, Me.Caption

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub EditMoney1_Change()
    TxtMontoPagar.Text = Format(lnTipCam * Me.EditMoney1.Value, "###,##0.00")
End Sub

Private Sub EditMoney1_GotFocus()
    Me.EditMoney1.MarcaTexto
End Sub

Private Sub EditMoney1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdBuscar.SetFocus
End If
End Sub

Private Sub EditMoney2_GotFocus()
    Me.EditMoney2.MarcaTexto
End Sub

Private Sub EditMoney2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.CmdGuardar.SetFocus
End If
End Sub

Private Sub Form_Activate()
    If (lnTipCam <= 0 And frmOperacion.txtCodOpe = gsCVComME) Or (lnTipCam <= 0 And frmOperacion.txtCodOpe = gsCVVenME) Then
        MsgBox "El tipo de Cambio no puede ser 0 o menor que 0, Atualiza el tipo de cambio del día.", vbInformation, "Aviso"
        Unload Me
    Else
        If Me.FraServicios.Visible Then
            txtNroRecibo.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    If Not AbreConexion Then Exit Sub
    
    Me.TxtMontoPagar.Text = ""
    Me.txtNroRecibo.Text = ""
    
    If frmOperacion.txtCodOpe = gsCVComME Or frmOperacion.txtCodOpe = gsCVVenME Then
        If frmOperacion.txtCodOpe = gsCVComME Then
            Me.lblTipoCambio.Caption = "Monto a Comprar"
            Me.lblMonto.Caption = "Neto a Pagar"
            lnTipCam = DameValorCambio(True)
            Me.lblTipcambiaDia.Caption = Me.lblTipcambiaDia.Caption + " Compra : S/. " + Format(lnTipCam, "#0.000")
        Else
            Me.lblTipoCambio.Caption = "Monto a Vender"
            Me.lblMonto.Caption = "Neto a Recibir"
            lnTipCam = DameValorCambio(False)
            Me.lblTipcambiaDia.Caption = Me.lblTipcambiaDia.Caption + " Venta : S/. " + Format(lnTipCam, "#0.000")
            
        End If
        
        Me.Width = 4245
        Me.Height = 4905
        Me.fraTipoCambio.Visible = True
        Me.FraServicios.Visible = False
        Me.cmdBuscar.Enabled = True
        Me.CmdGuardar.Top = 1150
        Me.CmdSalir.Top = 1150
    Else
        Me.lblTipcambiaDia.Visible = False
        Me.fraTipoCambio.Visible = False
        Me.FraServicios.Visible = True
        Me.cmdBuscar.Enabled = False
        Me.Width = 4245
        Me.Height = 2265
        Me.CmdGuardar.Top = 1440
        Me.CmdSalir.Top = 1440

        If frmOperacion.txtCodOpe = gsCobHid Then
            txtNroRecibo.Mask = "###-##-##-########"
            txtReferencia.Mask = "########"
            lblReferencia.Caption = "Num.Medidor"
        Else
            txtNroRecibo.Mask = "###-########-##"
            txtReferencia.Mask = "###########"
            lblReferencia.Caption = "Código"
        End If
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub txtNroRecibo_GotFocus()
    txtNroRecibo.SelStart = 0
    txtNroRecibo.SelLength = 50
End Sub

Private Sub TxtNroRecibo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtReferencia.SetFocus
End If
End Sub

Private Sub txtReferencia_GotFocus()
    txtReferencia.SelStart = 0
    txtReferencia.SelLength = 50
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Right(txtReferencia, 4) = "    " Then txtReferencia = Left(txtReferencia, 4) & "       "
        Me.EditMoney2.SetFocus
    End If
End Sub

Private Function DigChequeo(ByVal psRecibo As String, ByVal psCodigoCliente As String, ByVal pcMonto As Currency) As Integer
    Dim sfacsernro As String, sfacnro As String, sDigCheq As String
    Dim sValor As String
    Dim iLongValor As Integer, iValTotal As Integer, n As Integer
    sfacsernro = Mid(psRecibo, 1, 3)
    sfacnro = Mid(psRecibo, 4, 8)
    sDigCheq = Right(psRecibo, 2)
    sValor = psCodigoCliente & sfacsernro & sfacnro & Right(String((10 - Len(Str(pcMonto * 100))), " ") & Str(pcMonto * 100), 8)
    iLongValor = Len(sValor)
    n = 0
    iValTotal = 0
    Do While n < iLongValor
        n = n + 1
        iValTotal = iValTotal + n * Val(Mid(sValor, n, 1))
    Loop
    DigChequeo = 97 - (iValTotal - Int(iValTotal / 97) * 97)
End Function

Private Function ValSedalib(ByVal psRecibo As String, ByVal psCodigoCliente As String, ByVal pcMonto As Currency) As Boolean
    If DigChequeo(psRecibo, psCodigoCliente, pcMonto) = Val(Right(psRecibo, 2)) Then
        ValSedalib = True
    Else
        ValSedalib = False
    End If
End Function

Private Function ValidaDocIngresado(psNumRecibo As String, psReferencia As String) As Boolean
    Dim sqlO As String
    Dim rsO As ADODB.Recordset
    Set rsO = New ADODB.Recordset
    
    sqlO = " Select TA.dFecTran from OtrasOpe OO" _
         & " Inner Join Trandiaria TA On TA.dFecTran = OO.dFecTran And TA.nNumTran = OO.nNumTran" _
         & " where TA.dFecTran between '" & Format(gdFecSis, "mm/dd/yyyy") & "' And '" & Format(DateAdd("d", 1, gdFecSis), "mm/dd/yyyy") & "' And TA.cCodope in ('" & gsCobHid & "','" & gsCobSed & "') and OO.cGlosa = '" & psNumRecibo & "' and OO.cNumDoc = '" & psReferencia & "' and TA.cFlag is Null"
         rsO.Open sqlO, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If RSVacio(rsO) Then
        ValidaDocIngresado = True
    Else
        ValidaDocIngresado = False
    End If
    
    rsO.Close
    Set rsO = Nothing
End Function
