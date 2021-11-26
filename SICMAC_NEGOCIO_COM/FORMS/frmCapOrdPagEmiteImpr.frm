VERSION 5.00
Begin VB.Form frmCapOrdPagEmiteImpr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emision de Ordenes de Pago"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frmCapOrdPagEmiteImpr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir Ordenes"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7965
      TabIndex        =   17
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   4320
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   4320
      Width           =   990
   End
   Begin VB.Frame fraOrdPag 
      Caption         =   "Orden Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1860
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   10050
      Begin VB.OptionButton optSeleccion 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Top             =   270
         Width           =   1020
      End
      Begin VB.OptionButton optSeleccion 
         Caption         =   "&Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1455
         TabIndex        =   12
         Top             =   285
         Width           =   1020
      End
      Begin SICMACT.FlexEdit grdOrdPag 
         Height          =   1110
         Left            =   105
         TabIndex        =   14
         Top             =   675
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   3016
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Act-Cuenta-Inicio-Fin-Fecha-# Tal-Titular-nTipo-cMovNro"
         EncabezadosAnchos=   "300-500-1700-800-800-1000-600-3700-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Datos Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10140
      Begin VB.Frame fraDetCuenta 
         Height          =   1515
         Left            =   6480
         TabIndex        =   2
         Top             =   660
         Width           =   3495
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   330
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   750
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "# Firmas :"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1170
            Width           =   690
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
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
            Height          =   375
            Left            =   1140
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
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
            Height          =   375
            Left            =   1140
            TabIndex        =   4
            Top             =   660
            Width           =   2175
         End
         Begin VB.Label lblFirmas 
            Alignment       =   2  'Center
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
            Height          =   375
            Left            =   1140
            TabIndex        =   3
            Top             =   1080
            Width           =   2175
         End
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1455
         Left            =   60
         TabIndex        =   1
         Top             =   735
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2566
         Cols0           =   3
         EncabezadosNombres=   "#-Nombre-RE"
         EncabezadosAnchos=   "350-5000-500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0"
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
      End
      Begin VB.Label lblDatosCuenta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3945
         TabIndex        =   10
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmCapOrdPagEmiteImpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

Private Sub cmdCancelar_Click()
CmdGrabar.Enabled = True
fraCuenta.Enabled = True
cmdImprimir.Enabled = False
'fraSolicitud.Enabled = False
txtCuenta.Age = Right(gsCodAge, 2)
txtCuenta.Prod = gCapAhorros
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.Cuenta = ""
txtCuenta.SetFocusCuenta
cmdcancelar.Enabled = True
lblApertura = ""
lblTipoCuenta = ""
lblFirmas = ""
txtCuenta.EnabledCta = True
grdCliente.Rows = 2
grdCliente.Clear
grdCliente.FormaCabecera
optSeleccion(0).Enabled = False
optSeleccion(1).Enabled = False
GridClear
'grdHistoria.Rows = 2
'grdHistoria.Clear
'grdHistoria.FormaCabecera

End Sub

Private Sub CmdGrabar_Click()
Dim i As Long
Dim sCuenta As String, sMovNroAnt As String
Dim sMovNro As String
Dim oMov As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim oMant As COMNCaptaGenerales.NCOMCaptaGenerales  'NCapMantenimiento
Dim rsOrd As New ADODB.Recordset

Set rsOrd = grdOrdPag.GetRsNew()
Set oMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set oMov = New COMNContabilidad.NCOMContFunciones

sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
oMant.ActualizaCapOrdPagEstado gdFecSis, rsOrd, sMovNro, gCapTalOrdPagEstEntregado, gCapTalOrdPagEstSolicitado

'By Capi 21012009
    Do While Not rsOrd.EOF
        If rsOrd("Act") = "1" Then
        
            objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , rsOrd!Cuenta, gCodigoCuenta
            
        End If
    rsOrd.MoveNext
    Loop
'End by
            

'oMant.ImprimeConevenioOP rsOrd  'ppoa

Set oMant = Nothing
Set oMov = Nothing
'fsCadImp = ImprimirOrdenesPago
cmdImprimir.Enabled = True
CmdGrabar.Enabled = False
cmdcancelar.Enabled = False
'ObtieneDatosOrdenPago
End Sub

Private Sub cmdImprimir_Click()
Dim i As Long
Dim sCuerpo As String
Dim sCuenta As String
Dim sNro As String
Dim sNombre As String
Dim o As Long
Dim conta As Integer
Dim sTituloOP As String
Dim sTextImp As String 'RIRO20140905

If Not ExistenItemsMarcados() Then
    MsgBox "Debe seleccionar algún item para este proceso.", vbInformation, "Aviso"
    grdOrdPag.SetFocus
    Exit Sub
End If

For i = 1 To Me.grdOrdPag.Rows - 1
   If grdOrdPag.TextMatrix(i, 1) = "." Then
       Printer.Font.Size = 8
       conta = 1
       sCuenta = grdOrdPag.TextMatrix(i, 2)
       sNombre = Trim(grdOrdPag.TextMatrix(i, 7))
       'ALPA 20091118**********************************************************
       If Mid(sCuenta, 4, 2) = "01" Or Mid(sCuenta, 4, 2) = "04" Or Mid(sCuenta, 4, 2) = "09" Or Mid(sCuenta, 4, 2) = "31" Then
       '***********************************************************************
        If Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional Then
             sTituloOP = gsTitutloOP_Soles
        Else
             sTituloOP = gsTitutloOP_Dolares
        End If
       'ALPA 20091118**********************************************************
       End If
       '***********************************************************************
       
       If gImpresora <> gRICOH Then 'RIRO20140905
            For o = grdOrdPag.TextMatrix(i, 3) To grdOrdPag.TextMatrix(i, 4)
                sNro = Format(o, "00000000")
                Printer.Print Space(20) & sNro
                Printer.Print " "
                Printer.Print " "
                If conta = 1 Then Printer.Print " "
                Printer.Font.Size = 12
                Printer.Print Space(110) & sNro
                Printer.Font.Size = 8
                Printer.Print " "
                Printer.Print " "
                'If Conta <> 2 Then Printer.Print " "
                If conta <> 3 Then Printer.Print " "
                If conta <> 3 Then Printer.Print " "
                If conta <> 4 Then Printer.Print " "
                Printer.Print " "
                Printer.Print " "
                Printer.Print " "
                If conta = 3 Then Printer.Print " "
                If conta = 3 Then Printer.Print " "
                If conta = 4 Then Printer.Print " "
                Printer.Print Space(80) & Trim(sTituloOP)
                Printer.Print Space(80) & sCuenta
                Printer.Print Space(80) & sNombre
                If conta <> 4 Then
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                    Printer.Print " "
                Else
                    Printer.EndDoc
                    conta = 0
                End If
                conta = conta + 1
            Next o
        ' RIRO20140904 ************************************************
        Else
            For o = grdOrdPag.TextMatrix(i, 3) To grdOrdPag.TextMatrix(i, 4)
                'Modificado PASI20161009 ERS0552016**************************
                'sNro = Format(O, "00000000")
                'If conta = 1 Then
                '    Printer.Print " "
                '    Printer.Print " "
                '    Printer.Print " "
                'End If
                'Printer.Print Space(24) & sNro
                'Printer.Print " "
                'Printer.Print " "
                'If conta < 4 Then Printer.Print " "  ' ppp
                'Printer.Font.Size = 12
                'Printer.Print Space(110) & sNro
                'Printer.Font.Size = 8
                'Printer.Print " "
                'Printer.Print " "
                'Printer.Print " "
                'Printer.Print " "
                'Printer.Print " "
                'Printer.Print " "
                'Printer.Print " "
                '
                'Printer.Print " "
                'Printer.Print Space(80) & Trim(sTituloOP)
                'Printer.Print Space(80) & sCuenta
                'Printer.Print Space(80) & sNombre
                'If conta <> 4 Then
                '   Printer.Print " "
                '    Printer.Print " "
                '    Printer.Print " "
                '    Printer.Print " "
                '    Printer.Print " "
                '    Printer.Print " "
                '    If conta > 2 Then Printer.Print " "
                'Else
                '    Printer.EndDoc
                '    conta = 0
                'End If
                'end PASI****************************************************
                
                'PASI20161009 ERS0552016
                'Printer.Print " ":
                sNro = Format(o, "00000000")
                Printer.Print Space(24) & sNro
                Printer.Print " ": Printer.Print " ": Printer.Print " ": Printer.Print " ": Printer.Print " "
                Printer.Font.Size = 12
                Printer.Print Space(125) & sNro
                Printer.Print " ": Printer.Print " ": Printer.Print " ": Printer.Print " ": Printer.Print " ": Printer.Print " "
                Printer.Font.Size = 8
                Printer.Print Space(75) & Trim(sTituloOP)
                Printer.Print Space(75) & sCuenta
                Printer.Print Space(75) & sNombre
                If Not conta = 4 Then Printer.Print " ": Printer.Print " ": Printer.Print " ": Printer.Print " "
                If conta = 4 Then conta = 0: Printer.EndDoc
                conta = conta + 1
                'end PASI****************************************************
            Next o
        End If
        ' END RIRO ****************************************************
        If conta <> 4 Then Printer.EndDoc
   End If
   
Next i

'*** PEAC 20130208
Printer.EndDoc

If MsgBox("Se Imprimio correctamente las Ordenes de Pago", vbYesNo + vbInformation, "AVISO") = vbNo Then
    MsgBox "Vuelva a Imprimir las Ordenes de Pago"
    Exit Sub
End If
CmdGrabar.Enabled = True

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
cmdCancelar_Click
'By Capi 20012009
Set objPista = New COMManejador.Pista
gsOpeCod = gAhoEmisionOrdPago
'End By

End Sub
Private Function ExistenItemsMarcados() As Boolean
Dim i As Long
For i = 1 To grdOrdPag.Rows - 1
    If grdOrdPag.TextMatrix(i, 1) = "." Then
        ExistenItemsMarcados = True
        Exit Function
    End If
Next i
ExistenItemsMarcados = False
End Function
Private Sub GridClear()
grdOrdPag.Rows = 2
grdOrdPag.Clear
grdOrdPag.FormaCabecera
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub optSeleccion_Click(Index As Integer)
Dim i As Long
Select Case Index
    Case 0
        For i = 1 To grdOrdPag.Rows - 1
            grdOrdPag.TextMatrix(i, 1) = "1"
        Next
    Case 1
        For i = 1 To grdOrdPag.Rows - 1
            grdOrdPag.TextMatrix(i, 1) = ""
        Next
End Select
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCuenta As String
    sCuenta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCuenta
End If
End Sub
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim rsCta As New ADODB.Recordset, rsRel As New ADODB.Recordset
Dim oCuenta As COMNCaptaGenerales.NCOMCaptaGenerales
Dim nFila As Long
Dim sPersona As String
Dim oPar As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim nSaldoMinimo As Double, lblMon As String, lblDatosCuenta As String
Dim lbITFCtaExonerada As String

Set oCuenta = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsCta = oCuenta.GetDatosCuenta(sCuenta)
If Not (rsCta.EOF And rsCta.BOF) Then
    Dim nEstado As COMDConstantes.CaptacEstado
    Dim bOrdPag As Boolean
    nEstado = rsCta("nPrdEstado")
    If nEstado <> gCapEstActiva Then
        Select Case nEstado
            Case gCapEstBloqRetiro, gCapEstBloqTotal
                MsgBox "Cuenta Bloqueada.", vbInformation, "Aviso"
            Case gCapEstAnulada, gCapEstCancelada
                MsgBox "Cuenta Cancelada o Anulada.", vbInformation, "Aviso"
        End Select
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusCuenta
        Exit Sub
    End If
    bOrdPag = rsCta("bOrdPag")
    If Not bOrdPag Then
        MsgBox "Cuenta NO fue aperturada con Ordenes de Pago", vbInformation, "Aviso"
        txtCuenta.Cuenta = ""
        txtCuenta.SetFocusCuenta
        Exit Sub
    End If
    lblDatosCuenta = "CUENTA CON ORDEN DE PAGO" & Chr$(13)
    Set oPar = New COMNCaptaGenerales.NCOMCaptaDefinicion
    If CLng(Mid(sCuenta, 9, 1)) = gMonedaNacional Then
        'lblDescuento.BackColor = &H80000005
        lblDatosCuenta = lblDatosCuenta & "MONEDA NACIONAL"
        nSaldoMinimo = oPar.GetCapParametro(gSaldMinAhoMN)
        lblMon = "S/."
    Else
        'lblDescuento.BackColor = &HC0FFC0
        lblDatosCuenta = lblDatosCuenta & "MONEDA EXTRANJERA"
        nSaldoMinimo = oPar.GetCapParametro(gSaldMinAhoME)
        lblMon = "US$"
    End If
    
    lbITFCtaExonerada = fgITFVerificaExoneracion(sCuenta)
    fgITFParamAsume Mid(sCuenta, 4, 2), Mid(sCuenta, 6, 3)
    
    'Me.LblItf.BackColor = lblDescuento.BackColor
    
    Set oPar = Nothing
    lblApertura = Format$(rsCta("dApertura"), "dd-mmm-yyyy")
    lblTipoCuenta = Trim(rsCta("cTipoCuenta"))
    lblFirmas.Caption = rsCta("nFirmas")
    Set rsRel = oCuenta.GetPersonaCuenta(sCuenta)
    sPersona = ""
    Do While Not rsRel.EOF
        If sPersona <> rsRel("cPersCod") Then
            grdCliente.AdicionaFila
            nFila = grdCliente.Rows - 1
            grdCliente.TextMatrix(nFila, 1) = UCase(PstaNombre(rsRel("Nombre")))
            grdCliente.TextMatrix(nFila, 2) = Left(UCase(rsRel("Relacion")), 2)
            sPersona = rsRel("cPersCod")
        End If
        rsRel.MoveNext
    Loop
    rsRel.Close
    Set rsRel = Nothing
    
    'Dim rsHist As ADODB.Recordset
    'Set rsHist = oCuenta.GetHistOrdPagEmision(sCuenta)
    'AgregaHistoria rsHist
    'nMaxNumOP = oCuenta.GetMaxOrdPagEmitida(sCuenta)
    'If nMaxNumOP = 0 Then
    '    lblInicio = Mid(sCuenta, 9, 1) & Format$(1, "0000000")
    'Else
    '    lblInicio = Trim(nMaxNumOP + 1)
    'End If
    'CargaNumTalonario CLng(Mid(sCuenta, 9, 1))
    fraCuenta.Enabled = False
   ' fraHistoria.Enabled = True
    ObtieneDatosOrdenPago
    'fraSolicitud.Enabled = True
    'cmdGrabar.Enabled = True
    'cmdCancelar.Enabled = True
    'cboNumOP.SetFocus
Else
    MsgBox "Cuenta Cancelada, Anulada, o Sin Orden de Pago.", vbInformation, "Aviso"
    cmdCancelar_Click
End If
Set oCuenta = Nothing
End Sub

Private Sub ObtieneDatosOrdenPago()
Dim rsOrden As New ADODB.Recordset
Dim i As Long
Dim oCapMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento

GridClear
Set oCapMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsOrden = oCapMant.GetCapOrdPagEmision_Cuenta(gCapTalOrdPagEstSolicitado, Me.txtCuenta.GetCuenta)
If Not (rsOrden.EOF And rsOrden.BOF) Then
    i = 1
    Do While Not rsOrden.EOF
        If i >= grdOrdPag.Rows Then grdOrdPag.AdicionaFila
        grdOrdPag.TextMatrix(i, 0) = Trim(i)
        grdOrdPag.TextMatrix(i, 1) = "1"
        grdOrdPag.TextMatrix(i, 2) = rsOrden("cCtaCod")
        grdOrdPag.TextMatrix(i, 3) = rsOrden("nInicio")
        grdOrdPag.TextMatrix(i, 4) = rsOrden("nFin")
        grdOrdPag.TextMatrix(i, 5) = Format$(rsOrden("dFecha"), "dd/mm/yyyy")
        grdOrdPag.TextMatrix(i, 6) = rsOrden("nNumTal")
        grdOrdPag.TextMatrix(i, 7) = PstaNombre(rsOrden("cPersNombre"), False)
        grdOrdPag.TextMatrix(i, 8) = rsOrden("nTipo")
        grdOrdPag.TextMatrix(i, 9) = rsOrden("cMovNro")
        i = i + 1
        rsOrden.MoveNext
    Loop
    'cmdRefrescar.Enabled = True
    'cmdGenerar.Enabled = True
    'cmdEstado.Enabled = True
    optSeleccion(0).Enabled = True
    optSeleccion(1).Enabled = True
    cmdImprimir.Enabled = False
    'If nEstadoOrden = gCapTalOrdPagEstRecepcionado Then
    '   Call optSeleccion_Click(1)
    'End If
    
Else
    'cmdRefrescar.Enabled = True
    'cmdGenerar.Enabled = False
    'cmdEstado.Enabled = False
    optSeleccion(0).Enabled = False
    optSeleccion(1).Enabled = False
    cmdImprimir.Enabled = False
    MsgBox "No Existen Ordenes de Pago para este proceso", vbInformation, "Aviso"
End If
Set oCapMant = Nothing
End Sub

