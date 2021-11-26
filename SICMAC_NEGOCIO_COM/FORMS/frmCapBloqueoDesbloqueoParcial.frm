VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapBloqueoDesbloqueoParcial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
      ForeColor       =   &H00000080&
      Height          =   2445
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   9060
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3780
         TabIndex        =   17
         Top             =   250
         Width           =   435
      End
      Begin VB.Frame fraDatosCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   8790
         Begin SICMACT.FlexEdit grdCliente 
            Height          =   1395
            Left            =   120
            TabIndex        =   10
            Top             =   180
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   2461
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Nombre-RE"
            EncabezadosAnchos=   "250-4000-600"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   255
            RowHeight0      =   285
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   6480
            TabIndex        =   16
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   6480
            TabIndex        =   15
            Top             =   240
            Width           =   2190
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   6480
            TabIndex        =   14
            Top             =   960
            Width           =   2160
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5520
            TabIndex        =   13
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5520
            TabIndex        =   12
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5520
            TabIndex        =   11
            Top             =   1020
            Width           =   540
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   105
         TabIndex        =   18
         Top             =   250
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   4410
         TabIndex        =   19
         Top             =   210
         Width           =   3690
      End
   End
   Begin VB.Frame fraBloqueo 
      ForeColor       =   &H00000080&
      Height          =   4170
      Left            =   60
      TabIndex        =   3
      Top             =   2460
      Width           =   9060
      Begin VB.CommandButton cmdNuevoBlq 
         Caption         =   "&Nuevo Blq."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   3720
         Width           =   1275
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   1000
      End
      Begin TabDlg.SSTab tabBloqueo 
         Height          =   3435
         Left            =   105
         TabIndex        =   6
         Top             =   210
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   6059
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Bloqueo/Desbloqueo Parcial"
         TabPicture(0)   =   "frmCapBloqueoDesbloqueoParcial.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdBloqueo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin SICMACT.FlexEdit grdBloqueo 
            Height          =   2970
            Left            =   45
            TabIndex        =   7
            Top             =   420
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   5239
            Cols0           =   13
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Est-Fecha Blq-Hora Blq-Motivo-Monto-Usu Blq-Fecha Dbl-Hora Dbl-Usu Dbl-Comentario-cMovNro-Flag"
            EncabezadosAnchos=   "250-400-900-0-3500-1200-700-900-0-700-4500-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-X-4-5-X-X-X-X-10-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-4-0-0-3-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-C-C-L-R-C-C-C-C-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
         End
      End
      Begin VB.Label lblMontoAcumBlq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4620
         TabIndex        =   21
         Top             =   3720
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto Bloqueado"
         Height          =   195
         Left            =   2940
         TabIndex        =   20
         Top             =   3810
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8055
      TabIndex        =   2
      Top             =   6735
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6975
      TabIndex        =   1
      Top             =   6735
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   6735
      Width           =   1000
   End
End
Attribute VB_Name = "frmCapBloqueoDesbloqueoParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto
Public bConsulta As Boolean
Dim bBloqTotal As Boolean
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

'ande 20171013 mejora en el combo para seleccionar los motivos de bloqueos
Dim nFilaNuevoBloqueo As Integer



Private Sub ImprimeBoletas(ByVal rsRet As ADODB.Recordset)

    Dim sFlag As String, sMotivo As String
    Dim bEstado As Boolean
    Dim oCap As COMNCaptaGenerales.NCOMCaptaImpresion 'NCapImpBoleta
    Dim lsCliente As String
    Dim clsTit As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim lcCtaCod As String
    Dim lsTit As String, sComentario As String
    Dim lsCadImpB As String
    Dim lsCadImpBT As String
    Dim nFicSal As Integer
    Dim OptBt2 As String
    
    lcCtaCod = txtCuenta.NroCuenta
    
    Set clsTit = New COMNCaptaGenerales.NCOMCaptaGenerales
        lsCliente = ImpreCarEsp(clsTit.GetNombreTitulares(lcCtaCod))
    Set clsTit = Nothing
    
    If Not (rsRet Is Nothing) Then
         If bBloqTotal = True Then
                rsRet.MoveLast
                Do While Not rsRet.EOF
                    bEstado = IIf(rsRet("Est") = "0", False, True)
                    sFlag = rsRet("Flag")
                    sMotivo = rsRet("Motivo")
                    sComentario = Trim(rsRet("Comentario"))
                    
                    lsTit = IIf(bEstado, "Bloqueo", "Desbloqueo") & " Por " & sMotivo
                    Set oCap = New COMNCaptaGenerales.NCOMCaptaImpresion
                        lsCadImpB = lsCadImpB & gPrnSaltoLinea & oCap.ImprimeBoleta(lsTit, sComentario, "", "", lsCliente, lcCtaCod, "", lblMontoAcumBlq, 0, "", 0, 0, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt)
                    Set oCap = Nothing
                                          
            '        If sFlag = "N" And bEstado Then
            '            'Imprime boleta de nuevo bloqueo
            '            oCap.ImprimeBoleta lsTit, "", "", "", lsCliente, lcCtaCod, "", 0, 0, "", 0, 0, , , , , , , , , , gdFecSis, gsNomAge, gsCodUser, sLpt
            '
            '        ElseIf sFlag = "A" And Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            '            'Imprimir Boleta de desbloqueo
            '            'clsMant.ActualizaBloqueoRet lcCtaCod, sComentario, sMovNro, nMotRet, sMov
            '        End If
            
                    rsRet.MoveNext
                Loop
         
         Do
             OptBt2 = MsgBox("Desea Imprimir la Boleta", vbInformation + vbYesNo, "Aviso")
             If vbYes = OptBt2 Then
             nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
                    Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
                    Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
                    Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
                    Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
                    Print #nFicSal, lsCadImpB
                    Print #nFicSal, ""
                    Close #nFicSal
             End If
           Loop Until OptBt2 = vbNo
         End If
    End If
    
    
  
End Sub

Private Function ExisteBloqueoNuevo(ByRef grid As FlexEdit, ByVal nMotivo As Long) As Boolean
Dim i As Integer
Dim sFlag As String
Dim nMotivoNuevo As Long, nFila As Long
If grid.TextMatrix(1, 0) <> "" Then
    nFila = grid.row
    For i = 1 To grid.rows - 1
        sFlag = grid.TextMatrix(i, 12) '** Juez 20121214 Se cambio indice 9>>12
        If sFlag = "N" And i <> nFila Then
            If grid.TextMatrix(i, 4) <> "" Then '** Juez 20121214 Se cambio indice 3>>4
                nMotivoNuevo = CLng(Trim(Right(grid.TextMatrix(i, 4), 4))) '** Juez 20120522 Se cambió indice 3>>4
                If nMotivo = nMotivoNuevo Then
                    ExisteBloqueoNuevo = True
                    Exit Function
                End If
            End If
        End If
    Next i
End If
ExisteBloqueoNuevo = False
End Function

Private Function ValidaDatosBloqueo(ByRef grid As FlexEdit) As Boolean
Dim i As Integer
Dim sFlag As String
Dim nMontoAcum As Double
nMontoAcum = 0
If grid.TextMatrix(1, 0) <> "" Then
    For i = 1 To grid.rows - 1
        sFlag = grid.TextMatrix(i, 12) '** Juez 20120522 Se cambio indice 9>>12
        If grid.TextMatrix(i, 4) = "" And (sFlag = "N" Or sFlag = "A") Then '** Juez 20120522 Se cambió indice 3>>4
            MsgBox "Motivo no válido. Seleccione un motivo válido", vbInformation, "Aviso"
            grid.row = i
            grid.col = 3
            ValidaDatosBloqueo = False
            Exit Function
        End If
        If grid.TextMatrix(i, 1) = "." Then
            nMontoAcum = nMontoAcum + CDbl(grid.TextMatrix(i, 5)) '** Juez 20120522 Se cambio indice 4>>5
        End If
    Next i
End If
lblMontoAcumBlq = Format$(nMontoAcum, "#,##0.00")
ValidaDatosBloqueo = True
End Function

Private Sub ResaltaBloqueoActivo(ByRef grid As FlexEdit)
Dim i As Integer, j As Integer
Dim nCol As Long, nFila As Long
Dim nMontoAcum As Double

nMontoAcum = 0
For i = 1 To grid.rows - 1
    If grid.TextMatrix(i, 1) = "." Then
        nFila = grid.row
        nCol = grid.col
        grid.row = i
        For j = 1 To grid.cols - 1
            grid.col = j
            grid.CellBackColor = &HFFFFC0
        Next
        nMontoAcum = nMontoAcum + grid.TextMatrix(i, 5) '** Juez 20120522 Se cambio indide 4>>5
        grid.col = nCol
        grid.row = nFila
    End If
Next i
lblMontoAcumBlq.Caption = Format$(nMontoAcum, "#,##0.00")
End Sub

Public Sub inicia(ByVal nProd As Producto, Optional bCons As Boolean = False)
nProducto = nProd
bConsulta = bCons
Select Case nProd
    Case gCapAhorros
        txtCuenta.Prod = Trim(str(gCapAhorros))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo Parcial - Ahorros"
    Case gCapPlazoFijo
        txtCuenta.Prod = Trim(str(gCapPlazoFijo))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo Parcial - Plazo Fijo"
    Case gCapCTS
        txtCuenta.Prod = Trim(str(gCapCTS))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo Parcial - CTS"
End Select

If bConsulta Then
    cmdGrabar.Visible = False
    grdBloqueo.lbEditarFlex = False
Else
    cmdGrabar.Visible = True
    grdBloqueo.lbEditarFlex = True
End If
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledProd = False
cmdGrabar.Enabled = False
CmdCancelar.Enabled = False
fraBloqueo.Enabled = False
fraCuenta.Enabled = True
Dim rsMotivo As New ADODB.Recordset
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral

Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsMotivo = clsGen.GetConstante(gCaptacMotBloqueoParcial)
    grdBloqueo.CargaCombo rsMotivo
Set clsGen = Nothing

Me.Show 1
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim ssql As String, sMoneda As String, sPersona As String
Dim nRow As Long

Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
Set clsMant = Nothing

bBloqTotal = False

If Not (rsCta.EOF And rsCta.BOF) Then
    grdBloqueo.ColumnasAEditar = "X-1-X-X-X-X-X-X-X-X-10-X" 'ande 20171013
    nEstado = rsCta("nPrdEstado")
    lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy")
    lblEstado = UCase(rsCta("cEstado"))
    nEstado = rsCta("nPrdEstado")
    lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
    If Mid(rsCta("cCtaCod"), 9, 1) = "1" Then
        sMoneda = "NACIONAL"
        lblCuenta.ForeColor = &HC00000
        lblMontoAcumBlq.BackColor = &H80000005
    Else
        sMoneda = "EXTRANJERA"
        lblCuenta.ForeColor = &H8000&
        lblMontoAcumBlq.BackColor = &H80FF80
    End If
    Select Case nProducto
        Case gCapAhorros
            lblCuenta = "AHORROS " & IIf(rsCta("bOrdPag"), "CON ORDEN DE PAGO", "SIN ORDEN DE PAGO") & " - MONEDA " & sMoneda
        Case gCapPlazoFijo
            lblCuenta = "PLAZO FIJO - MONEDA " & sMoneda
        Case gCapCTS
            lblCuenta = "CTS - MONEDA " & sMoneda
    End Select
    Set rsRel = New ADODB.Recordset
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
    Set clsMant = Nothing
    sPersona = ""
    Do While Not rsRel.EOF
        If sPersona <> rsRel("cPersCod") Then
            grdCliente.AdicionaFila
            nRow = grdCliente.rows - 1
            grdCliente.TextMatrix(nRow, 1) = UCase(PstaNombre(rsRel("Nombre")))
            grdCliente.TextMatrix(nRow, 2) = Left(UCase(rsRel("Relacion")), 2)
            sPersona = rsRel("cPersCod")
        End If
        rsRel.MoveNext
    Loop
    rsRel.Close
    Set rsRel = Nothing
    
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    'Obtiene los datos del bloqueo
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsCta = clsMant.GetCapBloqueos(sCta, gCapTpoBlqParcial, gCaptacMotBloqueoParcial)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        Set grdBloqueo.Recordset = rsCta
        ResaltaBloqueoActivo grdBloqueo
    Else
        lblMontoAcumBlq = "0.00"
    End If
    
    Set rsCta = Nothing
    
    CmdCancelar.Enabled = True
    fraBloqueo.Enabled = True
    fraCuenta.Enabled = False
    cmdGrabar.Enabled = True
    cmdNuevoBlq.Enabled = True
    grdBloqueo.SetFocus
    
    If (nEstado = gCapEstAnulada) Or (nEstado = gCapEstCancelada) Then
        MsgBox "La cuenta se encuentra Cancelada o Anulada.", vbInformation, "Aviso"
        CmdCancelar.Enabled = True
        fraBloqueo.Enabled = True
        fraCuenta.Enabled = False
        cmdNuevoBlq.Enabled = False
        cmdGrabar.Enabled = False
        grdBloqueo.SetFocus
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona 'UPersona
Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    
    sPers = clsPers.sPersCod
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    
    Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto, , , , , gsCodAge)
    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = frmCapMantenimientoCtas.inicia
        If clsCuenta.sCtaCod <> "" Then
            txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
            txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
            txtCuenta.SetFocusCuenta
            SendKeys "{Enter}"
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsPers = Nothing
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelar_Click()
grdCliente.Clear
grdCliente.rows = 2
grdCliente.FormaCabecera
grdBloqueo.Clear
grdBloqueo.FormaCabecera
grdBloqueo.rows = 2
lblApertura = ""
lblCuenta = ""
lblTipoCuenta = ""
lblEstado = ""
lblMontoAcumBlq = "0.00"
cmdGrabar.Enabled = False
CmdCancelar.Enabled = False
txtCuenta.Cuenta = ""
txtCuenta.Age = ""
fraBloqueo.Enabled = False
fraCuenta.Enabled = True
txtCuenta.SetFocusCuenta
bBloqTotal = False
End Sub

Private Sub CmdGrabar_Click()

Dim rsRet As ADODB.Recordset
Dim rsRetTmp As ADODB.Recordset
Dim nMontoAcumBlq As Double

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim ClsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim sMovNro As String, sCuenta As String

''AGREGADO POR ANGC 20200306 -  VALIDA FECHA DEL SISTEMA Y EL SERVER
Dim Msj As String
Msj = gVarPublicas.ValidarFechaSistServer
If Msj <> "" Then
    MsgBox Msj, vbInformation, "Aviso"
    Unload frmCapBloqueoDesbloqueoParcial
Else
    If Not ValidaDatosBloqueo(grdBloqueo) Then
        grdBloqueo.SetFocus
        Exit Sub
    End If
    
    If bBloqTotal = False Then
        MsgBox "No existe Operaciones de Bloqueo Recientes", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Está seguro de grabar??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set rsRet = grdBloqueo.GetRsNew()
        
        Set rsRetTmp = rsRet
        
        Set ClsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set ClsMov = Nothing
        sCuenta = txtCuenta.NroCuenta
        nMontoAcumBlq = CDbl(lblMontoAcumBlq.Caption)
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        'ANDE 20171108
        clsMant.ActualizaBloqueosParciales sCuenta, rsRet, sMovNro, nMontoAcumBlq, gdFecSis
        'END ANDE
        'By Capi 21012009
         objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gInsertar, gModificar, sCuenta, gCodigoCuenta
        'End by
    
        Set clsMant = Nothing
        
        ImprimeBoletas rsRetTmp
       
        cmdCancelar_Click
    End If
End If      ''FIN AGREGADO POR ANGC 20200306


End Sub

Private Sub cmdNuevoBlq_Click()
    grdBloqueo.ColumnasAEditar = "X-1-X-X-4-5-X-X-X-X-10-X" 'ande 20171013
    grdBloqueo.AdicionaFila , , True
    grdBloqueo.TextMatrix(grdBloqueo.rows - 1, 1) = "1"
    grdBloqueo.TextMatrix(grdBloqueo.rows - 1, 2) = Format$(gdFecSis, "dd/mm/yyyy")
    grdBloqueo.TextMatrix(grdBloqueo.rows - 1, 5) = "0" '** Juez 20120522 Se cambio indice 4>>5
    grdBloqueo.TextMatrix(grdBloqueo.rows - 1, 6) = gsCodUser '** Juez 20120522 Se cambio indice 5>>6
    grdBloqueo.TextMatrix(grdBloqueo.rows - 1, 12) = "N" '** Juez 20120522 Se cambio indice 10>>12
    nFilaNuevoBloqueo = grdBloqueo.rows - 1  'ande 20171013
    grdBloqueo.SetFocus
    bBloqTotal = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(nProducto, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapBloqDesbParciales
    'End By


End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub grdBloqueo_OnCellChange(pnRow As Long, pnCol As Long)
Dim nMonto As Double
If pnCol = 4 Then
    nMonto = CDbl(grdBloqueo.TextMatrix(pnRow, 5)) '** Juez 20120522 Se cambio indice 4>>5
    lblMontoAcumBlq.Caption = 0
    If nMonto > 0 Then
        lblMontoAcumBlq.Caption = Format$(CDbl(lblMontoAcumBlq.Caption) + nMonto, "#,##0.00")
    End If
End If
End Sub

Private Sub grdBloqueo_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'Modif por JUEZ 20121004 segun OYP-RFC102-2012 Sólo Legal***********************************************
Dim nMonto As Double
Dim oGen As COMDConstSistema.DCOMGeneral
Dim rs As ADODB.Recordset
'Dim lsCodAreaBlq As String
Dim bAreaLegal As Boolean 'JUEZ 20130130
lblMontoAcumBlq.Caption = 0

Set oGen = New COMDConstSistema.DCOMGeneral
Set rs = New ADODB.Recordset
Set rs = oGen.GetDataUser(grdBloqueo.TextMatrix(pnRow, 6), False)
Set oGen = Nothing
'lsCodAreaBlq = rs!cAreaCod
bAreaLegal = IIf(rs!FechaCargo <= CDate(grdBloqueo.TextMatrix(pnRow, 2)) And rs!cAreaCod = "050", True, False) 'JUEZ 20130130

Set rs = Nothing

If grdBloqueo.TextMatrix(pnRow, 1) = "." Then
    If grdBloqueo.TextMatrix(pnRow, 12) = "A" Then '** Juez 20120522 Se cambio indice 10>>12
        grdBloqueo.TextMatrix(pnRow, 7) = "" '** Juez 20120522 Se cambio indice 6>>7
        grdBloqueo.TextMatrix(pnRow, 9) = "" '** Juez 20120522 Se cambio indice 7>>9
        nMonto = CDbl(grdBloqueo.TextMatrix(pnRow, 5)) '** Juez 20120522 Se cambio indice 4>>5
        lblMontoAcumBlq.Caption = Format$(CDbl(lblMontoAcumBlq.Caption) + nMonto, "#,##0.00")
    Else
        grdBloqueo.TextMatrix(pnRow, 1) = ""
    End If
    bBloqTotal = True
Else
   'If lsCodAreaBlq = "050" Then
    If bAreaLegal Then 'JUEZ 20130130
        If gsCodArea = "050" Then
            If grdBloqueo.TextMatrix(pnRow, 12) = "N" Then '** Juez 20120522 Se cambio indice 10>>12
                grdBloqueo.TextMatrix(pnRow, 1) = "1"
            Else
                If grdBloqueo.TextMatrix(pnRow, 7) = "" And grdBloqueo.TextMatrix(pnRow, 9) = "" Then '** Juez 20120522 Se cambio indice 6>>7 y 7>>9
                    grdBloqueo.TextMatrix(pnRow, 7) = Format$(gdFecSis, "dd/mm/yyyy") '** Juez 20120522 Se cambio indice 6>>7
                    grdBloqueo.TextMatrix(pnRow, 9) = gsCodUser '** Juez 20120522 Se cambio indice 7>>9
                    nMonto = CDbl(grdBloqueo.TextMatrix(pnRow, 5)) '** Juez 20120522 Se cambio indice 4>>5
                    lblMontoAcumBlq.Caption = Format$(CDbl(lblMontoAcumBlq.Caption) - nMonto, "#,##0.00")
                End If
            End If
            bBloqTotal = True
        Else
            MsgBox "Ud. no puede desbloquear esta cuenta, solo un personal del Area de Legal podrá realizarlo", vbInformation, "Aviso"
            grdBloqueo.TextMatrix(pnRow, 1) = "1"
            Exit Sub
        End If
    Else
        If grdBloqueo.TextMatrix(pnRow, 12) = "N" Then '** Juez 20120522 Se cambio indice 10>>12
            grdBloqueo.TextMatrix(pnRow, 1) = "1"
        Else
            If grdBloqueo.TextMatrix(pnRow, 7) = "" And grdBloqueo.TextMatrix(pnRow, 9) = "" Then '** Juez 20120522 Se cambio indice 6>>7 y 7>>9
                grdBloqueo.TextMatrix(pnRow, 7) = Format$(gdFecSis, "dd/mm/yyyy") '** Juez 20120522 Se cambio indice 6>>7
                grdBloqueo.TextMatrix(pnRow, 9) = gsCodUser '** Juez 20120522 Se cambio indice 7>>9
                nMonto = CDbl(grdBloqueo.TextMatrix(pnRow, 5)) '** Juez 20120522 Se cambio indice 4>>5
                lblMontoAcumBlq.Caption = Format$(CDbl(lblMontoAcumBlq.Caption) - nMonto, "#,##0.00")
            End If
        End If
        bBloqTotal = True
    End If
End If
'END JUEZ **********************************************************************************************
End Sub

Private Sub grdBloqueo_OnChangeCombo()
Dim nMotivo As Long, nFila As Long
nFila = grdBloqueo.row
'ande 20171013 mejora en evitar que se cargue el combo cuendo ya se ha registrado el bloqueo
If nFila = nFilaNuevoBloqueo Then
    If grdBloqueo.TextMatrix(nFila, 4) <> "" Then '** Juez 20121214 Se cambio indice 3>>4
        nMotivo = CLng(Trim(Right(grdBloqueo.TextMatrix(nFila, 4), 4))) '** Juez 20120522 Se cambio indice 3>>4
        If ExisteBloqueoNuevo(grdBloqueo, nMotivo) Then
            MsgBox "Motivo ya seleccionado, seleccione otro motivo.", vbInformation, "Aviso"
            grdBloqueo.TextMatrix(nFila, 4) = "" '** Juez 20121214 Se cambio indice 3>>4
        End If
    End If
Else
    cmdNuevoBlq.SetFocus
    grdBloqueo.SetFocus
End If
'end ande
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValidaCuenta(Me.txtCuenta.NroCuenta) = False Then Exit Sub
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub
 
Private Function ValidaCuenta(ByVal psCtaCod As String) As Boolean
    If Len(Trim(psCtaCod)) <> 18 Then
        MsgBox "Ingrese un Numero de Cuenta correcta", vbInformation, "Aviso"
        ValidaCuenta = False
    Else
        ValidaCuenta = True
    End If
End Function



