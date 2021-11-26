VERSION 5.00
Begin VB.Form FrmCFHonrarCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Honrar CF  con Credito"
   ClientHeight    =   7560
   ClientLeft      =   2265
   ClientTop       =   1845
   ClientWidth     =   7650
   Icon            =   "FrmCFHonrarCredito.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCFCredito 
      Caption         =   "Creditos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   3735
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   7335
      Begin VB.Frame Frame5 
         Height          =   1875
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   7125
         Begin VB.CommandButton CmdCredEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   5880
            TabIndex        =   35
            Top             =   720
            Width           =   1080
         End
         Begin VB.CommandButton CmdCredNuevo 
            Caption         =   "&Adicionar"
            Height          =   375
            Left            =   5880
            TabIndex        =   34
            Top             =   240
            Width           =   1080
         End
         Begin SICMACT.FlexEdit FECredRef 
            Height          =   1425
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   2514
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-Credito-Monto-Capital-Inter Comp.-Inter. Morat.-Inter. Gracia-Inter. Susp.-Inter. Reprog-Gastos-MontoSol"
            EncabezadosAnchos=   "400-2400-1200-1200-0-0-0-0-0-0-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-R-R-C"
            FormatosEdit    =   "0-0-2-2-2-2-2-2-2-2-0"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   8421376
         End
         Begin VB.Label LblMontoHonrarCred 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   285
            Left            =   5880
            TabIndex        =   38
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label lblTotalHonrarCred 
            AutoSize        =   -1  'True
            Caption         =   "A Honrar :"
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
            Left            =   5880
            TabIndex        =   37
            Top             =   1200
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1425
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   7095
         Begin VB.CommandButton cmdNueCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   345
            Left            =   5400
            TabIndex        =   27
            Top             =   960
            Width           =   1140
         End
         Begin VB.CommandButton cmdNueAceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4200
            TabIndex        =   26
            Top             =   960
            Width           =   1140
         End
         Begin SICMACT.ActXCodCta ActxCta 
            Height          =   420
            Left            =   165
            TabIndex        =   28
            Top             =   210
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   741
            Texto           =   "Credito :"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Saldo Capital "
            Height          =   225
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label lblsalcap 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1200
            TabIndex        =   31
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Titular :"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   525
         End
         Begin VB.Label lbltitular 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1200
            TabIndex        =   29
            Top             =   660
            Width           =   5430
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   120
      TabIndex        =   19
      Top             =   6840
      Width           =   7335
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5940
         TabIndex        =   21
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Afianzado /Acreedor"
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
      Height          =   1005
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   7425
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   40
         Tag             =   "txtcodigo"
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   39
         Tag             =   "txtnombre"
         Top             =   600
         Width           =   4875
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Tag             =   "txtcodigo"
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   11
         Tag             =   "txtnombre"
         Top             =   210
         Width           =   4905
      End
   End
   Begin VB.CommandButton CmdExaminar 
      Caption         =   "E&xaminar..."
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
      Left            =   6240
      TabIndex        =   1
      Top             =   180
      Width           =   1230
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Carta Fianza"
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
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7395
      Begin VB.Label Label9 
         Caption         =   "Analista"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   17
         Top             =   840
         Width           =   3420
      End
      Begin VB.Label lblFecVencCF 
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
         Height          =   285
         Left            =   5700
         TabIndex        =   15
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Left            =   4740
         TabIndex        =   14
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Monto"
         Height          =   255
         Left            =   4740
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4740
         TabIndex        =   7
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Modalidad"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label lblMontoSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   4
         Top             =   540
         Width           =   1590
      End
      Begin VB.Label lblModalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   3
         Top             =   540
         Width           =   3420
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   2
         Top             =   240
         Width           =   1590
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   180
      TabIndex        =   23
      Top             =   120
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   180
      Width           =   1725
   End
End
Attribute VB_Name = "FrmCFHonrarCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFHonrar
'*  CREACION: 10/09/2002     - LAYG
'*************************************************************************
'*  RESUMEN: HONRAR CARTA FIANZA
'***************************************************************************

Option Explicit
Dim vCodCta As String
Dim MatCred() As String

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatosCredito(ActxCta.NroCuenta) Then
            MsgBox "No se Pudo encontrar el Credito o No esta Vigente", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
    ActXCodCta.SetFocus
End Sub

Private Sub CmdCredEliminar_Click()
    FECredRef.EliminaFila FECredRef.Row
End Sub

Private Sub CmdCredNuevo_Click()
    HabiltaIngreso True
    Call LimpiaIngreso
    ActxCta.SetFocusProd
    cmdNueCancelar.Enabled = True
    
End Sub

Private Sub cmdExaminar_Click()
Dim lsCta As String
    'MAVM 20100606
    lsCta = frmCFPersEstado.Inicio(Array(gColocEstHonrada), "Honrar Carta Fianza", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto))
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        Call CargaDatosR(lsCta)
    Else
        Call LimpiarControles
    End If
End Sub

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Sub CargaDatosR(ByVal psCodCta As String)

Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
Dim R As New ADODB.Recordset
Dim loConstante As COMDConstantes.DCOMConstantes 'DConstante
Dim lbTienePermiso As Boolean


'On Error GoTo ErrorCargaDat
    
ActXCodCta.Enabled = False

    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaHonrarCredito(psCodCta)
    Set oCF = Nothing

    If Not R.BOF And Not R.EOF Then
        lblCodigo.Caption = R!cPersCod
        lblNombre.Caption = PstaNombre(R!cPersNombre)
    
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
    
        'MAVM 20100606
        'If Mid(Trim(psCodCta), 9, 1) = "1" Then
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion) '"COMERCIALES "
        'ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
            'lblTipoCF = "MICROEMPRESA "
        'End If
        
        If Mid(Trim(psCodCta), 9, 1) = "1" Then
            lblMoneda = "Soles"
        ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
            lblMoneda = "Dolares"
        End If
        lblAnalista.Caption = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
        lblMontoSol.Caption = IIf(IsNull(R!nMontoSol), "", Format(R!nMontoSol, "#0.00"))
        lblFecVencCF.Caption = IIf(IsNull(R!dVencSol), "", Format(R!dVencSol, "dd/mm/yyyy"))
        'lblFinalidad.Caption = IIf(IsNull(R!cFinalidad), "", R!cFinalidad)

        Set loConstante = New COMDConstantes.DCOMConstantes
            lblModalidad.Caption = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
        Set loConstante = Nothing
        
        'lblEstado.Caption = GetEstCartaFianza(Trim(reg!cEstado))
        'lblModalidad = DatoTablaCodigo("D1", IIf(IsNull(Trim(reg!cModalidad)), "", reg!cModalidad))
        
        fraCFCredito.Enabled = True
        CmdCredNuevo.SetFocus
        CmdGrabar.Enabled = True
    End If
Exit Sub

ErrorCargaDat:
    MsgBox "Error Nº [" & str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    Exit Sub
End Sub



Private Sub cmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza 'NCartaFianza

Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnDevuelta As Integer
Dim lsComenta As String
Dim lnMonto As Double

vCodCta = ActXCodCta.NroCuenta
Call LllenaMatrizCreditos

If MsgBox("Desea Grabar Creditos para Honrar Carta Fianza", vbInformation + vbYesNo, "Honrar Carta Fianza") = vbYes Then

    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        Call loNCartaFianza.nCFHonrarCredito(vCodCta, lsFechaHoraGrab, lsMovNro, MatCred)
    Set loNCartaFianza = Nothing
            
    CmdGrabar.Enabled = False
    LimpiarControles
End If

'  Call RestMonto(vCodCta)

End Sub

Private Sub cmdNueAceptar_Click()
Dim oNegCred As COMNCredito.NCOMCredito 'NCredito
Dim i As Integer
Dim nTipoCambioFijo As Double
Dim oGeneral As COMDConstSistema.DCOMGeneral 'DGeneral
    
    If Len(Trim(lbltitular.Caption)) <= 0 Then
        MsgBox "Digite un Credito", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set oNegCred = New COMNCredito.NCOMCredito
        FECredRef.AdicionaFila
        FECredRef.TextMatrix(FECredRef.Rows - 1, 1) = ActxCta.NroCuenta
        FECredRef.TextMatrix(FECredRef.Rows - 1, 2) = CDbl(Format(CDbl(lblsalcap.Caption), "#0.00"))
        FECredRef.TextMatrix(FECredRef.Rows - 1, 3) = lblsalcap.Caption
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 4) = Format(oNegCred.MatrizInteresCompAFecha(ActxCta.NroCuenta, MatCalend, gdFecSis, nPrestamo, nTasa) + oNegCred.MatrizInteresCompVencidoFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 5) = Format(oNegCred.MatrizInteresMorFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 6) = Format(oNegCred.MatrizInteresGraciaFecha(ActxCta.NroCuenta, MatCalend, gdFecSis, nPrestamo), "#0.00")
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 7) = Format(oNegCred.MatrizInteresSuspensoFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 8) = Format(oNegCred.MatrizInteresReprogramadoFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 9) = Format(oNegCred.MatrizGastosFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
    Set oNegCred = Nothing
    
    Set oGeneral = New COMDConstSistema.DCOMGeneral
        nTipoCambioFijo = oGeneral.EmiteTipoCambio(gdFecSis, TCFijoMes)
    Set oGeneral = Nothing
    If CInt(Mid(ActXCodCta.NroCuenta, 9, 1)) <> CInt(Mid(ActxCta.NroCuenta, 9, 1)) Then
        If CInt(Mid(ActXCodCta.NroCuenta, 9, 1)) = gMonedaNacional Then 'De Dolares a Soles
            FECredRef.TextMatrix(FECredRef.Rows - 1, 10) = Format(CDbl(FECredRef.TextMatrix(FECredRef.Rows - 1, 2)) * nTipoCambioFijo, "#0.00")
        Else 'De Soles a Dolares
            FECredRef.TextMatrix(FECredRef.Rows - 1, 10) = Format(CDbl(FECredRef.TextMatrix(FECredRef.Rows - 1, 2)) / nTipoCambioFijo, "#0.00")
        End If
    Else
        FECredRef.TextMatrix(FECredRef.Rows - 1, 10) = FECredRef.TextMatrix(FECredRef.Rows - 1, 2)
    End If
    
    Call HabiltaIngreso(False)
    Call LimpiaIngreso
    LblMontoHonrarCred.Caption = "0.00"
    For i = 1 To FECredRef.Rows - 1
        LblMontoHonrarCred.Caption = CDbl(LblMontoHonrarCred.Caption) + CDbl(FECredRef.TextMatrix(i, 10))
    Next i
    cmdNueAceptar.Enabled = False
    cmdNueCancelar.Enabled = False
End Sub

Private Sub cmdNueCancelar_Click()
    Call HabiltaIngreso(False)
    Call LimpiaIngreso
    cmdNueAceptar.Enabled = False
    cmdNueCancelar.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    LimpiarControles
End Sub

Sub LimpiarControles()
   ActXCodCta.Enabled = True
   ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
   lblCodigo.Caption = ""
   lblNombre.Caption = ""
   lblCodAcreedor.Caption = ""
   lblNomAcreedor.Caption = ""
   lblTipoCF.Caption = ""
   lblMoneda.Caption = ""
   lblMontoSol.Caption = ""
   lblModalidad.Caption = ""
   lblAnalista.Caption = ""
   lblFecVencCF.Caption = ""
   lblEstado.Caption = ""
   LblMontoHonrarCred.Caption = ""
   FECredRef.Clear
   FECredRef.TextMatrix(0, 1) = "Credito"
   FECredRef.TextMatrix(0, 2) = "Monto"
   FECredRef.TextMatrix(0, 3) = "Capital"
   fraCFCredito.Enabled = False
   CmdGrabar.Enabled = False
End Sub


Private Sub HabiltaIngreso(ByVal pbHabilita As Boolean)
    ActxCta.Enabled = pbHabilita
    'Label5.Enabled = pbHabilita
    'lbltitular.Enabled = pbHabilita
    'Label3.Enabled = pbHabilita
    'lblsalcap.Enabled = pbHabilita
    cmdNueAceptar.Enabled = pbHabilita
    cmdNueCancelar.Enabled = pbHabilita
    'fraDatos.Enabled = Not pbHabilita
End Sub

Private Sub LimpiaIngreso()
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    Me.lbltitular.Caption = ""
    Me.lblsalcap.Caption = ""
End Sub

Private Function CargaDatosCredito(ByVal psCtaCod As String) As Boolean
Dim oDCred As COMDCredito.DCOMCredito 'DCredito
Dim oNCred As COMNCredito.NCOMCredito 'NCredito
Dim R As New ADODB.Recordset

    On Error GoTo ErrorCargaDatos
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosComunes(psCtaCod, True)
    Set oDCred = Nothing
    If Not R.BOF And Not R.EOF Then
        CargaDatosCredito = True
        lbltitular.Caption = Trim(R!cTitular)
        lblsalcap.Caption = Format(R!nSaldo, "#0.00")
        'nPrestamo = CDbl(Format(R!nMontoCol, "#0.00"))
        'dFecVig = CDate(Format(R!dVigencia, "dd/mm/yyyy"))
        'nTasa = CDbl(Format(R!nTasaInteres, "#0.00"))
        R.Close
        Set R = Nothing
        cmdNueAceptar.SetFocus
        
        'Set oNCred = New COMNCredito.NCOMCredito
        '    MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
        '    ActxCta.Enabled = False
        '    cmdNueAceptar.Enabled = True
        '    cmdNueCancelar.Enabled = True
        'Set oNCred = Nothing
    Else
        CargaDatosCredito = False
        R.Close
        ActxCta.Enabled = True
        Set R = Nothing
        lbltitular.Caption = ""
        lblsalcap.Caption = ""

    End If
    Exit Function

ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Function
Private Sub LllenaMatrizCreditos()
    Dim i As Integer
    If Trim(FECredRef.TextMatrix(1, 1)) <> "" Then
        ReDim MatCred(FECredRef.Rows - 1, 10)
        For i = 1 To FECredRef.Rows - 1
            MatCred(i - 1, 0) = FECredRef.TextMatrix(i, 1)
            MatCred(i - 1, 1) = FECredRef.TextMatrix(i, 2)
            'MatCred(i - 1, 2) = FECredRef.TextMatrix(i, 3)
            'MatCred(i - 1, 3) = FECredRef.TextMatrix(i, 4)
            'MatCred(i - 1, 4) = FECredRef.TextMatrix(i, 5)
            'MatCred(i - 1, 5) = FECredRef.TextMatrix(i, 6)
            'MatCred(i - 1, 6) = FECredRef.TextMatrix(i, 7)
            'MatCred(i - 1, 7) = FECredRef.TextMatrix(i, 8)
            'MatCred(i - 1, 8) = FECredRef.TextMatrix(i, 9)
            'MatCred(i - 1, 9) = FECredRef.TextMatrix(i, 10)
        Next i
    Else
        ReDim MatCred(0, 0)
    End If
End Sub
