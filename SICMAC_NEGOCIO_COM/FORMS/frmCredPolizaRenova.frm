VERSION 5.00
Begin VB.Form frmCredPolizaRenova 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renovación de Polizas"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "frmCredPolizaRenova.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   360
      Left            =   7560
      TabIndex        =   12
      Top             =   6480
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   8880
      TabIndex        =   11
      Top             =   6480
      Width           =   1050
   End
   Begin VB.Frame fraAplicacion 
      Caption         =   "Aplicación en cuotas"
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
      Height          =   675
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Width           =   5415
      Begin VB.Label lblPrimaCuota 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDesde 
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
         Height          =   315
         Left            =   840
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Prima Cuota:"
         Height          =   195
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame fraPrima 
      Caption         =   "Prima"
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
      Height          =   1755
      Left            =   4560
      TabIndex        =   9
      Top             =   4440
      Width           =   5295
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "="
         Height          =   195
         Left            =   3675
         TabIndex        =   34
         Top             =   1395
         Width           =   90
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "+"
         Height          =   195
         Left            =   2595
         TabIndex        =   33
         Top             =   1395
         Width           =   90
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "+"
         Height          =   195
         Left            =   1395
         TabIndex        =   32
         Top             =   1395
         Width           =   90
      End
      Begin VB.Label lblPrimaTotal 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3840
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblIGV 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   2760
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblDerechoEm 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1560
         TabIndex        =   29
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblPrimaNeta 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblTCMes 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Prima Total"
         Height          =   195
         Left            =   3960
         TabIndex        =   20
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "IGV"
         Height          =   195
         Left            =   3000
         TabIndex        =   19
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label6 
         Caption         =   "Derecho Emisión"
         Height          =   435
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Prima Neta"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T/C Cierre de Mes:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Frame fraCertificado 
      Caption         =   "Nº de Certificado:"
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
      Height          =   1755
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   4335
      Begin VB.TextBox txtNumCertif 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtValorEdificacion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2280
         TabIndex        =   37
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblClaseInmueble 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Clase de Inmueble:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor Edificación:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº de Certificado:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1245
      End
   End
   Begin VB.Frame fraAseguradora 
      Caption         =   "Aseguradora"
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
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   9735
      Begin SICMACT.TxtBuscar txtAseguradora 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblNombreAseguradora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   24
         Top             =   360
         Width           =   7455
      End
   End
   Begin VB.Frame fraGarantias 
      Caption         =   "Garantìas"
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
      Height          =   2835
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9735
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   255
         Left            =   8400
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optTasacion 
         Caption         =   "Tasación"
         Height          =   255
         Left            =   8400
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         Height          =   360
         Left            =   8520
         TabIndex        =   4
         Top             =   1560
         Width           =   1050
      End
      Begin SICMACT.FlexEdit feGarantias 
         Height          =   2460
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   4339
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Descripcion-Codigo"
         EncabezadosAnchos=   "400-5000-2000"
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
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0"
         BackColorControl=   65535
         BackColorControl=   65535
         BackColorControl=   65535
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483635
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   767
      Texto           =   "Crédito:"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "109"
   End
End
Attribute VB_Name = "frmCredPolizaRenova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre      : frmCredPolizaRenova
'** Descripción : Formulario para la renovacion manual de las polizas
'** Creación    : WIOR, 20140620 10:00:00 AM
'*****************************************************************************************************
Option Explicit
Private pnValorPolizaTC As Double
Private fnIndex As Integer
Private fMatPolizaRenova() As PolizaRenova

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ActxCta.Enabled = False
        CargaDatos (ActxCta.NroCuenta)
    End If
End Sub

Private Sub cmdAplicar_Click()
    If fnIndex <> -1 Then
        Call HabilitaControles(True, , True, , IIf(optTasacion.value, 0, 1))
        lblClaseInmueble.Caption = fMatPolizaRenova(fnIndex).cDesInmueble
        txtNumCertif.Text = fMatPolizaRenova(fnIndex).cNumCertificado
        txtValorEdificacion.Text = Format(fMatPolizaRenova(fnIndex).nValorEdificacion, "###," & String(15, "#") & "#0.00")
        lblDesde.Caption = fMatPolizaRenova(fnIndex).nDesde
    
        If optTasacion.value Then
            Call MostrarDatosARenovar(CDbl(txtValorEdificacion.Text), fMatPolizaRenova(fnIndex).nInmueble, fMatPolizaRenova(fnIndex).cAgeCod, fMatPolizaRenova(fnIndex).nmoneda)
        End If
        
        lblMoneda.Caption = IIf(CInt(fMatPolizaRenova(fnIndex).nMonedaAux) = 1, "S/.", "$.")
        
        txtAseguradora.Text = fMatPolizaRenova(fnIndex).cPersCodAseguradora
        txtAseguradora.SetFocus
        SendKeys "{Enter}"
    Else
        MsgBox "Seleccione una Garantía", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oPers As comdpersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.Inicio(oPers.sPersCod, , , True, ActxCta)
        ActxCta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Sub cmdCancelar_Click()
Limpiar
End Sub

Private Sub cmdGrabar_Click()
Dim bError As Boolean
Dim oCredito As COMNCredito.NCOMCredito
Dim oPoliza As COMDCredito.DCOMPoliza
Dim nMontoCalend As Double
Dim i As Integer
If ValidaDatos Then
    If MsgBox("Desea Guardar los Datos", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCredito = New COMNCredito.NCOMCredito
        Set oPoliza = New COMDCredito.DCOMPoliza
        nMontoCalend = 0
        For i = 0 To UBound(fMatPolizaRenova)
            If i = fnIndex Then
                nMontoCalend = nMontoCalend + CDbl(lblPrimaCuota.Caption)
            Else
                nMontoCalend = nMontoCalend + fMatPolizaRenova(i).nPrimaCuota
            End If
        Next i
        
        
        Call oCredito.RenovarPoliza(Trim(ActxCta.NroCuenta), fMatPolizaRenova(fnIndex).cNumGarant, fMatPolizaRenova(fnIndex).cnumpoliza, IIf(optTasacion.value, 1, 2), _
            Trim(txtAseguradora.Text), Trim(txtNumCertif.Text), fMatPolizaRenova(fnIndex).nmoneda, fMatPolizaRenova(fnIndex).dTasacion, CDbl(txtValorEdificacion.Text), _
            fMatPolizaRenova(fnIndex).nInmueble, CDbl(lblPrimaNeta.Caption), CDbl(lblDerechoEm.Caption), CDbl(lblIGV.Caption), CDbl(lblPrimaTotal.Caption), pnValorPolizaTC, _
            CInt(lblDesde.Caption), CDbl(lblPrimaCuota.Caption), fMatPolizaRenova(fnIndex).dVigencia, fMatPolizaRenova(fnIndex).dVencimiento, fMatPolizaRenova(fnIndex).dVigenciaNueva, _
            fMatPolizaRenova(fnIndex).dVencimientoNueva, nMontoCalend, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), bError)
        
        If Not bError Then
        
            Call oPoliza.ModificaPoliza(fMatPolizaRenova(fnIndex).cnumpoliza, , Trim(txtAseguradora.Text), CDbl(lblPrimaTotal.Caption), , , , , , , , , , CDbl(txtValorEdificacion.Text), Trim(txtNumCertif.Text), CDbl(lblPrimaNeta.Caption), fMatPolizaRenova(fnIndex).nmoneda, , pnValorPolizaTC)
            Call oPoliza.Elimina_GarantPoliza(fMatPolizaRenova(fnIndex).cnumpoliza, fMatPolizaRenova(fnIndex).cNumGarant, fMatPolizaRenova(fnIndex).dTasacion)
            Call oPoliza.Adiciona_GarantPoliza(fMatPolizaRenova(fnIndex).cnumpoliza, fMatPolizaRenova(fnIndex).cNumGarant, fMatPolizaRenova(fnIndex).dTasacion, gdFecSis)
            
            MsgBox "Datos Guardados Satisfactoriamente", vbInformation, "Aviso"
            Limpiar
            Call CargaDatos(ActxCta.NroCuenta)
        End If
    End If
End If
End Sub

Private Sub feGarantias_Click()
fnIndex = feGarantias.row - 1
End Sub

Private Sub Form_Load()
CargaControles
ActxCta.Age = gsCodAge
pnValorPolizaTC = 0
fnIndex = -1
Call HabilitaControles(False, True, True, , 1)
End Sub


Private Sub txtAseguradora_EmiteDatos()
On Error GoTo ErrorPersona
lblNombreAseguradora.Caption = ""

If Trim(txtAseguradora.psCodigoPersona) <> "" Then
    lblNombreAseguradora.Caption = txtAseguradora.psDescripcion
End If

    Exit Sub
ErrorPersona:
    MsgBox err.Description, vbInformation, "Error"
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
Dim oCred As COMDCredito.DCOMCredito
Dim rsGarant As ADODB.Recordset
Dim i As Integer

Call HabilitaControles(True, True)

Set oCred = New COMDCredito.DCOMCredito
Set rsGarant = oCred.RecuperaCredGarantiaPoliza(psCtaCod)
Set oCred = Nothing


If Not (rsGarant.EOF And rsGarant.BOF) Then
    ReDim fMatPolizaRenova(rsGarant.RecordCount - 1)
    For i = 0 To rsGarant.RecordCount - 1
        fMatPolizaRenova(i).cCtaCod = psCtaCod
        fMatPolizaRenova(i).cAgeCod = rsGarant!cAgeCod
        fMatPolizaRenova(i).cNumGarant = rsGarant!cNumGarant
        fMatPolizaRenova(i).cnumpoliza = rsGarant!cnumpoliza
        fMatPolizaRenova(i).cPersCodAseguradora = rsGarant!cPersCodAseg
        fMatPolizaRenova(i).cNumCertificado = rsGarant!cCodPolizaAseg
        fMatPolizaRenova(i).nMonedaAux = CInt(rsGarant!nmoneda)
        fMatPolizaRenova(i).nmoneda = CInt(rsGarant!nMonedaPoliza)
        fMatPolizaRenova(i).dTasacion = CDate(rsGarant!dTasacion)
        fMatPolizaRenova(i).nValorEdificacion = CDbl(rsGarant!nValorEdificacion)
        fMatPolizaRenova(i).nInmueble = CInt(rsGarant!nInmueble)
        fMatPolizaRenova(i).cDesInmueble = rsGarant!Inmueble
        fMatPolizaRenova(i).nDesde = CInt(rsGarant!CuotaDesde)
        fMatPolizaRenova(i).nPrimaCuota = CDbl(rsGarant!MontoPorPoliza)
        fMatPolizaRenova(i).dVigencia = CDate(rsGarant!dVigenciaAseg)
        fMatPolizaRenova(i).dVencimiento = CDate(rsGarant!dVencAseg)
        fMatPolizaRenova(i).dVigenciaNueva = CDate(rsGarant!NuevaVigPoliza)
        fMatPolizaRenova(i).dVencimientoNueva = CDate(rsGarant!NuevaVencPoliza)
        fMatPolizaRenova(i).cDescripcion = rsGarant!cDescripcion
        rsGarant.MoveNext
    Next i
Else
    ReDim fMatPolizaRenova(0)
    MsgBox "Crédito no cuentas con Garantías que cuente con pólizas", vbInformation, "Aviso"
    Exit Sub
End If

LimpiaFlex feGarantias
If fMatPolizaRenova(0).cCtaCod <> "" Then
    For i = 0 To UBound(fMatPolizaRenova)
            feGarantias.AdicionaFila
            feGarantias.TextMatrix(i + 1, 1) = fMatPolizaRenova(i).cDescripcion
            feGarantias.TextMatrix(i + 1, 2) = fMatPolizaRenova(i).cNumGarant
    Next i
End If

End Sub

Private Sub TxtValorEdificacion_GotFocus()
fEnfoque txtValorEdificacion
End Sub

Private Sub TxtValorEdificacion_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtValorEdificacion, KeyAscii, , , True)

If Not IsNumeric(txtValorEdificacion.Text) Then Exit Sub

If KeyAscii = 13 Then
    Call MostrarDatosARenovar(CDbl(txtValorEdificacion.Text), fMatPolizaRenova(fnIndex).nInmueble, fMatPolizaRenova(fnIndex).cAgeCod, fMatPolizaRenova(fnIndex).nmoneda)
End If

If KeyAscii = 8 Then
    lblPrimaNeta.Caption = ""
    lblDerechoEm.Caption = ""
    lblIGV.Caption = ""
    lblPrimaTotal.Caption = ""
    lblPrimaCuota.Caption = ""
End If

End Sub

Private Sub TxtValorEdificacion_LostFocus()
 If Trim(txtValorEdificacion.Text) = "" Then
    txtValorEdificacion.Text = "0.00"
Else
    txtValorEdificacion.Text = Format(txtValorEdificacion.Text, "###," & String(15, "#") & "#0.00")
End If
End Sub

Private Sub CargaControles()
Dim oTipCambio As nTipoCambio
Set oTipCambio = New nTipoCambio
    lblTCMes.Caption = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
Set oTipCambio = Nothing
End Sub

Private Sub Limpiar()
txtAseguradora.Text = ""
txtNumCertif.Text = ""
txtValorEdificacion = ""
lblNombreAseguradora.Caption = ""
lblClaseInmueble.Caption = ""
lblDerechoEm.Caption = ""
lblIGV.Caption = ""
lblDesde.Caption = ""
lblPrimaCuota.Caption = ""
lblPrimaNeta.Caption = ""
lblPrimaTotal.Caption = ""
lblMoneda.Caption = ""
pnValorPolizaTC = 0
fnIndex = -1
Call HabilitaControles(False, IIf(fraGarantias.Enabled, True, False), True)

If Not fraGarantias.Enabled Then
    ActxCta.Enabled = True
    'ActxCta.NroCuenta = ""
    LimpiaFlex feGarantias
End If
End Sub

Private Sub HabilitaControles(ByVal pbHabilita As Boolean, Optional ByVal pbCargaDatos As Boolean = False, Optional ByVal pbAplicar As Boolean = False, _
                            Optional ByVal pbGrabaCancel As Boolean = False, Optional ByVal pnTipoApl As Integer = 0)

If pbCargaDatos Then
    ActxCta.Enabled = Not pbHabilita
    cmdBuscar.Enabled = Not pbHabilita
    CmdAplicar.Enabled = pbHabilita
    fraGarantias.Enabled = pbHabilita
    cmdCancelar.Enabled = pbHabilita
End If

'Cuando Damos Aplicar
If pbAplicar Then
cmdGrabar.Enabled = pbHabilita
txtAseguradora.Enabled = pbHabilita
txtNumCertif.Enabled = pbHabilita

    If pbHabilita Or (Not fraGarantias.Enabled And Not pbCargaDatos) Then
        fraGarantias.Enabled = Not pbHabilita
    End If
    
    If pnTipoApl = 1 Or Not pbHabilita Then   'Manual
        txtValorEdificacion.Enabled = pbHabilita
    End If
End If

End Sub

Private Sub MostrarDatosARenovar(ByVal pnSumaAseg As Double, ByVal pnInmueble As Integer, ByVal psAgeCod As String, ByVal pnMoneda As Integer)
Dim oPoliza As COMDCredito.DCOMPoliza
Dim rsPoliza As ADODB.Recordset

Set oPoliza = New COMDCredito.DCOMPoliza
Set rsPoliza = oPoliza.CargaDatosPrimaNetaPolizaRenova(pnSumaAseg, pnInmueble, psAgeCod)

If Not (rsPoliza.EOF And rsPoliza.BOF) Then
    lblPrimaNeta.Caption = Format(rsPoliza!PrimaNeta, "###," & String(15, "#") & "#0.00")
    lblDerechoEm.Caption = Format(rsPoliza!Drcho_emi, "###," & String(15, "#") & "#0.00")
    lblIGV.Caption = Format(rsPoliza!IGV, "###," & String(15, "#") & "#0.00")
    lblPrimaTotal.Caption = Format(rsPoliza!Total, "###," & String(15, "#") & "#0.00")
    
    If pnMoneda = 1 Then
        pnValorPolizaTC = Round(CDbl(rsPoliza!Total) * CDbl(lblTCMes.Caption), 2)
    Else
        pnValorPolizaTC = CDbl(rsPoliza!Total)
    End If
    Set rsPoliza = oPoliza.MontoPolizaRenova(CInt(Mid(ActxCta.NroCuenta, 9, 1)), CDbl(lblPrimaTotal.Caption), pnValorPolizaTC)

    If Not (rsPoliza.EOF And rsPoliza.BOF) Then
        lblPrimaCuota.Caption = Format(rsPoliza!MontoPorPoliza, "###," & String(15, "#") & "#0.00")
    End If
End If
Set oPoliza = Nothing
Set rsPoliza = Nothing
End Sub

Private Function ValidaDatos() As Boolean
If Len(txtAseguradora.Text) < 13 Or Trim(txtAseguradora.Text) = "" Or Trim(lblNombreAseguradora.Caption) = "" Then
    MsgBox "Seleccione una Aseguradora correctamente", vbInformation, "Aviso"
    txtAseguradora.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(txtNumCertif.Text) = "" Then
    MsgBox "Ingrese el Nº Certificado", vbInformation, "Aviso"
    txtNumCertif.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Not IsNumeric(txtValorEdificacion.Text) Then
    MsgBox "Ingrese el Valor de Edificación", vbInformation, "Aviso"
    txtValorEdificacion.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(lblPrimaNeta.Caption) = "" Then
    MsgBox "Presione Enter en Valor de Edificación para calcular la prima de la póliza", vbInformation, "Aviso"
    txtValorEdificacion.SetFocus
    ValidaDatos = False
    Exit Function
End If

ValidaDatos = True
End Function
