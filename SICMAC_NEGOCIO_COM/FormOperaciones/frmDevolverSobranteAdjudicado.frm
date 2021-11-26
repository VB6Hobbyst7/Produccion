VERSION 5.00
Begin VB.Form frmDevolverSobranteAdjudicado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolución Sobrantes de Adjudicados"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12495
   Icon            =   "frmDevolverSobranteAdjudicado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Efectuar Pago"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10800
      TabIndex        =   10
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame FRBeneficiario 
      Caption         =   "Datos del Beneficiario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.ComboBox cmbTipoDOI 
         Height          =   315
         ItemData        =   "frmDevolverSobranteAdjudicado.frx":030A
         Left            =   1080
         List            =   "frmDevolverSobranteAdjudicado.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox txtITF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox txtDOI 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4440
         MaxLength       =   11
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin SICMACT.FlexEdit FEDebitosDevolucion 
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   3836
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-NroContrato-Fecha Adj.-Deuda-Tasación-Importe-Estado-Agencia"
         EncabezadosAnchos=   "700-2100-1500-1400-1400-1400-1400-1600"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-4-4-4-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   705
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblDireccion 
         Caption         =   "Label7"
         Height          =   255
         Left            =   7920
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo DOI:"
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
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "TOTAL S/:"
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
         Left            =   8880
         TabIndex        =   14
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ITF:"
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
         Left            =   9480
         TabIndex        =   11
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Caption         =   "Sub Total S/:"
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
         Left            =   8640
         TabIndex        =   7
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
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
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "DOI:"
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
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDevolverSobranteAdjudicado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'** Nombre : frmArqueoTarjDebVent
'** Descripción : Formulario para Devolución en Efectivo de Sobrantes de Joyas Adjudicadas
'** Creación : GIPO, 20161220
'** Referencia : TI-ERS070-2016
'*****************************************************************************************
Dim oVisto As frmVistoElectronico
Dim codigoOperacionDevolucion As String
Dim codigoOperacionAutoriza As String
Dim cUsuarioVisto As String
Dim bDevuelto As Boolean
Dim bAutorizado As Boolean
Dim cCodigoPersCliente As String

Private Sub cmdCancelar_Click()
    limpiarCamposBeneficiario
    txtDOI.Enabled = True
    cmbTipoDOI.Enabled = True
End Sub

Private Sub Cmdguardar_Click()
    cmdGuardar.Enabled = False
    Dim lsBoleta As String
    Dim oNCOMCaptaGenerales As New NCOMCaptaGenerales
    'aca la validación
    Dim rsGeneracionCarta As New ADODB.Recordset
    Set rsGeneracionCarta = oNCOMCaptaGenerales.obtenerFechaGeneracionCarta()
        If Not (rsGeneracionCarta.EOF And rsGeneracionCarta.BOF) Then
            Dim fechaGeneracion As Date
            If rsGeneracionCarta!nConsSisValor <> Noone Then
                fechaGeneracion = rsGeneracionCarta!nConsSisValor
                If mayorATresMeses(fechaGeneracion) Then
                    'comprobar aprobación
                      If AutorizacionAprobada = False Then Exit Sub
                End If
            End If
        Else
            MsgBox "No se podrá efectuar ninguna devolución mientras no se haya registrado la Fecha de Generación de Carta. Por favor, comunique a su Supervisor", vbInformation, "Aviso"
            cmdGuardar.Enabled = True
            Exit Sub
        End If

    If MsgBox("¿Esta seguro que desea efecuar la Devolución?", vbYesNo, "Aviso") = vbYes Then
        Dim oNCOM As NCOMCaptaGenerales
        Set oNCOM = New NCOMCaptaGenerales
        Dim OnComPrint As New NCOMCaptaImpresion
        Dim lsConfirmar As Long
        Dim lsMovNro As String
        Dim lsDescip As String
        Dim cOpecod As String
        cOpecod = gCapConSerPagDeb
        Dim ITF As Currency
        Dim Total As Currency
        
        Dim lbReimp As Boolean
        Dim nFicSal As Integer

        ITF = CCur(txtITF)
        Total = CCur(txtTotal2)
        lsDescrip = "DEVOLUCIÓN DE SOBRANTES DE JOYAS ADJUDICADAS"
        lsMovNro = oNCOM.registrarDevolucionDeSobranteAdjudicado(gdFecSis, Right(gsCodAge, 2), gsCodUser, cUsuarioVisto, codigoOperacionDevolucion, _
                        FEDebitosDevolucion.GetRsNew(), Total, ITF, txtDOI)
        'MARG ERS052-2017----
        Dim oMov As COMDMov.DCOMMov
        Dim lnMovNro As Long
        Set oMov = New COMDMov.DCOMMov
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        oVisto.RegistraVistoElectronico lnMovNro, , gsCodUser, lnMovNro
        'END MARG- -------------
        lsBoleta = OnComPrint.imprimirBoletaDevolucionSobrantesPignoraticio("DEVOLUCIÓN DE SOBRANTES - CRED. PIGNORATICIO", "", txtNombre, lblDireccion, txtDOI, 0, gsOpeCod, _
                              CCur(txtTotal2), CCur(txtITF), FEDebitosDevolucion.GetRsNew(), gsNomAge, lsMovNro, sLpt, gsCodCMAC, gsNomCmac, gbImpTMU)
        
         lbReimp = True
            Do While lbReimp
                 If Trim(lsBoleta) <> "" Then
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, lsBoleta
                        Print #nFicSal, ""
                    Close #nFicSal
                 End If
               
            
                If MsgBox("Devolución efectuada con éxito!¿Desea Reimprimir boleta de Operación?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbReimp = False
                End If
            Loop
        cmdCancelar_Click
        Set oImp = Nothing
        
    Else
        cmdGuardar.Enabled = True
    End If
End Sub

Private Function AutorizacionAprobada() As Boolean
    Dim oNCOM As New NCOMCaptaGenerales
    Dim rsAutorizacion As New ADODB.Recordset
    Set rsAutorizacion = oNCOM.obtenerAutorizacionDevolucion(Trim(txtDOI), gsCodAge, gsCodUser, gdFecSis)
    If Not (rsAutorizacion.EOF And rsAutorizacion.BOF) Then
        If rsAutorizacion!nEstado = 1 Then
             MsgBox "¡Solicitud Aprobada! El usuario " & rsAutorizacion!cUsuarioAutoriza & " ha autorizado la operación." & _
             " Pulse Aceptar para continuar con el proceso."
             AutorizacionAprobada = True
        ElseIf rsAutorizacion!nEstado = 2 Then
            MsgBox "¡Solicitud Rechazada!, El usuario " & rsAutorizacion!cUsuarioAutoriza & " ha rechazado la operación." & Chr$(10) & _
             " Consulte con su Supervisor los detalles del rechazo."
             cmdGuardar.Enabled = False
             AutorizacionAprobada = False
        ElseIf rsAutorizacion!nEstado = 0 Then
             MsgBox "¡Solicitud Pendiente! Aún no se ha Aprobado o Rechazado la Solicitud."
             cmdGuardar.Enabled = True
             AutorizacionAprobada = False
        End If
    Else
        Call oNCOM.registrarAutorizacionDevolucionSobrante(gdFecSis, gsCodAge, gsCodUser, CDbl(txtTotal2), cCodigoPersCliente)
        MsgBox "La carta del cliente se ha generado hace más de 3 meses. Se ha procedido a realizar una Solicitud de Autorización. Por favor comunique a su Supervisor para la Aprobación...", vbInformation, "Solicitud de Autorización"
        cmdGuardar.Enabled = True
        AutorizacionAprobada = False
    End If
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub calcularSubTotal()
    Dim i As Integer
    txtTotal = "0.00"
    For i = 1 To FEDebitosDevolucion.Rows - 1
        'If Trim(FEDebitosDevolucion.TextMatrix(i, 5)) = "." Then
            txtTotal = Format$(CCur(txtTotal) + CCur(FEDebitosDevolucion.TextMatrix(i, 5)), "##,##0.00")
        'End If
    Next i
End Sub


Private Sub Form_Load()
    Call cargarTiposDOI
    codigoOperacionDevolucion = "123100"
    codigoOperacionAutoriza = "999995"
    bDevuelto = False
End Sub

Private Sub cargarTiposDOI()
    CargaComboConstante 1003, cmbTipoDOI 'tipos de DOI
    cmbTipoDOI.ListIndex = 0
End Sub

Private Sub txtDOI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDOI.Text) <> "" Then
            Call cargarDebitoParaSobranteAdjudicados
        End If
    End If
End Sub

Private Sub limpiarCamposBeneficiario()
    txtDOI = ""
    txtNombre = ""
    txtTotal = ""
    txtITF = ""
    txtTotal2.Text = ""
    LimpiaFlex FEDebitosDevolucion
'    LimpiaFlex FEDebitoOtrasAgencias
End Sub

'GIPO Sobrantes de Adjudicados
Private Sub cargarDebitoParaSobranteAdjudicados()
Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
Dim rsDebito As ADODB.Recordset
'Dim rsDebito2 As ADODB.Recordset
Set rsDebito = New ADODB.Recordset
'Set rsDebito2 = New ADODB.Recordset
Dim rsGeneracionCarta As New ADODB.Recordset

txtNombre = ""
txtITF = ""
txtTotal = ""


Set rsDebito = oNCOMCaptaGenerales.obtenerBeneficiarioSobrantesDeAdjudicados(Trim(txtDOI), Trim(Right(cmbTipoDOI.Text, 2)))
'Set rsDebito2 = rsDebito.Clone

LimpiaFlex FEDebitosDevolucion
'LimpiaFlex FEDebitoOtrasAgencias
If Not (rsDebito.BOF And rsDebito.EOF) Then
   Set oVisto = New frmVistoElectronico
   Dim bResultadoVisto As Boolean
   bResultado = oVisto.Inicio(18)
   cUsuarioVisto = oVisto.ObtieneUsuarioVisto
   If bResultado Then 'si el resultado del visto es correcto
'        If Not (rsGeneracionCarta.EOF And rsGeneracionCarta.BOF) Then
            cCodigoPersCliente = rsDebito!cPersCod
            txtNombre = rsDebito!cPersNombre
            lblDireccion = rsDebito!Direccion
            FEDebitosDevolucion.lbEditarFlex = True
            FEDebitosDevolucion.SetFocus
            FEDebitosDevolucion.lbEditarFlex = True
            Dim conteoDev As Integer
            conteoDev = 0
            Do While Not rsDebito.EOF
                'If rsDebito!CAgencia = gsCodAge Then
                    FEDebitosDevolucion.AdicionaFila
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 1) = rsDebito!cCtaCod
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 2) = rsDebito!FechaAdjud
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 3) = Format$(rsDebito!nDeuda, "##,##0.00")
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 4) = Format$(rsDebito!nTasacion, "##,##0.00")
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 5) = Format$(rsDebito!nDevolver, "##,##0.00")
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 6) = rsDebito!cEstadoDevolucion
                    FEDebitosDevolucion.TextMatrix(FEDebitosDevolucion.row, 7) = rsDebito!cAgeDescripcion
                    conteoDev = conteoDev + 1
                    If rsDebito!cEstadoDevolucion = "PENDIENTE" Then
                        bDevuelto = False
                    Else
                        bDevuelto = True
                    End If
                'End If
                rsDebito.MoveNext
            Loop
            
            If conteoDev > 0 Then
                Call calcularSubTotal
                If Not bDevuelto Then
                    cmdGuardar.Enabled = True
                Else
                    cmdGuardar.Enabled = False
                End If
                
            Else
                cmdGuardar.Enabled = False
            End If
            
            cmbTipoDOI.Enabled = False
            txtDOI.Enabled = False
            
'        End If
   
    
   Else
    MsgBox "Existe monto disponible pero se requiere el Visto del Supervisor", vbInformation, "Aviso"
    cmbTipoDOI.Enabled = True
    txtDOI.Enabled = True
   End If
 
Else
    MsgBox "No existe debito a pagar.", vbInformation, "Aviso"
    cmbTipoDOI.Enabled = True
    txtDOI.Enabled = True
End If

End Sub

Private Function mayorATresMeses(fechaGeneracion As Date) As Boolean
    Dim nMeses As Integer
    nMeses = DateDiff("m", fechaGeneracion, gdFecSis)
    If nMeses >= 3 Then
        mayorATresMeses = True
    Else
        mayorATresMeses = False
    End If
End Function

Private Sub txtTotal_Change()
    If Trim(txtTotal) <> "" Then
        Dim nRedondeoITF As Double
        txtITF = Format(fgITFCalculaImpuesto(CCur(txtTotal)), "#,##0.00")
        nRedondeoITF = fgDiferenciaRedondeoITF(CCur(txtITF))
        txtITF = Format(CCur(txtITF) - nRedondeoITF, "#,##0.00")
        txtTotal2 = Format(CCur(txtTotal) - CCur(txtITF), "#,##0.00")
    End If
End Sub
