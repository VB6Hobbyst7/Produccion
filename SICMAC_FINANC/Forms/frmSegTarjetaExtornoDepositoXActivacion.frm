VERSION 5.00
Begin VB.Form frmSegTarjetaExtornoDepositoXActivacion 
   Caption         =   "Extorno MN: Extorno Depósito por Activación de Seguro de Tarjeta"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmSegTarjetaExtornoDepositoXActivacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   10695
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "Extornar"
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   9480
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   615
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   9135
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.TextBox txtDescripcionAnt 
         Height          =   615
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   10335
      End
      Begin Sicmact.FlexEdit feExtorno 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   3625
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Institución Financiera-Cuenta-Doc-Nro. Doc-Importe-cMovNro"
         EncabezadosAnchos=   "400-0-1800-1200-1200-1200-1200-1200"
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
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmSegTarjetaExtornoDepositoXActivacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCaja As nCajaGeneral
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub feExtorno_RowColChange()
    'txtDescripcionAnt.Text = feExtorno.TextMatrix(feExtorno.row, 7)'FRHU 20140814: ANEXO 3-ERS068-2014
    txtDescripcionAnt.Text = feExtorno.TextMatrix(feExtorno.row, 18)
    txtDescripcion.Text = txtDescripcionAnt.Text
End Sub
Private Sub Form_Load()
    Dim lsOperacion As String
    Dim oOpe As New DOperacion
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Set oCaja = New nCajaGeneral
    Set rs = New ADODB.Recordset
    
    'FRHU 20140814: ANEXO 3-ERS068-2014
    'feExtorno.EncabezadosAnchos = "350-0-2600-1700-400-1200-1200-0-0-0-0-2600-0-0-0"
    'feExtorno.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo-nImporteIntDev"
    'feExtorno.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L"
    'feExtorno.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0"
    
    feExtorno.EncabezadosAnchos = "350-0-2600-1700-400-1200-1200-0-0-0-0-0-0-0-0-0-2600-0-0"
    feExtorno.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo-nImporteIntDev-cOpeCod-cMovOpe-nMovOpe-cMovDescOpe"
    feExtorno.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L-C-C-C-L"
    feExtorno.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0-0-0-0-0"
    'FIN FRHU 20140814: ANEXO 3-ERS068-2014
    
    lsOperacion = oOpe.GetOperacionRefencia(gsOpeCod)
    'Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, gdFecSis, gdFecSis) 'FRHU 20140814: ANEXO 3-ERS068-2014
    Set rs = oCaja.GetCajaBancosOperacionesSegTarjeta(lsOperacion, gdFecSis, gdFecSis)
    
    feExtorno.Clear
    feExtorno.FormaCabecera
    feExtorno.Rows = 2
    If Not rs.BOF And Not rs.EOF Then
        Set feExtorno.Recordset = rs
        feExtorno.row = 1
        feExtorno.col = 1
    Else
         MsgBox "Datos no encontrados para proceso seleccionado", vbInformation, "Aviso"
    End If
    
    feExtorno_RowColChange
    RSClose rs
    Set rs = Nothing
    Set oCaja = Nothing
End Sub
Private Sub cmdExtornar_Click()
Dim oCon As NContFunciones
Dim oOpe As DOperacion
Dim lnMovNro As String
Dim lsMovNroExt As String
Dim lnNumTran As Long
Dim lnImporte As Currency
Dim ldFechaMov As Date
Dim lbEliminaMov As Boolean
Dim lsMovNro As String
Dim lsAgeCodRef As String
Dim lsDocNRo As String
Dim lsObjetoCod As String
Dim lsAreaCod As String
Dim lsAgeCod  As String
Dim ldFecReg  As Date
Dim oCaja As New nCajaGeneral
Dim oCajaIF As New DCajaCtasIF
Dim lsPersCod As String
Dim lsIFTpo As String
Dim lsCtaIFCod As String
Dim lnNroCuota As Integer
Dim lsOpeCod As String 'FRHU 20140814: ANEXO 3-ERS068-2014

On Error GoTo ExtornarErr

Set oCon = New NContFunciones
Set oOpe = New DOperacion

    If feExtorno.TextMatrix(1, 1) = "" Then
        MsgBox "No existen Movimientos para Extornar", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    
    If Len(Trim(txtDescripcion.Text)) = 0 Then
        MsgBox "Falta indicar motivo de Extorno", vbInformation, "Aviso"
        txtDescripcion.SetFocus
        Exit Sub
    End If

    lnMovNro = 0
    lsDocNRo = ""
    lsObjetoCod = ""
    'FRHU 20140814: ANEXO 3-ERS068-2014
    'lnMovNro = feExtorno.TextMatrix(feExtorno.row, 12)
    'lsMovNro = feExtorno.TextMatrix(feExtorno.row, 11)
    lnMovNro = feExtorno.TextMatrix(feExtorno.row, 17)
    lsMovNro = feExtorno.TextMatrix(feExtorno.row, 16)
    'FIN FRHU 20140814: ANEXO 3-ERS068-2014
    lnImporte = CCur(feExtorno.TextMatrix(feExtorno.row, 6))
    ldFechaMov = CDate(feExtorno.TextMatrix(feExtorno.row, 1))
    lsAgeCodRef = "" 'Trim(feExtorno.TextMatrix(feExtorno.row, 15))
    lnNumTran = feExtorno.TextMatrix(feExtorno.row, 14)
    lsOpeCod = Trim(feExtorno.TextMatrix(feExtorno.row, 15)) 'FRHU 20140814: ANEXO 3-ERS068-2014
    
    If lnMovNro = 0 Or lsMovNro = "" Then
        MsgBox "Aún no esta implementado Extorno de esta Operación. Consultar con Sistemas", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    
    'FRHU 20140531 ERS068-2014
    Dim oDMov As New DMov
    Dim rsMov As ADODB.Recordset
    Set rsMov = oDMov.ObtenerSolicitudesParaExtornarSegTarjeta(lsMovNro, 2)
    If Not rsMov.BOF And Not rsMov.EOF Then
        MsgBox "No se puede extornar porque la solicitud fue aceptada en el Sicmact Negocio", vbInformation, "Aviso"
        Exit Sub
    End If
    'FIN FRHU 20140531
    lsMovNroExt = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If MsgBox("Desea Realizar el Extorno respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        'FRHU 20140814: ANEXO 3-ERS068-2014
        'Dim oFun As New NContFunciones
        'lbEliminaMov = oFun.PermiteModificarAsiento(lsMovNro, False)
        'If oCaja.GrabaExtornoMov(gdFecSis, ldFechaMov, lsMovNroExt, lnMovNro, gsOpeCod, txtDescripcion.Text, lnImporte, lsMovNroExt, lbEliminaMov, lsMovNro, lsAgeCodRef, lsDocNRo, lsObjetoCod, gbBitCentral, ldFecReg, lnNumTran) = 0 Then
        '    Set oCon = Nothing
        '    If Not lbEliminaMov Then
        '        ImprimeAsientoContable lsMovNroExt, , , , True, False, txtDescripcion.Text
        '    End If
        '    feExtorno.EliminaFila feExtorno.row
        '    If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        '        txtDescripcion.Text = ""
        '        txtDescripcionAnt.Text = ""
        '        feExtorno.SetFocus
        '    Else
        '        Unload Me
        '        Exit Sub
        '    End If
        'End If
        'cmdExtornar.Enabled = True
        If lsOpeCod = "401587" Then
            Dim oFun As New NContFunciones
            lbEliminaMov = oFun.PermiteModificarAsiento(lsMovNro, False)
            If oCaja.GrabaExtornoMov(gdFecSis, ldFechaMov, lsMovNroExt, lnMovNro, gsOpeCod, txtDescripcion.Text, lnImporte, lsMovNroExt, lbEliminaMov, lsMovNro, lsAgeCodRef, lsDocNRo, lsObjetoCod, gbBitCentral, ldFecReg, lnNumTran) = 0 Then
                Set oCon = Nothing
                If Not lbEliminaMov Then
                    ImprimeAsientoContable lsMovNroExt, , , , True, False, txtDescripcion.Text
                End If
                feExtorno.EliminaFila feExtorno.row
                If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    txtDescripcion.Text = ""
                    txtDescripcionAnt.Text = ""
                    feExtorno.SetFocus
                Else
                    Unload Me
                    Exit Sub
                End If
            End If
        Else
            'Extorna solo la referencia porque no se genero asiento
            Call oCaja.GrabaExtornoMovDepositoSegTarjeta(lsMovNroExt, lnMovNro)
            feExtorno.EliminaFila feExtorno.row
            If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                txtDescripcion.Text = ""
                txtDescripcionAnt.Text = ""
                feExtorno.SetFocus
            Else
                Unload Me
                Exit Sub
            End If
        End If
        cmdExtornar.Enabled = True
        'FIN FRHU 20140814
    End If
    
Exit Sub
ExtornarErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
    cmdExtornar.Enabled = True
End Sub
