VERSION 5.00
Begin VB.Form frmCajaChicaLista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Caja Chica"
   ClientHeight    =   3735
   ClientLeft      =   810
   ClientTop       =   2520
   ClientWidth     =   9765
   Icon            =   "frmCajaChicaLista.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8400
      TabIndex        =   14
      Top             =   240
      Width           =   1245
   End
   Begin VB.CommandButton cmdExtorno 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7140
      TabIndex        =   13
      Top             =   3240
      Width           =   1245
   End
   Begin VB.CommandButton cmdRendicion 
      Caption         =   "&Rendición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5880
      TabIndex        =   12
      Top             =   3240
      Width           =   1245
   End
   Begin VB.CommandButton cmdSustentacion 
      Caption         =   "Sus&tentación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7140
      TabIndex        =   11
      Top             =   3240
      Width           =   1245
   End
   Begin VB.CommandButton cmdAtender 
      Caption         =   "&Atender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7140
      TabIndex        =   10
      Top             =   3240
      Width           =   1245
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8385
      TabIndex        =   8
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Frame Frame1 
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
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   9540
      Begin Sicmact.FlexEdit fgListaCH 
         Height          =   1995
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3519
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Nro Doc.-Fecha-Solicitante-Importe-cCodPers-Glosa-Area-cMovNroArendir-cCodArea"
         EncabezadosAnchos=   "450-1-1200-4000-1200-0-0-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L-R-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-2-0-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         RowHeightMin    =   150
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
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
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   8100
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   225
         Width           =   1095
         _ExtentX        =   1535
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2310
         TabIndex        =   4
         Top             =   225
         Width           =   4440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   195
         Left            =   6915
         TabIndex        =   3
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7260
         TabIndex        =   2
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7140
      TabIndex        =   9
      Top             =   3240
      Width           =   1245
   End
End
Attribute VB_Name = "frmCajaChicaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oArendir As NARendir
Dim oCH As nCajaChica
Dim lnArendirFase As ARendirFases
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsCtaFondofijo As String
Dim lbSalir As Boolean

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(Optional ByVal pnArendirFase As ARendirFases = ArendirRechazo)
lnArendirFase = pnArendirFase
Me.Show 1
End Sub

Private Sub cmdAtender_Click()
Dim lsMovNroSol As String
Dim lnImporte As Currency
Dim lsMovNro As String
Dim oConFun As NContFunciones
Set oConFun = New NContFunciones
Dim lsTexto As String
Dim oConImp As NContImprimir
Dim lsPersCod As String

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
lsMovNroSol = fgListaCH.TextMatrix(fgListaCH.row, 8)
lnImporte = CCur(fgListaCH.TextMatrix(fgListaCH.row, 4))
lsPersCod = fgListaCH.TextMatrix(fgListaCH.row, 5)
gsDocNro = fgListaCH.TextMatrix(fgListaCH.row, 1)
gdFecha = CDate(fgListaCH.TextMatrix(fgListaCH.row, 2))

'***Comentado por ELRO el 20120723, según OYP-RFC047-2012
'If Len(Trim(txtMovDesc)) = 0 Then
'    MsgBox "Descripcion de la Operación no Válida", vbInformation, "Aviso"
'    txtMovDesc.SetFocus
'    Exit Sub
'End If
'***Fin Comentado por ELRO el 20120723*******************
If MsgBox("Desea Realizar la Atencion del Arendir en efectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    Dim lsSubCta As String
    lsMovNro = oConFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    lsSubCta = oConFun.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaFondofijo, Trim(txtBuscarAreaCH), False)
    lsCtaFondofijo = lsCtaFondofijo + IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
    
    '***Modificado por ELRO el 20120723, según OYP-RFC047-2012
    'If oArendir.GrabaAtencionArendirCH(gArendirTipoCajaChica, gsFormatoFecha, lsMovNroSol, lsMovNro, gsOpeCod, txtMovDesc, _
    '            lnImporte, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaFondofijo, gsDocNro, gdFecha) = 0 Then
    If oArendir.GrabaAtencionArendirCH(gArendirTipoCajaChica, gsFormatoFecha, lsMovNroSol, lsMovNro, gsOpeCod, "Atender Solicitud", _
                lnImporte, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaFondofijo, gsDocNro, gdFecha) = 0 Then
    '***Fin Modificado por ELRO el 20120723*******************
        
        Set oConImp = New NContImprimir
        '***Modificado por ELRO el 20120723, según OYP-RFC047-2012
        'lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, Trim(txtMovDesc), gsNomCmac, gsOpeCod, lsPersCod, Abs(lnImporte), _
        '        gnColPage, gArendirTipoCajaChica, "", False, True, Trim(txtBuscarAreaCH) & "-" & Val(lblNroProcCH), _
        '        lblCajaChicaDesc, ArendirSustentacion, fgListaCH.TextMatrix(fgListaCH.Row, 1), fgListaCH.TextMatrix(fgListaCH.Row, 2))
        lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, "", gsNomCmac, gsOpeCod, lsPersCod, Abs(lnImporte), _
                gnColPage, gArendirTipoCajaChica, "", False, True, Trim(txtBuscarAreaCH) & "-" & Val(lblNroProcCH), _
                lblCajaChicaDesc, ArendirSustentacion, fgListaCH.TextMatrix(fgListaCH.row, 1), fgListaCH.TextMatrix(fgListaCH.row, 2))
        '***Fin Modificado por ELRO el 20120723*******************
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Atendio la Chica del  Area : " & lblCajaChicaDesc & " A Rendir " & fgListaCH.TextMatrix(fgListaCH.row, 1)
            Set objPista = Nothing
            '*******
        EnviaPrevio lsTexto, Me.Caption, gnLinPage / 2
        fgListaCH.EliminaFila fgListaCH.row
        fgListaCH.SetFocus
        '***Comentado por ELRO el 20120723, según OYP-RFC047-2012
        'txtMovDesc = ""
        'txtMovDescAux = ""
        '***Fin Comentado por ELRO el 20120723*******************
    End If
End If
Set oConImp = Nothing

End Sub

Private Sub cmdExtorno_Click()
Dim ldFechaAtenc As Date
Dim lsMovAtenc As String
Dim lsMovNroSol As String
Dim lnImporte As Currency
Dim lsDocTpo As String
Dim lsOpeDoc As String
Dim lsMovNro As String
Dim lsMovNroRend As String
Dim ldFechaRend As Date
Dim lnSaldo As Currency
Dim lsOpeRend As String
Dim oContFunc As NContFunciones
Set oContFunc = New NContFunciones
Dim lsMovNroExtSol  As String '***Agregado por ELRO 20120625, según OYP-RFC047-2012
Dim lnRendido As Currency '***Agregado por ELRO 20120705, según OYP-RFC047-2012
Dim oDOperacion As DOperacion '***Agregado por ELRO 20120705, según OYP-RFC047-2012

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
'***Modificado por ELRO el 20120620, según OYP-RFC047
'If Len(Trim(txtMovDesc)) = 0 Then
'    MsgBox "Ingrese la descripcion del extorno ", vbInformation, "Aviso"
'    txtMovDesc.SetFocus
'    Exit Sub
'End If
'***Fin Modificacion por ELRO************************
lnImporte = CCur(fgListaCH.TextMatrix(fgListaCH.row, 4))
lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.row, 5))
lsMovAtenc = fgListaCH.TextMatrix(fgListaCH.row, 9)
lsMovNroSol = fgListaCH.TextMatrix(fgListaCH.row, 8)
'ldFechaAtenc = CDate(fgListaCH.TextMatrix(fgListaCH.Row, 2))
lsMovNroRend = fgListaCH.TextMatrix(fgListaCH.row, 14)

'If lsMovNroRend <> "" Then
    'ldFechaRend = CDate(Mid(lsMovNroRend, 7, 2) & "/" & Mid(lsMovNroRend, 5, 2) & "/" & Mid(lsMovNroRend, 1, 4))
'End If

lsOpeDoc = gsOpeCod
If MsgBox("Desea Realizar el extorno??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Select Case lnArendirFase
        Case ArendirExtornoAtencion
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            '***Modificado por ELRO el 20120620, según OYP-RFC047
            'If ldFechaAtenc <> gdFecSis Then
                'oArendir.ExtornaArendir lsMovNro, lsOpeDoc, txtMovDesc, _
                '           lsMovAtenc, lsMovNroSol, gArendirTipoCajaChica, lnImporte, _
                '            lnSaldo, lnArendirFase, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
                oArendir.ExtornaArendir lsMovNro, lsOpeDoc, "Extorno de Atención", _
                           lsMovAtenc, lsMovNroSol, gArendirTipoCajaChica, lnImporte, _
                           lnSaldo, lnArendirFase, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
                
                lsMovNroExtSol = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
               
                oArendir.GrabaRechazoSolARendir lsMovNroExtSol, IIf(Mid(gsOpeCod, 3, 1) = "1", gCHEgreDirectoRechMN, gCHEgreDirectoRechME), lsMovNroSol, "Rechazado la Solcitud " & fgListaCH.TextMatrix(fgListaCH.row, 1)
                           
                'ImprimeAsientoContable lsMovNro
            'Else
                'oArendir.EliminaArendir lsMovNro, lsOpeDoc, txtMovDesc, lsMovAtenc, _
                '                        lsMovNroSol, gArendirTipoCajaChica, lnImporte, _
                '                        lnSaldo, lnArendirFase, False, _
                '                        Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
            '    oArendir.EliminaArendir lsMovNro, lsOpeDoc, "Extorno de Atención", lsMovAtenc, _
                                        lsMovNroSol, gArendirTipoCajaChica, lnImporte, _
                                        lnSaldo, lnArendirFase, False, _
                                        Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
                                        
            '    lsMovNroExtSol = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
               
            '    oArendir.GrabaRechazoSolARendir lsMovNroExtSol, IIf(Mid(gsOpeCod, 3, 1) = "1", gCHEgreDirectoRechMN, gCHEgreDirectoRechME), lsMovNroSol, "Rechazado la Solcitud " & fgListaCH.TextMatrix(fgListaCH.Row, 1)
            'End If
            '***Fin Modificado por ELRO**************************
        Case ArendirExtornoRendicion
            '***Agregado por ELRO el 20120705, según OYP-RFC047-2012
            If gsOpeCod = CStr(gCHArendirCtaRendExtMN2) Or gsOpeCod = CStr(gCHArendirCtaRendExtME2) Then
                lnRendido = CCur(fgListaCH.TextMatrix(fgListaCH.row, 5))
                Set oDOperacion = New DOperacion
                
                lsCtaArendir = oDOperacion.EmiteOpeCta(IIf(lnRendido = 0, gCHArendirCtaRendExtExactMN, gCHArendirCtaRendExtIngMN), "H")
                lsCtaPendiente = oDOperacion.EmiteOpeCta(IIf(lnRendido = 0, gCHArendirCtaRendExtExactMN, gCHArendirCtaRendExtIngMN), "H", 1)
                If Trim(lsCtaArendir) = "" Or Trim(lsCtaPendiente) = "" Then
                    MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
                    lbSalir = True
                    Exit Sub
                End If
                Set oDOperacion = Nothing
            End If
            '***Fin Agregado por ELRO*******************************
            lnSaldo = 0
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

            
            '***Modificado por ELRO el 20120620, según OYP-RFC047-2012
            'If ldFechaRend <> gdFecSis Then
                'oArendir.ExtornaArendir lsMovNro, lsOpeDoc, txtMovDesc, _
                '            lsMovNroRend, lsMovNroSol, gArendirTipoCajaChica, lnImporte, _
                '            lnSaldo, lnArendirFase, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
                '
                oArendir.ExtornaArendir lsMovNro, lsOpeDoc, "Extorno de Rendición", _
                            lsMovNroRend, lsMovNroSol, gArendirTipoCajaChica, lnImporte, _
                            lnSaldo, lnArendirFase, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
                'ImprimeAsientoContable lsMovNro
            'Else
            '    oArendir.EliminaArendir lsMovNro, lsOpeDoc, txtMovdesc, lsMovNroRend, _
            '                lsMovNroSol, gArendirTipoCajaChica, lnImporte, lnSaldo, lnArendirFase, False, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
            '    oArendir.EliminaArendir lsMovNro, lsOpeDoc, "Extorno de Rendición", lsMovNroRend, _
            '                lsMovNroSol, gArendirTipoCajaChica, lnImporte, lnSaldo, lnArendirFase, False, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
            'End If
            '***Fin Modificado por ELRO**************************
    End Select
               'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " Se Extorno de Chica del  Area : " & lblCajaChicaDesc & " A Rendir " & fgListaCH.TextMatrix(fgListaCH.row, 1)
            Set objPista = Nothing
            '*******
    fgListaCH.EliminaFila fgListaCH.row
    MsgBox "Extorno realizado con éxito", vbInformation, "Aviso"
    '***Modificado por ELRO el 20120620, según OYP-RFC047
    'txtMovDesc = ""
    'Me.txtMovDescAux = ""
    '***Fin Modificado por ELRO**************************
End If
Set oContFunc = Nothing
Set oDOperacion = Nothing
End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Dim lnTipoRend As RendicionTipo
Dim lsSubCta As String '***Agregado por ELRO el 20120620, según OYP-RFC047-2012
Dim lsCtaPendiente2 As String '***Agregado por ELRO el 20120620, según OYP-RFC047-2012
Dim oNContFunciones As New NContFunciones '***Agregado por ELRO el 20120620, según OYP-RFC047-2012

Set rs = New ADODB.Recordset

Me.MousePointer = 11

'***Comentado por ELRO el 20130221, según SATI INC1301300007
''***Agregado por ELRO el 20120620, según OYP-RFC047-2012
'If gsOpeCod = CStr(gCHArendirCtaRendExtIngMN) Or gsOpeCod = CStr(gCHArendirCtaRendExtIngME) Or _
'   gsOpeCod = CStr(gCHArendirCtaRendExtMN2) Or gsOpeCod = CStr(gCHArendirCtaRendExtME2) Then
'    lsSubCta = oNContFunciones.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaPendiente, Trim(txtBuscarAreaCH), False)
'
'    If lsSubCta <> "" Then
'       If Mid(txtBuscarAreaCH, 1, 3) = "042" Then 'Recuperaciones
'              lsCtaPendiente2 = lsCtaPendiente & lsSubCta
'
'       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "043" Then  'Secretaria
'              lsCtaPendiente2 = lsCtaPendiente & lsSubCta
'
'       ElseIf Mid(txtBuscarAreaCH, 1, 3) = "023" Then  'Logistica
'              lsCtaPendiente2 = lsCtaPendiente & lsSubCta
'       Else
'              lsCtaPendiente2 = lsCtaPendiente & lsSubCta
'       End If
'
'    Else
'        MsgBox "Sub Cuenta Contable no definida.", vbInformation, "Aviso"
'        Exit Sub
'    End If
'End If
''***Fin Agregado por ELRO*******************************
'***Comentado por ELRO el 20130221**************************

Select Case lnArendirFase
    Case ArendirRechazo, ArendirAtencion
        Set rs = oCH.GetSolicitudesArendirCH(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
    Case ArendirSustentacion, ArendirRendicion
        '***Modificado por ELRO el 20120618, según OYP-RFC047-2012
        'Set rs = oCH.GetCHSustSinRendicion(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir)
        Set rs = oCH.GetCHSustSinRendicion(Mid(txtBuscarAreaCH, 1, 3), _
                                           Mid(txtBuscarAreaCH, 4, 2), _
                                           Val(lblNroProcCH), _
                                           lsCtaArendir, _
                                           IIf(lnArendirFase = ArendirRendicion, 1, 0), _
                                           IIf(lnArendirFase = ArendirRendicion, "", ""))
        '***Fin Modificado por ELRO*******************************
    Case ArendirExtornoAtencion
        Set rs = oCH.GetAtencionesArendir(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir)
    Case ArendirExtornoRendicion
        Select Case gsOpeCod
            '***Modificado por ELRO el 20120620, según OYP-RFC047-2012
            'Case gCHArendirCtaRendExacMN, gCHArendirCtaRendExacME
            Case gCHArendirCtaRendExtExactMN, gCHArendirCtaRendExtExactME
            '***Fin Modificado por ELRO*******************************
                lnTipoRend = Exacta
            Case gCHArendirCtaRendExtIngMN, gCHArendirCtaRendExtIngME
                lnTipoRend = ConIngreso
            Case gCHArendirCtaRendExtEgrMN, gCHArendirCtaRendExtEgrME
                lnTipoRend = ConEgreso
            '***Agregado por ELRO el 20120705, según OYP-RFC047-2012
            Case gCHArendirCtaRendExtMN2, gCHArendirCtaRendExtME2
                lnTipoRend = Todo
            '***Fin Agregado por ELRO*******************************
        End Select
        '***Modificado por ELRO el 20120620, según OYP-RFC047-2012
        'Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente, lnTipoRend)
        If gsOpeCod = CStr(gCHArendirCtaRendExtIngMN) Or gsOpeCod = CStr(gCHArendirCtaRendExtIngME) Then
            '***Modificado por ELRO el 20130221, según SATI INC1301300007
            'Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente2, lnTipoRend)
            Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente, lnTipoRend)
            '***Fin Modificado por ELRO el 20130221**********************
        '***Agregado por ELRO el 20120705, según OYP-RFC047-2012
        ElseIf gsOpeCod = CStr(gCHArendirCtaRendExtMN2) Or gsOpeCod = CStr(gCHArendirCtaRendExtME2) Then
            '***Modificado por ELRO el 20130221, según SATI INC1301300007
            'Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente2, lnTipoRend)
            Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente, lnTipoRend)
            '***Fin Modificado por ELRO el 20130221**********************
        '***Fin Agregado por ELRO*******************************
        Else
            Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente, lnTipoRend)
        End If
        '***Fin Modificado por ELRO*******************************

        
End Select
If Not rs.EOF And Not rs.BOF Then
    Set fgListaCH.Recordset = rs
    fgListaCH.FormatoPersNom (3)
    If fgListaCH.Enabled Then
        Me.fgListaCH.SetFocus
    End If
Else
    MsgBox "No se encuentran solicitudes pendientes de Arendir", vbInformation, "Aviso"
    If txtBuscarAreaCH.Enabled Then
        '***Modificado por ELRO el 20120626, según OYP-RFC047-2012
        Me.cmdProcesar.SetFocus
        '***Fin Modificado por ELRO*******************************
    Else
        cmdSalir.SetFocus
    End If
End If
rs.Close
Set rs = Nothing
Me.MousePointer = 0
End Sub

Private Sub cmdRechazar_Click()
On Error GoTo ErrCmdRechazar
Dim lsMovNro As String
Dim oContFunc As NContFunciones
Set oContFunc = New NContFunciones

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
'***Comentado por ELRO el 20120623, según OYP-RFC047-2012
'If Len(Trim(txtMovDesc)) = 0 Then
'    MsgBox "Descripcion de la Operación no Válida", vbInformation, "Aviso"
'    txtMovDesc.SetFocus
'    Exit Sub
'End If
'***Fin Comentado por ELRO*******************************
If MsgBox(" ¿ Seguro de Rechazar A Rendir N° :" & fgListaCH.TextMatrix(fgListaCH.row, 1) & vbCrLf & vbCrLf & "Solicitado por :" & fgListaCH.TextMatrix(fgListaCH.row, 3), vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    '***Modificado por ELRO el 20120623, según OYP-RFC047-2012
    'If oArendir.GrabaRechazoSolARendir(lsMovNro, gsOpeCod, fgListaCH.TextMatrix(fgListaCH.Row, 8), Trim(txtMovDesc)) = 0 Then
    If oArendir.GrabaRechazoSolARendir(lsMovNro, gsOpeCod, fgListaCH.TextMatrix(fgListaCH.row, 8), "Rechazado la Solcitud " & fgListaCH.TextMatrix(fgListaCH.row, 1)) = 0 Then
    '***Fin Modificado por ELRO*******************************
        fgListaCH.EliminaFila fgListaCH.row
        fgListaCH.SetFocus
        '***Comentado por ELRO el 20120623, según OYP-RFC047-2012
        'txtMovDesc = ""
        '***Comentado por ELRO***********************************
           'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Rechazo de Chica del  Area : " & lblCajaChicaDesc & " A Rendir " & fgListaCH.TextMatrix(fgListaCH.row, 1)
            Set objPista = Nothing
            '*******
    End If
End If
Set oContFunc = Nothing
Exit Sub
ErrCmdRechazar:
    MsgBox "Error N° [" & Err.Number & "]" & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdRendicion_Click()
Dim oContFunc As NContFunciones
Set oContFunc = New NContFunciones
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String, lsMovNroSol As String
Dim lsMovNro As String, lsOpeCod As String
Dim lnSaldoCH As Currency
Dim lsPersCod As String
Dim oConImp As NContImprimir
Dim lsTexto As String
Dim vCtaFondoFijo As String
Dim lsSubCta As String

On Error GoTo cmdRendicionErr

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub
'***Comentado de ELRO el 20120618, según OYP-RFC047-2012
'If Len(Trim(txtMovDesc)) = 0 Then
'    MsgBox "Descripcion de la Operación no Válida", vbInformation, "Aviso"
'    txtMovDesc.SetFocus
'    Exit Sub
'End If
'***Fin Comentado de ELRO*******************************
'***Modificación de ELRO el 20120618, según OYP-RFC047-2012
'lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.Row, 5))
lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.row, 4))
'lsMovNroAtenc = fgListaCH.TextMatrix(fgListaCH.Row, 9)
lsMovNroAtenc = fgListaCH.TextMatrix(fgListaCH.row, 10)
'lsMovNroSol = fgListaCH.TextMatrix(fgListaCH.Row, 8)
lsMovNroSol = fgListaCH.TextMatrix(fgListaCH.row, 9)
'lsPersCod = fgListaCH.TextMatrix(fgListaCH.Row, 13)
lsPersCod = fgListaCH.TextMatrix(fgListaCH.row, 14)
'***Fin Modificación de ELRO*******************************
Set oConImp = New NContImprimir
lsTexto = ""
'***Modificado por ELRO el 20120618, segú OYP-RFC047-2012
'If MsgBox(" ¿ Seguro de Realizar la Rendicion del A Rendir N° :" & fgListaCH.TextMatrix(fgListaCH.Row, 1) & vbCrLf & vbCrLf & "Solicitado por :" & fgListaCH.TextMatrix(fgListaCH.Row, 3), vbQuestion + vbYesNo, "Confirmación") = vbYes Then
If MsgBox(" ¿ Seguro de Realizar la Rendicion del A Rendir N° :" & fgListaCH.TextMatrix(fgListaCH.row, 1) & vbCrLf & vbCrLf & "Solicitado por :" & fgListaCH.TextMatrix(fgListaCH.row, 6), vbQuestion + vbYesNo, "Confirmación") = vbYes Then
'***ELRO Modificado por ELRO*****************************
    lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    lnSaldoCH = oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), SaldoActual)
    Select Case lnSaldo
        Case 0  'RENDICION EXACTA
            lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaRendExacMN, gCHArendirCtaRendExacME)
            lnSaldoCH = 0
            '***Modificado por ELRO el 20120618, según OYP-RFC047-2012
            'oArendir.GrabaRendicionExacta gArendirTipoCajaChica, gsFormatoFecha, lsMovNro, lsOpeCod, Trim(txtMovDesc), _
            '            lsMovNroAtenc, lsMovNroSol, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lnSaldoCH
            oArendir.GrabaRendicionExacta gArendirTipoCajaChica, _
                                          gsFormatoFecha, _
                                          lsMovNro, _
                                          lsOpeCod, _
                                          "RENDICION EXACTA", _
                                          lsMovNroAtenc, _
                                          lsMovNroSol, _
                                          Mid(txtBuscarAreaCH, 1, 3), _
                                          Mid(txtBuscarAreaCH, 4, 2), _
                                          Val(lblNroProcCH), _
                                          lnSaldoCH
            '***Fin Modificado por ELRO*******************************

        Case Is < 0 'EGRESO CON EFECTIVO
            lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaRendEgreMN, gCHArendirCtaRendEgreME)
            lsSubCta = oContFunc.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaFondofijo, Trim(txtBuscarAreaCH), False)
            If lsSubCta <> "" Then lsCtaFondofijo = lsCtaFondofijo + IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
            vCtaFondoFijo = lsCtaFondofijo
            
            '***Modificado por ELRO el 20120618, según OYP-RFC047-2012
            'oArendir.GrabaRendicionGiroDocumento gArendirTipoCajaChica, lsMovNro, lsMovNroSol, lsMovNroAtenc, _
            '            lsOpeCod, Trim(txtMovDesc), lsCtaPendiente, vCtaFondoFijo, lsPersCod, lnSaldo, TpoDocVoucherEgreso, _
            '           "", gdFecSis, "", "", "", False, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
            oArendir.GrabaRendicionGiroDocumento gArendirTipoCajaChica, _
                                                 lsMovNro, _
                                                 lsMovNroSol, _
                                                 lsMovNroAtenc, _
                                                 lsOpeCod, _
                                                 "EGRESO CON EFECTIVO", _
                                                 lsCtaPendiente, _
                                                 vCtaFondoFijo, _
                                                 lsPersCod, _
                                                 lnSaldo, _
                                                 TpoDocVoucherEgreso, _
                                                 "", gdFecSis, "", "", _
                                                 "", False, Mid(txtBuscarAreaCH, 1, 3), _
                                                 Mid(txtBuscarAreaCH, 4, 2), _
                                                 Val(lblNroProcCH)
                       
            'lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, Trim(txtMovDesc), gsNomCmac, gsOpeCod, lsPersCod, Abs(lnSaldo), _
            '                    gnColPage, gArendirTipoCajaChica, "", False, True, Trim(txtBuscarAreaCH) & "-" & Val(lblNroProcCH), _
            '                    lblCajaChicaDesc, ArendirRendicion, fgListaCH.TextMatrix(fgListaCH.Row, 1), fgListaCH.TextMatrix(fgListaCH.Row, 2))
            lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, _
                                                         gdFecSis, _
                                                         "EGRESO CON EFECTIVO", _
                                                         gsNomCmac, _
                                                         gsOpeCod, _
                                                         lsPersCod, _
                                                         Abs(lnSaldo), _
                                                         gnColPage, _
                                                         gArendirTipoCajaChica, _
                                                         "", False, True, _
                                                         Trim(txtBuscarAreaCH) & "-" & Val(lblNroProcCH), _
                                                         lblCajaChicaDesc, _
                                                         ArendirRendicion, _
                                                         fgListaCH.TextMatrix(fgListaCH.row, 1), _
                                                         fgListaCH.TextMatrix(fgListaCH.row, 2))
            '***Fin Modificado por ELRO*******************************
            
            EnviaPrevio lsTexto, Me.Caption, gnLinPage / 2
            
        Case Is > 0 'INGRESO CON EFECTIVO
            lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaRendIngMN, gCHArendirCtaRendIngME)
            lsSubCta = oContFunc.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaFondofijo, Trim(txtBuscarAreaCH), False)
            If lsSubCta <> "" Then lsCtaFondofijo = lsCtaFondofijo + IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
            vCtaFondoFijo = lsCtaFondofijo
            '***Modificado por ELRO el 20120618, según OYP-RFC047-2012
            'oArendir.GrabaRendicionGiroDocumento gArendirTipoCajaChica, lsMovNro, lsMovNroSol, lsMovNroAtenc, _
            '            lsOpeCod, Trim(txtMovDesc), vCtaFondoFijo, lsCtaArendir, lsPersCod, Abs(lnSaldo), TpoDocRecEgreso, _
            '            "", gdFecSis, "", "", "", False, Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH)
            oArendir.GrabaRendicionGiroDocumento gArendirTipoCajaChica, _
                                                 lsMovNro, lsMovNroSol, _
                                                 lsMovNroAtenc, _
                                                 lsOpeCod, "INGRESO CON EFECTIVO", _
                                                 vCtaFondoFijo, lsCtaArendir, _
                                                 lsPersCod, Abs(lnSaldo), _
                                                 TpoDocRecEgreso, _
                                                 "", gdFecSis, "", "", "", False, _
                                                 Mid(txtBuscarAreaCH, 1, 3), _
                                                 Mid(txtBuscarAreaCH, 4, 2), _
                                                 Val(lblNroProcCH)
            'lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, gdFecSis, Trim(txtMovDesc), gsNomCmac, gsOpeCod, lsPersCod, Abs(lnSaldo), _
            '                    gnColPage, gArendirTipoCajaChica, "", True, True, Trim(txtBuscarAreaCH) & "-" & Val(lblNroProcCH), _
            '                   lblCajaChicaDesc, ArendirRendicion, fgListaCH.TextMatrix(fgListaCH.Row, 1), fgListaCH.TextMatrix(fgListaCH.Row, 2))
            lsTexto = oConImp.ImprimeReciboIngresoEgreso(lsMovNro, _
                                                         gdFecSis, _
                                                         "INGRESO CON EFECTIVO", _
                                                         gsNomCmac, _
                                                         gsOpeCod, _
                                                         lsPersCod, _
                                                         Abs(lnSaldo), _
                                                         gnColPage, _
                                                         gArendirTipoCajaChica, _
                                                         "", True, True, _
                                                         Trim(txtBuscarAreaCH) & "-" & Val(lblNroProcCH), _
                                                         lblCajaChicaDesc, _
                                                         ArendirRendicion, _
                                                         fgListaCH.TextMatrix(fgListaCH.row, 1), _
                                                         fgListaCH.TextMatrix(fgListaCH.row, 2))
            '***Fin Modificado por ELRO*******************************
            EnviaPrevio lsTexto, Me.Caption, gnLinPage / 2
    End Select
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " A Rendir " & lsPersCod
            Set objPista = Nothing
            '*******
    fgListaCH.EliminaFila fgListaCH.row
    fgListaCH.SetFocus
    '***Comentado de ELRO el 20120618, según OYP-RFC047-2012
    'txtMovDesc = ""
    'txtMovDescAux = ""
    '***Fin Comentado de ELRO*******************************
End If
Set oConImp = Nothing
Set oContFunc = Nothing
Exit Sub
cmdRendicionErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdSustentacion_Click()
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String
Dim lsAreaCh As String
Dim lsAgeCh As String
Dim lnNroProc As Integer

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub

lsNroArendir = fgListaCH.TextMatrix(fgListaCH.row, 1)
lsNroDoc = fgListaCH.TextMatrix(fgListaCH.row, 1)
lsFechaDoc = fgListaCH.TextMatrix(fgListaCH.row, 2)
'***Modificado por ELRO el 20120618, según OYP-RFC047-2012
'lsPersCod = fgListaCH.TextMatrix(fgListaCH.Row, 13)
lsPersCod = fgListaCH.TextMatrix(fgListaCH.row, 14)
'lsPersNomb = fgListaCH.TextMatrix(fgListaCH.Row, 3)
lsPersNomb = fgListaCH.TextMatrix(fgListaCH.row, 6)
'lsAreaCod = fgListaCH.TextMatrix(fgListaCH.Row, 10)
lsAreaCod = fgListaCH.TextMatrix(fgListaCH.row, 11)
'lsAreaDesc = fgListaCH.TextMatrix(fgListaCH.Row, 7)
lsAreaDesc = fgListaCH.TextMatrix(fgListaCH.row, 8)
'***Fin Modificado por ELRO*******************************

lsDescDoc = fgListaCH.TextMatrix(fgListaCH.row, 1)
'***Modificado por ELRO el 20120618, según OYP-RFC047-2012
'lnImporte = CCur(fgListaCH.TextMatrix(fgListaCH.Row, 4))
lnImporte = CCur(fgListaCH.TextMatrix(fgListaCH.row, 3))
'lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.Row, 5))
lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.row, 4))
'lsMovNroAtenc = fgListaCH.TextMatrix(fgListaCH.Row, 9)
lsMovNroAtenc = fgListaCH.TextMatrix(fgListaCH.row, 10)
'lsMovNroSolicitud = fgListaCH.TextMatrix(fgListaCH.Row, 8)
lsMovNroSolicitud = fgListaCH.TextMatrix(fgListaCH.row, 9)
'lsAgeDesc = fgListaCH.TextMatrix(fgListaCH.Row, 11)
lsAgeDesc = fgListaCH.TextMatrix(fgListaCH.row, 12)
'lsAgeCod = fgListaCH.TextMatrix(fgListaCH.Row, 12)
lsAgeCod = fgListaCH.TextMatrix(fgListaCH.row, 13)
'***Fin Modificado por ELRO*******************************

lsAreaCh = Mid(txtBuscarAreaCH, 1, 3)
lsAgeCh = Mid(txtBuscarAreaCH, 4, 2)
lnNroProc = Val(lblNroProcCH)
frmOpeRegDocs.Inicio lnArendirFase, gArendirTipoCajaChica, False, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, lsAreaCh, lsAgeCh, lnNroProc

fgListaCH.TextMatrix(fgListaCH.row, 5) = Format(frmOpeRegDocs.lnSaldo, gsFormatoNumeroView)
fgListaCH.SetFocus

End Sub
'***Comentado por ELRO el 20120618, según OYP-RFC047-2012
'Private Sub fgListaCH_EnterCell()
'If fgListaCH.TextMatrix(1, 0) <> "" Then
'    txtMovDescAux = Me.fgListaCH.TextMatrix(Me.fgListaCH.Row, 6)
'    txtMovDesc = Me.fgListaCH.TextMatrix(Me.fgListaCH.Row, 6)
'End If
'End Sub
'***Fin Comentado por ELRO*******************************

'***Comentado por ELRO el 20120723, según OYP-RFC047-2012
'Private Sub fgListaCH_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    txtMovDesc.SetFocus
'End If
'End Sub
'***Fin Comentado por ELRO el 20120723*******************

'***Comentado por ELRO el 20120618, según OYP-RFC047-2012
'Private Sub fgListaCH_OnRowChange(pnRow As Long, pnCol As Long)
'If fgListaCH.TextMatrix(1, 0) <> "" Then
'    txtMovDescAux = Me.fgListaCH.TextMatrix(Me.fgListaCH.Row, 6)
'End If
'End Sub
'***Fin Comentado por ELRO*******************************

Private Sub Form_Activate()
    If lbSalir Then
        Unload Me
    End If
    txtBuscarAreaCH_EmiteDatos
End Sub

Private Sub Form_Load()
Set oArendir = New NARendir
Set oCH = New nCajaChica
Dim oOpe As DOperacion

lbSalir = False
Me.Caption = gsOpeDesc
Set oOpe = New DOperacion
cmdAtender.Visible = False
cmdRechazar.Visible = False
cmdSustentacion.Visible = False
cmdRendicion.Visible = False
cmdExtorno.Visible = False
Select Case lnArendirFase
    Case ArendirRechazo
        cmdRechazar.Visible = True
    Case ArendirAtencion
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "D")
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "H")
        cmdAtender.Visible = True
        If Trim(lsCtaArendir) = "" Or Trim(lsCtaFondofijo) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
    Case ArendirSustentacion
        Me.cmdSustentacion.Visible = True
        fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha Doc-Solicitante-Importe - Saldo -Glosa-Area-cMovSol -cMovNroAtenc-cAreaCod- Agencia- cAgeCod - cPersCod"
        fgListaCH.EncabezadosAnchos = "450-1300-1100-3000-1200-1200-0-2000-0-0-0-2000-0-0"
        fgListaCH.EncabezadosAlineacion = "C-L-C-L-R-R-L-L-L-C-C-L-L"
        fgListaCH.FormatosEdit = "0-0-0-0-2-2-0-0-0-0-0"
        
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaPendiente = oOpe.EmiteOpeCta(gsOpeCod, "H", 1)
        If Trim(lsCtaArendir) = "" Or Trim(lsCtaPendiente) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
        '***Comentado por ELRO el 20120616, según OYP-RFC047-2012
        'Me.txtMovDesc.Visible = False
        'lblMovDesc.Visible = False
        '***Fin Comentado por ELRO*******************************
    Case ArendirRendicion
        '***Modificado por ELRO 20120618, según OYP-RFC047-2012
        'fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha Doc-Solicitante-Importe - Saldo -Glosa-Area-cMovSol -cMovNroAtenc-cAreaCod- Agencia- cAgeCod - cPersCod"
        fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha-Importe-Saldo-Usuario-Solicitante-Glosa-Area-cMovSol-cMovNroAtenc-cAreaCod-Agencia-cAgeCod-cPersCod-nProcNro"
        'fgListaCH.EncabezadosAnchos = "450-1300-1100-3000-1200-1200-0-2000-0-0-0-2000-0-0"
        fgListaCH.EncabezadosAnchos = "450-0-1100-1000-1000-1000-0-4500-0-0-0-0-0-0-0-0"
        'fgListaCH.EncabezadosAlineacion = "C-L-C-L-R-R-L-L-L-C-C-L-L"
        fgListaCH.EncabezadosAlineacion = "C-L-R-R-C-L-L-L-L-L-L-L-L-L-L-L"
        'fgListaCH.FormatosEdit = "0-0-0-0-2-2-0-0-0-0-0"
        '***Fin Modificado por ELRO****************************
        fgListaCH.FormatosEdit = "0-0-0-2-2-2-0-0-0-0-0-0-0-0-0-0"
        cmdRendicion.Visible = True
        '***Modificado por ELRO el 20120618, según OYP-RFC047-2012
        cmdSustentacion.Visible = True
        'cmdSustentacion.Visible = False
        '***Fin Modificado por ELRO*******************************
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaPendiente = oOpe.EmiteOpeCta(gsOpeCod, "D", 1)
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
        If Trim(lsCtaArendir) = "" Or Trim(lsCtaPendiente) = "" Or Trim(lsCtaFondofijo) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
    Case ArendirExtornoAtencion
        cmdExtorno.Visible = True
        fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha Doc-Solicitante-Importe - Saldo -Glosa-Area-cMovSol -cMovNroAtenc-cAreaCod- Agencia- cAgeCod - cPersCod - cMovRend "
        fgListaCH.EncabezadosAnchos = "450-1300-1100-3000-1200-1200-0-2000-0-0-0-2000-0-0-0"
        fgListaCH.EncabezadosAlineacion = "C-L-C-L-R-R-L-L-L-C-C-L-L"
        fgListaCH.FormatosEdit = "0-0-0-0-2-2-0-0-0-0-0"
        
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "H")
        If Trim(lsCtaArendir) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
    Case ArendirExtornoRendicion
        cmdExtorno.Visible = True
        fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha Doc-Solicitante-Importe Real- Rendido-Glosa-Area-cMovSol -cMovNroAtenc-cAreaCod- Agencia- cAgeCod - cPersCod"
        fgListaCH.EncabezadosAnchos = "450-1400-1100-3000-0-1200-0-2000-0-0-0-2000-0-0"
        fgListaCH.EncabezadosAlineacion = "C-L-C-L-R-R-L-L-L-C-C-L-L"
        fgListaCH.FormatosEdit = "0-0-0-0-2-2-0-0-0-0-0"
        
           
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaPendiente = oOpe.EmiteOpeCta(gsOpeCod, "H", 1)
        If Trim(lsCtaArendir) = "" Or Trim(lsCtaPendiente) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
            


End Select
CentraForm Me
txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
'***Agregado por ELRO el 20120623, según OYP-RFC047-2012
fraCajaChica.Enabled = False
verificarEncargadoCH
If Len(Trim(txtBuscarAreaCH)) = 0 Then
    Exit Sub
End If
'***Fin Agregado por ELRO*******************************
Set oOpe = Nothing
End Sub
Private Sub txtBuscarAreaCH_EmiteDatos()
Dim oCajaCH As nCajaChica
Set oCajaCH = New nCajaChica
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
fgListaCH.Clear
fgListaCH.FormaCabecera
fgListaCH.Rows = 2
'***Comentado por ELRO el 20120616, según OYP-RFC047-2012
'txtMovDescAux = ""
'***Fin Comentado por ELRO*******************************
If lblCajaChicaDesc <> "" Then
    If cmdProcesar.Visible Then
       cmdProcesar.SetFocus
    End If
End If
Set oCajaCH = Nothing
End Sub

Private Sub txtBuscarAreaCH_Validate(Cancel As Boolean)
If txtBuscarAreaCH = "" Then
    Cancel = True
End If
End Sub

'***Comentado por ELRO el 20120723, según OYP-RFC047-2012
'Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    KeyAscii = 0
'    If cmdRechazar.Visible Then
'        cmdRechazar.SetFocus
'    Else
'        If cmdAtender.Visible Then
'            cmdAtender.SetFocus
'        Else
'            If cmdRendicion.Visible Then
'                cmdRendicion.SetFocus
'            ElseIf cmdExtorno.Visible Then
'                cmdExtorno.SetFocus
'            ElseIf cmdSustentacion.Visible Then
'                    cmdSustentacion.SetFocus
'                End If
'        End If
'    End If
'End If
'End Sub
'***Comentado por ELRO el 20120723***********************
'***Agregado por ELRO el 20120623, según OYP-RFC047-2012
Private Sub verificarEncargadoCH()
    Dim oNCajaChica As nCajaChica
    Set oNCajaChica = New nCajaChica
    Dim rsEncargado As ADODB.Recordset
    Set rsEncargado = New ADODB.Recordset
    
    Set rsEncargado = oNCajaChica.verificarEncargadoCH(gsCodPersUser)
    
    If Not rsEncargado.BOF And Not rsEncargado.EOF Then
        txtBuscarAreaCH = rsEncargado!cAreaCod & rsEncargado!cAgeCod
    Else
        MsgBox "No carga el código de la Caja Chica por los siguientes motivos:" & Chr(10) & "1. No esta encargado de la Caja Chica." & Chr(10) & "2. Aún no esta Autorizado el nuevo proceso de la Caja Chica." & Chr(10) & "3. Aún no cobra el efectivo habilitado por la Caja Chica.", vbInformation, "Aviso"
    End If
    Set rsEncargado = Nothing
    Set oNCajaChica = Nothing
End Sub
'***Fin Agregado por ELRO*******************************
