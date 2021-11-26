VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapBloqueoDesbloqueo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "frmCapBloqueoDesbloqueo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6780
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   6780
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   6780
      Width           =   1000
   End
   Begin VB.Frame fraBloqueo 
      Caption         =   "Bloqueo"
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
      Height          =   4170
      Left            =   105
      TabIndex        =   8
      Top             =   2520
      Width           =   9060
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Width           =   1000
      End
      Begin VB.CommandButton cmdNuevoBlq 
         Caption         =   "&Nuevo Blq."
         Height          =   375
         Left            =   7620
         TabIndex        =   18
         Top             =   3720
         Width           =   1275
      End
      Begin TabDlg.SSTab tabBloqueo 
         Height          =   3435
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   8850
         _ExtentX        =   15610
         _ExtentY        =   6059
         _Version        =   393216
         Style           =   1
         Tabs            =   2
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
         TabCaption(0)   =   "Bloqueo Retiro"
         TabPicture(0)   =   "frmCapBloqueoDesbloqueo.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdRetiro"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Bloqueo Total"
         TabPicture(1)   =   "frmCapBloqueoDesbloqueo.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdTotal"
         Tab(1).ControlCount=   1
         Begin SICMACT.FlexEdit grdRetiro 
            Height          =   2970
            Left            =   45
            TabIndex        =   2
            Top             =   420
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   5239
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Est-Fecha Blq-Hora Blq-Motivo-Usu Blq-Fecha Dbl-Hora Dbl-Usu Dbl-Comentario-cMovNro-Flag"
            EncabezadosAnchos=   "250-400-900-900-3500-700-900-900-700-4500-0-0"
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
            ColumnasAEditar =   "X-1-X-X-4-X-X-X-X-9-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-4-0-0-3-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-C-C-L-C-C-C-C-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
         End
         Begin SICMACT.FlexEdit grdTotal 
            Height          =   2970
            Left            =   -74955
            TabIndex        =   19
            Top             =   420
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   5239
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Est-Fecha Blq-Hora Blq-Motivo-Usu Blq-Fecha Dbl-Hora Dbl-Usu Dbl-Comentario-cMovNro-Flag"
            EncabezadosAnchos=   "250-400-900-900-3500-700-900-900-700-4500-0-0"
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
            ColumnasAEditar =   "X-1-X-X-4-X-X-X-X-9-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-4-0-0-3-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-C-C-L-C-C-C-C-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   255
            RowHeight0      =   300
         End
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
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
      Height          =   2445
      Left            =   105
      TabIndex        =   7
      Top             =   45
      Width           =   9060
      Begin VB.Frame fraDatosCuenta 
         Height          =   1710
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   8790
         Begin SICMACT.FlexEdit grdCliente 
            Height          =   1395
            Left            =   120
            TabIndex        =   17
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   5520
            TabIndex        =   15
            Top             =   1020
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   5520
            TabIndex        =   14
            Top             =   300
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   5520
            TabIndex        =   13
            Top             =   660
            Width           =   960
         End
         Begin VB.Label lblEstado 
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
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   6480
            TabIndex        =   12
            Top             =   960
            Width           =   2160
         End
         Begin VB.Label lblApertura 
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
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   6480
            TabIndex        =   11
            Top             =   240
            Width           =   2190
         End
         Begin VB.Label lblTipoCuenta 
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
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   6480
            TabIndex        =   10
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3780
         TabIndex        =   1
         Top             =   250
         Width           =   435
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   105
         TabIndex        =   0
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   4410
         TabIndex        =   16
         Top             =   210
         Width           =   3690
      End
   End
End
Attribute VB_Name = "frmCapBloqueoDesbloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto
Public bConsulta As Boolean
Public nEstado As CaptacEstado
Dim bBloqTotal As Boolean
Dim bBloqRetiro As Boolean
'By capi 21012009
Dim objPista As COMManejador.Pista

'ande 20171013 mejora en el combo para seleccionar los motivos de bloqueos
Dim nFilaNuevoBloqueo As Integer

Public nDestinoCredito As Long 'FRHU20140308
Public nMotivoCarga As Long 'FRHU20140308


Private Sub ImprimeBoletas(ByVal rsRet As ADODB.Recordset, ByVal rsTot As ADODB.Recordset)

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
    Dim lsDOI As String 'RIRO20160311
    
    lcCtaCod = txtCuenta.NroCuenta
     
    Set clsTit = New COMNCaptaGenerales.NCOMCaptaGenerales
        lsCliente = ImpreCarEsp(clsTit.GetNombreTitulares(lcCtaCod))
    
    'RIRO20160310 **************************
    lsDOI = clsTit.GetDoiCta(lcCtaCod)
    lsDOI = Left(lsDOI & Space(12), 12)
    'END RIRO ******************************
        
    Set clsTit = Nothing
    
    If Not (rsRet Is Nothing) Then
    
         If bBloqRetiro = True Then
                rsRet.MoveLast
                Do While Not rsRet.EOF
                    bEstado = IIf(rsRet("Est") = "0", False, True)
                    sFlag = rsRet("Flag")
                    sMotivo = rsRet("Motivo")
                    sComentario = Trim(rsRet("Comentario"))
                    
                    lsTit = IIf(bEstado, "Bloqueo", "Desbloqueo") & " Por " & sMotivo
                    Set oCap = New COMNCaptaGenerales.NCOMCaptaImpresion
                        'lsCadImpB = lsCadImpB & gPrnSaltoLinea & oCap.ImprimeBoleta(lsTit, sComentario, "", "", lsCliente, lcCtaCod, "", 0, 0, "", 0, 0, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt) 'RIRO 20160408 COMENTÉ LINEA
                        lsCadImpB = lsCadImpB & gPrnSaltoLinea & oCap.ImprimeBoletaNew(lsTit, sComentario, gsOpeCod, "", lsCliente, lcCtaCod, 0, False, True, , gdFecSis, gsNomAge, gsCodUser, , , , , , gbImpTMU, , , lsDOI) 'RIRO 20160408 ADD LINEA
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
    
    If Not (rsTot Is Nothing) Then
          If bBloqTotal = True Then
                    rsTot.MoveLast
                    Do While Not rsTot.EOF
                        bEstado = IIf(rsTot("Est") = "0", False, True)
                        sFlag = rsTot("Flag")
                        sMotivo = rsTot("Motivo")
                        sComentario = Trim(rsTot("Comentario"))
                        Set oCap = New COMNCaptaGenerales.NCOMCaptaImpresion
                            lsTit = IIf(bEstado, "Bloqueo", "Desbloqueo") & " Por " & sMotivo
                            'lsCadImpBT = lsCadImpBT & gPrnSaltoLinea & oCap.ImprimeBoleta(lsTit, sComentario, "", "", lsCliente, lcCtaCod, "", 0, 0, "", 0, 0, False, False, , , , True, , , , gdFecSis, gsNomAge, gsCodUser, sLpt, , False) 'RIRO 20160408 COMENTÉ LINEA
                            lsCadImpBT = lsCadImpBT & gPrnSaltoLinea & oCap.ImprimeBoletaNew(lsTit, sComentario, gsOpeCod, "", lsCliente, lcCtaCod, 0, False, True, , gdFecSis, gsNomAge, gsCodUser, , , , , , gbImpTMU, , , lsDOI) 'RIRO 20160408 ADD LINEA
                        Set oCap = Nothing
                        rsTot.MoveNext
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
                    Print #nFicSal, lsCadImpBT
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
    For i = 1 To grid.Rows - 1
        sFlag = grid.TextMatrix(i, 11) '** Juez 20120411 se cambió el nro de la fila (9>>11)
        If sFlag = "N" And i <> nFila Then
            If grid.TextMatrix(i, 3) <> "" Then
                nMotivoNuevo = CLng(Trim(Right(grid.TextMatrix(i, 4), 4))) '** Juez 20121214 se cambió el nro de la fila (3>>4)
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
If grid.TextMatrix(1, 0) <> "" Then
    For i = 1 To grid.Rows - 1
        sFlag = grid.TextMatrix(i, 11) '** Juez 20120411 se cambió el nro de la fila (9>>11)
        If grid.TextMatrix(i, 4) = "" And (sFlag = "N" Or sFlag = "A") Then '** Juez 20120411 se cambió el nro de la fila (3>>4)
            MsgBox "Motivo no válido. Seleccione un motivo válido", vbInformation, "Aviso"
            grid.row = i
            grid.Col = 3
            ValidaDatosBloqueo = False
            Exit Function
        End If
    Next i
End If
ValidaDatosBloqueo = True
End Function

Private Sub ResaltaBloqueoActivo(ByRef grid As FlexEdit)
Dim i As Integer, j As Integer
Dim nCol As Long, nFila As Long
For i = 1 To grid.Rows - 1
    If grid.TextMatrix(i, 1) = "." Then
        nFila = grid.row
        nCol = grid.Col
        grid.row = i
        For j = 1 To grid.Cols - 1
            grid.Col = j
            grid.CellBackColor = &HFFFFC0
        Next
        grid.Col = nCol
        grid.row = nFila
    End If
Next i
End Sub

Public Sub inicia(ByVal nProd As Producto, Optional bCons As Boolean = False)
nProducto = nProd
bConsulta = bCons
Select Case nProd
    Case gCapAhorros
        txtCuenta.Prod = Trim(str(gCapAhorros))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo - Ahorros"
        'By capi 21012009
        gsOpeCod = gAhoBloqDesbCuenta
        '
    
    Case gCapPlazoFijo
        txtCuenta.Prod = Trim(str(gCapPlazoFijo))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo - Plazo Fijo"
        'By capi 21012009
        gsOpeCod = gPFBloqDesbCuenta
        '
    Case gCapCTS
        txtCuenta.Prod = Trim(str(gCapCTS))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo - CTS"
        'By capi 21012009
        gsOpeCod = gCTSBloqDesbCuenta
        '
End Select

If bConsulta Then
    cmdGrabar.Visible = False
    grdRetiro.lbEditarFlex = False
    grdTotal.lbEditarFlex = False
Else
    cmdGrabar.Visible = True
    grdRetiro.lbEditarFlex = True
    grdTotal.lbEditarFlex = True
End If
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledProd = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraBloqueo.Enabled = False
fraCuenta.Enabled = True
Dim rsMotivo As New ADODB.Recordset
Dim clsGen As COMDConstSistema.DCOMGeneral 'DGeneral

Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsMotivo = clsGen.GetConstante(gCaptacMotBloqueoRet, "'" & gCapMotBlqRetCuentaInactiva & "'")
    grdRetiro.CargaCombo rsMotivo
    Set rsMotivo = clsGen.GetConstante(gCaptacMotBloqueoTot, "'" & gCapMotBlqTotCuentaInactiva & "'")
    grdTotal.CargaCombo rsMotivo
    tabBloqueo.Tab = 0
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
bBloqRetiro = False


If Not (rsCta.EOF And rsCta.BOF) Then
    grdRetiro.ColumnasAEditar = "X-1-X-X-X-X-X-X-X-9-X-X" 'ande 20171013

    nEstado = rsCta("nPrdEstado")
    lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy")
    lblEstado = UCase(rsCta("cEstado"))
    nEstado = rsCta("nPrdEstado")
    lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
    If Mid(rsCta("cCtaCod"), 9, 1) = "1" Then
        sMoneda = "NACIONAL"
        lblCuenta.ForeColor = &HC00000
    Else
        sMoneda = "EXTRANJERA"
        lblCuenta.ForeColor = &H8000&
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
            nRow = grdCliente.Rows - 1
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
        Set rsCta = clsMant.GetCapBloqueos(sCta, gCapTpoBlqRetiro, gCaptacMotBloqueoRet)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        Set grdRetiro.Recordset = rsCta
        ResaltaBloqueoActivo grdRetiro
    End If
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsCta = clsMant.GetCapBloqueos(sCta, gCapTpoBlqTotal, gCaptacMotBloqueoTot)
    Set clsMant = Nothing
    If Not (rsCta.EOF And rsCta.BOF) Then
        Set grdTotal.Recordset = rsCta
        ResaltaBloqueoActivo grdTotal
    End If
    Set rsCta = Nothing
    
    cmdCancelar.Enabled = True
    fraBloqueo.Enabled = True
    fraCuenta.Enabled = False
    cmdGrabar.Enabled = True
    cmdNuevoBlq.Enabled = True
    grdRetiro.SetFocus
    
    If (nEstado = gCapEstAnulada) Or (nEstado = gCapEstCancelada) Then
        MsgBox "La cuenta se encuentra Cancelada o Anulada.", vbInformation, "Aviso"
        cmdCancelar.Enabled = True
        fraBloqueo.Enabled = True
        fraCuenta.Enabled = False
        cmdNuevoBlq.Enabled = False
        cmdGrabar.Enabled = False
        grdRetiro.SetFocus
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
'FRHU20140301 RQ14008
Call ObtenerRelacion(txtCuenta.NroCuenta)
Dim rsMotivo As New ADODB.Recordset
Dim oGen As New COMDConstSistema.DCOMGeneral
If nEstado = gCapEstActiva Then
    If nDestinoCredito <> 14 Then '14 - Destino de Credito: Cambio de Estructura de Pasivo
        Set rsMotivo = oGen.GetConstante(gCaptacMotBloqueoRet, "'" & gCapMotBlqRetCuentaInactiva & "','" & gCapMotBlqRetCambioEstructuraPasivo & "'")
        grdRetiro.CargaCombo rsMotivo
    End If
Else
    Set rsMotivo = oGen.GetConstante(gCaptacMotBloqueoRet, "'" & gCapMotBlqRetCuentaInactiva & "'")
    grdRetiro.CargaCombo rsMotivo
End If
'FIN FRHU 20140301 RQ14008
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
grdCliente.Rows = 2
grdCliente.FormaCabecera
grdRetiro.Clear
grdTotal.Clear
grdRetiro.FormaCabecera
grdRetiro.Rows = 2
grdTotal.FormaCabecera
grdTotal.Rows = 2
lblApertura = ""
lblCuenta = ""
lblTipoCuenta = ""
lblEstado = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
txtCuenta.Cuenta = ""
txtCuenta.Age = ""
fraBloqueo.Enabled = False
fraCuenta.Enabled = True
txtCuenta.SetFocusCuenta
bBloqRetiro = False
bBloqTotal = False

End Sub

Private Sub cmdGrabar_Click()
Dim rsRet As ADODB.Recordset, rsTot As ADODB.Recordset
Dim rsRetTmp As ADODB.Recordset, rsTotTmp As ADODB.Recordset

Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim ClsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim sMovNro As String, sCuenta As String

If Not ValidaDatosBloqueo(grdRetiro) Then
    grdRetiro.SetFocus
    Exit Sub
End If
If Not ValidaDatosBloqueo(grdTotal) Then
    grdTotal.SetFocus
    Exit Sub
End If
If bBloqTotal = False And bBloqRetiro = False Then
    MsgBox "No existe Operaciones de Bloqueo Recientes", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("Está seguro de grabar??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Set rsRet = grdRetiro.GetRsNew()
    Set rsTot = grdTotal.GetRsNew()
    
    Set rsRetTmp = rsRet
    Set rsTotTmp = rsTot
    
    Set ClsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set ClsMov = Nothing
    sCuenta = txtCuenta.NroCuenta
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    'ANDE 20170811 SE AGREGÓ LA FECHA DE SISTEMA COMO PARAMETRO
    clsMant.ActualizaBloqueos sCuenta, rsRet, rsTot, sMovNro, nEstado, gdFecSis
    'END ANDE 20170811
    'By Capi 21012009
    objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , sCuenta, gCodigoCuenta
    '
    Set clsMant = Nothing
    
    ImprimeBoletas rsRetTmp, rsTotTmp
   
    cmdCancelar_Click
End If
End Sub

Private Sub cmdNuevoBlq_Click()
grdRetiro.ColumnasAEditar = "X-1-X-X-4-X-X-X-X-9-X-X" 'ande 20171013
If tabBloqueo.Tab = 0 Then
    grdRetiro.AdicionaFila , , True
    grdRetiro.TextMatrix(grdRetiro.Rows - 1, 1) = "1"
    grdRetiro.TextMatrix(grdRetiro.Rows - 1, 2) = Format$(gdFecSis, "dd/mm/yyyy")
    grdRetiro.TextMatrix(grdRetiro.Rows - 1, 3) = Format$(Time, "hh:mm:ss") '** Juez 20120411 ******
    grdRetiro.TextMatrix(grdRetiro.Rows - 1, 5) = gsCodUser '** Juez 20120411 se cambió el nro de la fila (4>>5)
    grdRetiro.TextMatrix(grdRetiro.Rows - 1, 11) = "N" '** Juez 20120411 se cambió el nro de la fila (9>>11)
    nFilaNuevoBloqueo = grdRetiro.Rows - 1  'ande 20171013
    grdRetiro.SetFocus
    
    bBloqRetiro = True
    
ElseIf tabBloqueo.Tab = 1 Then
    grdTotal.AdicionaFila , , True
    grdTotal.TextMatrix(grdTotal.Rows - 1, 1) = "1"
    grdTotal.TextMatrix(grdTotal.Rows - 1, 2) = Format$(gdFecSis, "dd/mm/yyyy")
    grdTotal.TextMatrix(grdTotal.Rows - 1, 3) = Format$(Time, "hh:mm:ss") '** Juez 20120411 ******
    grdTotal.TextMatrix(grdTotal.Rows - 1, 5) = gsCodUser '** Juez 20120411 se cambió el nro de la fila (4>>5)
    grdTotal.TextMatrix(grdTotal.Rows - 1, 11) = "N" '** Juez 20120411 se cambió el nro de la fila (9>>11)
    nFilaNuevoBloqueo = grdRetiro.Rows - 1 'ande 20171013
    grdTotal.SetFocus
    
    bBloqTotal = True
    
End If
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
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
'By Capi 20012009
Set objPista = New COMManejador.Pista
'
End Sub

Private Sub grdRetiro_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'Modif por JUEZ 20121004 segun OYP-RFC102-2012 Sólo Legal***********************************************
Dim oGen As COMDConstSistema.DCOMGeneral

Set oGen = New COMDConstSistema.DCOMGeneral
Dim rs As ADODB.Recordset
'Dim lsCodAreaBlq As String
Dim bAreaLegal As Boolean 'JUEZ 20130130
Set rs = New ADODB.Recordset
Set rs = oGen.GetDataUser(grdRetiro.TextMatrix(pnRow, 5), False)
Set oGen = Nothing
'lsCodAreaBlq = rs!cAreaCod
'FRHU20170919 INC1709150014
'bAreaLegal = IIf(rs!FechaCargo <= CDate(grdRetiro.TextMatrix(pnRow, 2)) And rs!cAreaCod = "050", True, False) 'JUEZ 20130130
If Not (rs.EOF And rs.BOF) Then
    bAreaLegal = IIf(rs!FechaCargo <= CDate(grdRetiro.TextMatrix(pnRow, 2)) And rs!cAreaCod = "050", True, False) 'JUEZ 20130130
Else
    bAreaLegal = False
End If
'FIN FRHU20170919
Set rs = Nothing
If grdRetiro.TextMatrix(pnRow, 1) = "." Then
    If grdRetiro.TextMatrix(pnRow, 11) = "A" Then '** Juez 20120411 se cambió el nro de la fila (9>>11)
        grdRetiro.TextMatrix(pnRow, 6) = "" '** Juez 20120411 se cambió el nro de la fila (5>>6)
        grdRetiro.TextMatrix(pnRow, 7) = "" '** Juez 20120411 se agregó el campo dHoraDbl
        grdRetiro.TextMatrix(pnRow, 8) = "" '** Juez 20120411 se cambió el nro de la fila (6>>8)
    Else
        grdRetiro.TextMatrix(pnRow, 1) = ""
    End If
    
    bBloqRetiro = True
Else
    'If lsCodAreaBlq = "050" Then
    If bAreaLegal Then 'JUEZ 20130130
        If gsCodArea = "050" Then
            If grdRetiro.TextMatrix(pnRow, 11) = "N" Then '** Juez 20120411 se cambió el nro de la fila (9>>11)
                grdRetiro.TextMatrix(pnRow, 1) = "1"
            Else
                If grdRetiro.TextMatrix(pnRow, 6) = "" And grdRetiro.TextMatrix(pnRow, 6) = "" Then '** Juez 20120411 se cambió el nro de la fila (5>>6)
                    grdRetiro.TextMatrix(pnRow, 6) = Format$(gdFecSis, "dd/mm/yyyy") '** Juez 20120411 se cambió el nro de la fila (5>>6)
                    grdRetiro.TextMatrix(pnRow, 7) = Format$(Time, "hh:mm:ss") '** Juez 20120411 se agregó el campo dHoraDbl
                    grdRetiro.TextMatrix(pnRow, 8) = gsCodUser '** Juez 20120411 se cambió el nro de la fila (6>>8)
                End If
            End If
            bBloqRetiro = True
        Else
            MsgBox "Ud. no puede desbloquear esta cuenta, solo un personal del Area de Legal podrá realizarlo", vbInformation, "Aviso"
            grdRetiro.TextMatrix(pnRow, 1) = "1"
            Exit Sub
        End If
    Else
        If grdRetiro.TextMatrix(pnRow, 11) = "N" Then '** Juez 20120411 se cambió el nro de la fila (9>>11)
            grdRetiro.TextMatrix(pnRow, 1) = "1"
        Else
            If grdRetiro.TextMatrix(pnRow, 6) = "" And grdRetiro.TextMatrix(pnRow, 6) = "" Then '** Juez 20120411 se cambió el nro de la fila (5>>6)
                grdRetiro.TextMatrix(pnRow, 6) = Format$(gdFecSis, "dd/mm/yyyy") '** Juez 20120411 se cambió el nro de la fila (5>>6)
                grdRetiro.TextMatrix(pnRow, 7) = Format$(Time, "hh:mm:ss") '** Juez 20120411 se agregó el campo dHoraDbl
                grdRetiro.TextMatrix(pnRow, 8) = gsCodUser '** Juez 20120411 se cambió el nro de la fila (6>>8)
            End If
        End If
        bBloqRetiro = True
    End If
End If
'END JUEZ **********************************************************************************************
End Sub

Private Sub grdRetiro_OnChangeCombo()
Dim nMotivo As Long, nFila As Long
nFila = grdRetiro.row
'FRHU 20140308 RQ14008
'ande 20171013 mejora en evitar que se cargue el combo cuendo ya se ha registrado el bloqueo
If nFila = nFilaNuevoBloqueo Then
    If nMotivoCarga = 24 Then
        If grdRetiro.TextMatrix(nFila, 4) <> "" Then
            nMotivo = CLng(Trim(Right(grdRetiro.TextMatrix(nFila, 4), 4)))
            If nMotivo <> nMotivoCarga Then
                MsgBox "Solo puede seleccionar el Motivo: Cambio de Estructura de Pasivos", vbInformation, "Aviso"
                grdRetiro.TextMatrix(nFila, 4) = ""
            End If
        End If
    End If
    'FIN FRHU 20140301 RQ14008
    If grdRetiro.TextMatrix(nFila, 4) <> "" Then '** Juez 20120411 se cambió el nro de la fila (3>>4)
        nMotivo = CLng(Trim(Right(grdRetiro.TextMatrix(nFila, 4), 4))) '** Juez 20120411 se cambió el nro de la fila (3>>4)
        If ExisteBloqueoNuevo(grdRetiro, nMotivo) Then
            MsgBox "Motivo ya seleccionado, seleccione otro motivo.", vbInformation, "Aviso"
            grdRetiro.TextMatrix(nFila, 4) = "" '** Juez 20120411 se cambió el nro de la fila (3>>4)
        End If
    End If
Else
    cmdNuevoBlq.SetFocus
    grdRetiro.SetFocus
End If
'end ande
End Sub

Private Sub grdTotal_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'Modif por JUEZ 20121004 segun OYP-RFC102-2012 Sólo Legal***********************************************
Dim oGen As COMDConstSistema.DCOMGeneral
'Dim lsCodAreaBlq As String
Dim bAreaLegal As Boolean 'JUEZ 20130130

Set oGen = New COMDConstSistema.DCOMGeneral
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = oGen.GetDataUser(grdTotal.TextMatrix(pnRow, 5), False)
Set oGen = Nothing
'lsCodAreaBlq = rs!cAreaCod
'FRHU20170919 INC1709150014
'bAreaLegal = IIf(rs!FechaCargo <= CDate(grdTotal.TextMatrix(pnRow, 2)) And rs!cAreaCod = "050", True, False) 'JUEZ 20130130
If Not (rs.EOF And rs.BOF) Then
    bAreaLegal = IIf(rs!FechaCargo <= CDate(grdTotal.TextMatrix(pnRow, 2)) And rs!cAreaCod = "050", True, False) 'JUEZ 20130130
Else
    bAreaLegal = False
End If
'FIN FRHU20170919
Set rs = Nothing

If grdTotal.TextMatrix(pnRow, 1) = "." Then
    If grdTotal.TextMatrix(pnRow, 11) = "A" Then '** Juez 20120411 se cambió el nro de la fila (9>>11)
        grdTotal.TextMatrix(pnRow, 6) = "" '** Juez 20120411 se cambió el nro de la fila (5>>6)
        grdTotal.TextMatrix(pnRow, 7) = "" '** Juez 20120411 se agregó el campo dHoraDbl
        grdTotal.TextMatrix(pnRow, 8) = "" '** Juez 20120411 se cambió el nro de la fila (6>>8)
    Else
        grdTotal.TextMatrix(pnRow, 1) = ""
    End If
    
    bBloqTotal = True
Else
    'If lsCodAreaBlq = "050" Then
    If bAreaLegal Then 'JUEZ 20130130
        If gsCodArea = "050" Then
            If grdTotal.TextMatrix(pnRow, 11) = "N" Then '** Juez 20120411 se cambió el nro de la fila (9>>11)
                grdTotal.TextMatrix(pnRow, 1) = "1"
            Else
                If grdTotal.TextMatrix(pnRow, 6) = "" And grdTotal.TextMatrix(pnRow, 6) = "" Then '** Juez 20120411 se cambió el nro de la fila (5>>6)
                    grdTotal.TextMatrix(pnRow, 6) = Format$(gdFecSis, "dd/mm/yyyy") '** Juez 20120411 se cambió el nro de la fila (5>>6)
                    grdTotal.TextMatrix(pnRow, 7) = Format$(Time, "hh:mm:ss") '** Juez 20120411 se agregó el campo dHoraDbl
                    grdTotal.TextMatrix(pnRow, 8) = gsCodUser '** Juez 20120411 se cambió el nro de la fila (6>>8)
                End If
            End If
            bBloqTotal = True
        Else
            MsgBox "Ud. no puede desbloquear este motivo, solo un personal del Area de Legal podrá realizarlo", vbInformation, "Aviso"
            grdTotal.TextMatrix(pnRow, 1) = "1"
            Exit Sub
        End If
    Else
        If grdTotal.TextMatrix(pnRow, 11) = "N" Then '** Juez 20120411 se cambió el nro de la fila (9>>11)
            grdTotal.TextMatrix(pnRow, 1) = "1"
        Else
            If grdTotal.TextMatrix(pnRow, 6) = "" And grdTotal.TextMatrix(pnRow, 6) = "" Then '** Juez 20120411 se cambió el nro de la fila (5>>6)
                grdTotal.TextMatrix(pnRow, 6) = Format$(gdFecSis, "dd/mm/yyyy") '** Juez 20120411 se cambió el nro de la fila (5>>6)
                grdTotal.TextMatrix(pnRow, 7) = Format$(Time, "hh:mm:ss") '** Juez 20120411 se agregó el campo dHoraDbl
                grdTotal.TextMatrix(pnRow, 8) = gsCodUser '** Juez 20120411 se cambió el nro de la fila (6>>8)
            End If
        End If
        bBloqTotal = True
    End If
End If
'END JUEZ **********************************************************************************************
End Sub

Private Sub grdTotal_OnChangeCombo()
Dim nMotivo As Long, nFila As Long
nFila = grdTotal.row
nMotivo = CLng(Trim(Right(grdTotal.TextMatrix(nFila, 4), 4))) '** Juez 20120411 se cambió el nro de la fila (3>>4)
If ExisteBloqueoNuevo(grdTotal, nMotivo) Then
    MsgBox "Motivo ya seleccionado, seleccione otro motivo.", vbInformation, "Aviso"
    grdTotal.TextMatrix(nFila, 4) = "" '** Juez 20120411 se cambió el nro de la fila (3>>4)
End If
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
'FRHU 20140301 RQ14008
Private Sub ObtenerRelacion(ByVal cCtaAho As String)
Dim oCap As New COMNCaptaGenerales.NCOMCaptaGenerales
Dim rs As New ADODB.Recordset
Set rs = oCap.ObtenerRelacionAhoCred(txtCuenta.NroCuenta)
If Not (rs.EOF And rs.BOF) Then
    nDestinoCredito = rs!nMotivo
    If nDestinoCredito <> 14 Then nMotivoCarga = gCapMotBlqRetCambioEstructuraPasivo
Else
    nDestinoCredito = 0
    nMotivoCarga = 0
End If
Set oCap = Nothing
End Sub
'FIN FRHU
