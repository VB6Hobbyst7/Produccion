VERSION 5.00
Begin VB.Form frmMntCredSaldosAdeudo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Paquetes de Adeudados "
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10920
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
   Icon            =   "frmMntCredSaldosAdeudo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalle Adeudados"
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   10695
      Begin Sicmact.FlexEdit FePagares 
         Height          =   2775
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4895
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Documento-Moneda-Linea-Monto Requerido-Monto Adicional-Monto Colocaciones-Superavit/Deficit"
         EncabezadosAnchos=   "400-900-1500-800-1700-1700-1700-1600"
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
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.TextBox txtporcentaje 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Text            =   "20"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "Generar Informacion Formato Excel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox TxtFecha 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   4440
      TabIndex        =   7
      Top             =   2880
      Width           =   1020
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Width           =   1020
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Activar"
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   1020
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paquetes"
      Height          =   2385
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   10665
      Begin Sicmact.FlexEdit FeAdeuLineas 
         Height          =   1875
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   3307
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Paquete-Linea de Credito-Documento-Fecha Adeudado-Saldo-Justificacion-Moneda-Chk"
         EncabezadosAnchos=   "400-800-2500-1000-1400-1400-1200-800-400"
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
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8"
         ListaControles  =   "0-0-3-0-2-0-0-0-4"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-R-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-2-1-1-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8520
      TabIndex        =   1
      Top             =   2880
      Width           =   1020
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9720
      TabIndex        =   0
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Label Label5 
      Caption         =   "(%) Adicional"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8280
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Proceso"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Cambio"
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
End
Attribute VB_Name = "frmMntCredSaldosAdeudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoOperacion As Integer
Dim nFilaActual As Integer
Dim nPagActual As Integer
Public oConecta As DConecta
Private Sub txtCtaIFCod_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub cmdAceptar_Click()
Dim cCodPaq As String
Dim cCodPaqGenerado As String
cmdGenerar.Enabled = True



FeAdeuLineas.Row = nFilaActual

If nTipoOperacion = 1 Then  'Nuevo
    Dim oAdeud As New DCajaCtasIF
    
    With FeAdeuLineas
        If .TextMatrix(.Row, 2) = "" Or .TextMatrix(.Row, 3) = "" Or .TextMatrix(.Row, 4) = "" Or .TextMatrix(.Row, 5) = "" Or .TextMatrix(.Row, 6) = "" Or .TextMatrix(.Row, 7) = "" Or (.TextMatrix(.Row, 6) <> "SI" And .TextMatrix(.Row, 6) <> "NO") Or (.TextMatrix(.Row, 7) <> "MN" And .TextMatrix(.Row, 7) <> "ME") Then
            MsgBox "Faltan datos para el registro o Estan errados", vbInformation, "Mensaje"
            Call cmdCancelar_Click
            Exit Sub
        End If

        cCodPaq = .TextMatrix(.Row, 1) 'Left(.TextMatrix(.Row, 1), 5)
        If BuscarPaquete(cCodPaq) = True Then
            MsgBox "Ya existe el paquete seleccionado", vbInformation, "Mensaje"
            Call cmdCancelar_Click
            Exit Sub
        End If
        Call oAdeud.InsertaCredSaldosAdeudo(cCodPaq, .TextMatrix(.Row, 2), .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), .TextMatrix(.Row, 5), .TextMatrix(.Row, 6), .TextMatrix(.Row, 7))
    End With
    
    Set oAdeud = Nothing
End If
If nTipoOperacion = 2 Then   'Modificar
    With FeAdeuLineas
    '    If .TextMatrix(.Row, 5) = "" Then
    '        MsgBox "Debe especificar el Saldo", vbInformation, "Mensaje"
    '        Call cmdCancelar_Click
    '        Exit Sub
    '    End If
        cCodPaq = .TextMatrix(.Row, 1)
        'Call oAdeud.ModificaCredSaldosAdeudo(cCodPaq,  0, 0)
        'cCodPaqGenerado = cCodPaq
    End With
    Set oAdeud = Nothing
    
End If

'MsgBox "Se realizó la operación con éxito", vbInformation, "Mensaje"
'By capi 22112007 no es necesario
'Call CargarPaquetes(cCodPaq)

'Call CargarLineasCredito
'End By
FeAdeuLineas.lbEditarFlex = False
cmdNuevo.Enabled = True
cmdEditar.Enabled = True
cmdEliminar.Enabled = True
cmdCancelar.Enabled = True
cmdAceptar.Enabled = False
End Sub

Private Function BuscarPaquete(ByVal pcCodPaq As String) As Boolean
BuscarPaquete = False
Dim I As Integer

With FeAdeuLineas
    For I = 1 To .Rows - 2
        If pcCodPaq = Left(.TextMatrix(I, 1), 5) Then
            BuscarPaquete = True
            Exit Function
        End If
    Next I
End With

End Function

Private Sub cmdCancelar_Click()
    nTipoOperacion = -1
    Call FeAdeuLineas.BackColorRow(vbWhite)
    FeAdeuLineas.ColumnasAEditar = "X-1-X-3-4-5"
    Call CargarLineasCredito
    FeAdeuLineas.lbEditarFlex = False
    cmdCancelar.Enabled = False
    cmdAceptar.Enabled = False
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
End Sub

Sub CargarPaquetes(ByVal psCodPaquete As String)

'If nTipoOperacion = 1 Then
    Call frmMntCredSaldosAdeudoDet.Inicio(FeAdeuLineas.TextMatrix(FeAdeuLineas.Row, 2), psCodPaquete)
'Else
'    Call frmMntCredSaldosAdeudoDet.Inicio(FeAdeuLineas.TextMatrix(FeAdeuLineas.Row, 2), psCodPaquete)
'End If
'FeAdeuLineas.TextMatrix(FeAdeuLineas.Row, 4) = frmMntCredSaldosAdeudoDet.nSaldoTotal
Dim oAdeud As New DCajaCtasIF

If frmMntCredSaldosAdeudoDet.bCancel = False Then
    Call oAdeud.ModificaCredSaldosAdeudo(psCodPaquete, frmMntCredSaldosAdeudoDet.nSaldoSoles, frmMntCredSaldosAdeudoDet.nSaldoDolares)
End If
Set oAdeud = Nothing
End Sub
Private Sub CmdEditar_Click()
Dim I As Integer
Dim MatColEdit(3) As String
Dim oGen As New NConstSistemas
Dim ldFecCie As Date

If FeAdeuLineas.Row = 0 Then Exit Sub

nTipoOperacion = 2
cmdNuevo.Enabled = False
cmdEditar.Enabled = False
cmdEliminar.Enabled = False
cmdAceptar.Enabled = False
cmdCancelar.Enabled = True
cmdGenerar.Enabled = True

ldFecCie = CDate(oGen.LeeConstSistema(gConstSistCierreMesNegocio))
txtFecha.Text = ldFecCie


FeAdeuLineas.lbEditarFlex = True
nFilaActual = FeAdeuLineas.Row
'MatColEdit(0) = 1
'MatColEdit(1) = 2
'MatColEdit(2) = 3

'For i = 1 To FeAdeuLineas.Rows - 1
'If i <> nFilaActual Then
'    Call HabilitaFilaFlex(nFilaActual, MatColEdit)
'End If
'Next
End Sub

Private Sub cmdEliminar_Click()
Dim cCodPaq As String
Dim oAdeud As New DCajaCtasIF

If MsgBox("Está seguro de Eliminar el Paquete?", vbQuestion + vbYesNo, "Mensaje") = vbNo Then Exit Sub

nFilaActual = FeAdeuLineas.Row
With FeAdeuLineas
    cCodPaq = .TextMatrix(.Row, 1)
End With
    Call oAdeud.EliminaCredSaldosAdeudo(cCodPaq)
Set oAdeud = Nothing
Call CargarLineasCredito

End Sub


Private Sub cmdGenerar_Click()
    Dim oDCaja As New DCajaCtasIF
    Dim X, a As Integer
    Dim cControl As String
    Dim sSql As String
    Dim cPaquete As String
    Dim cJustificacion As String
    Dim rs As ADODB.Recordset
    Dim rs0 As ADODB.Recordset
    Dim nAccion As Integer
    Dim cmoneda As String
    Dim nRequerido As Currency
    Dim cLinea As String
    Dim ldFecha As Date
    Dim lsFecha As String
    Dim cPagare As String
        
    
    If Val(txtTipoCambio.Text) = 0 Then
        MsgBox "Ingrese un Tipo de Cambio Válido", vbExclamation, "Aviso"
        txtTipoCambio.SetFocus
        Exit Sub
    End If
    If Val(txtporcentaje.Text) = 0 Then
        MsgBox "Ingrese un Porcentaje Adicional Válido", vbExclamation, "Aviso"
        txtTipoCambio.SetFocus
        Exit Sub
    End If
    
    

    'lnTipoCambio = Val(txtTipoCambio.Text)
    'lnPorcentaje = Val(txtporcentaje.Text)
    cControl = ""
    For X = 1 To FeAdeuLineas.Rows - 1
         With FeAdeuLineas
            If .TextMatrix(X, 8) = "." Then
                If cControl = "" Then
                    cControl = Mid(.TextMatrix(X, 2), 1, 2)
                End If
                If cControl <> Mid(.TextMatrix(X, 2), 1, 2) Then
                    MsgBox "No Puede Activar mas de una Linea a la vez", vbInformation, "Aviso"
                    Exit Sub
                End If
        
            End If
        End With
    Next
    If cControl = "01" Or cControl = "02" Then
        nAccion = 1
    ElseIf cControl = "03" Then
        nAccion = 2
    ElseIf cControl = "99" Then 'Para Cyrano
        nAccion = 3
    ElseIf cControl = "04" Then
        nAccion = 4
    Else
        MsgBox "Fuente no se encuentra activa", vbInformation, "Aviso"
        Exit Sub
    End If
       
    Set oConecta = New DConecta
    oConecta.AbreConexion
    sSql = " Exec stp_sel_GeneraInformacionAdeudados '" & Format(CDate(txtFecha), "yyyymmdd") & "'," & Val(txtTipoCambio) & ",'" & cControl & "'," & nAccion
    oConecta.Ejecutar (sSql)
    
    nPagActual = 1
    For X = 1 To FeAdeuLineas.Rows - 1
        FePagares.lbEditarFlex = True
        FePagares.lbEditarFlex = True
     
        With FeAdeuLineas
        If .TextMatrix(X, 8) = "." Then
          
            FePagares.AdicionaFila
            cPaquete = .TextMatrix(X, 1)
            lsFecha = Mid(.TextMatrix(X, 4), 1, 4) & "." & Mid(.TextMatrix(X, 4), 6, 2) & "." & Mid(.TextMatrix(X, 4), 9, 2)
            'ldFecha = CDate(lsFecha)
            cJustificacion = .TextMatrix(X, 6)
            FePagares.TextMatrix(nPagActual, 1) = .TextMatrix(X, 3)
            If .TextMatrix(X, 7) = "MN" Then
                FePagares.TextMatrix(nPagActual, 2) = "NUEVOS SOLES"
                cmoneda = "MN"
            Else
                FePagares.TextMatrix(nPagActual, 2) = "US DOLARES"
                cmoneda = "ME"
            End If
            cPagare = .TextMatrix(X, 3)
            FePagares.TextMatrix(nPagActual, 3) = Mid(.TextMatrix(X, 2), 1, 4)
            FePagares.TextMatrix(nPagActual, 4) = .TextMatrix(X, 5)
            FePagares.TextMatrix(nPagActual, 5) = Round(.TextMatrix(X, 5) * Val(txtporcentaje) / 100, 2)
            cLinea = Mid(FePagares.TextMatrix(nPagActual, 3), 1, 4)
            nRequerido = .TextMatrix(X, 5) + Val(FePagares.TextMatrix(nPagActual, 5))
            If cControl = "01" Then
                sSql = " Select *,Paquete= '" & cPaquete & "',Justificacion='" & cJustificacion & "',Pagare='" & cPagare & " ' From ##TmpCofide Where  Control='NO' And Substring(CuentaCod,6,1)='2' And Clave='" & cLinea & "' And  dDesembolso>='" & lsFecha & "'"
            ElseIf cControl = "02" Then
                sSql = " Select *,Paquete= '" & cPaquete & "',Justificacion='" & cJustificacion & "',Pagare='" & cPagare & " ' From ##TmpCofide Where  Control='NO' And Clave='" & cLinea & "'And  dDesembolso>='" & lsFecha & "'"
            ElseIf cControl = "03" Then
            sSql = " Select *,Paquete= '" & cPaquete & "',Justificacion='" & cJustificacion & "',Pagare='" & cPagare & " ' From ##TmpAgrobanco Where Control='NO' And Clave='" & cLinea & "'"
            Else
            sSql = " Select *,Paquete= '" & cPaquete & "',Justificacion='" & cJustificacion & "',Pagare='" & cPagare & " ' From ##TmpFoncodes Where Control='SI' And Clave='0404'"
            End If
            Set rs = oConecta.CargaRecordSet(sSql)
            If rs.RecordCount > 1 Then
                Call SayInformacionAdeudados(rs, FePagares.TextMatrix(nPagActual, 3), Val(txtTipoCambio), cmoneda, nRequerido, CInt(Format(CDate(txtFecha), "yyyy")))
'                Exit For
            Else
                MsgBox "No Existe informacion para el presente paquete"
                Exit For
            End If
            nPagActual = nPagActual + 1
        End If
        End With
    Next
    oConecta.CierraConexion
    Set oConecta = Nothing
    
End Sub

Private Sub cmdNuevo_Click()
    nTipoOperacion = 1
    cmdEditar.Enabled = False
    cmdNuevo.Enabled = False
    cmdEliminar.Enabled = False
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
    FeAdeuLineas.lbEditarFlex = True
    FeAdeuLineas.AdicionaFila
    FeAdeuLineas.lbEditarFlex = True
    nFilaActual = FeAdeuLineas.Rows - 1
    
    Dim oAdeud As New DCajaCtasIF
    FeAdeuLineas.CargaCombo oAdeud.ObtenerLineas_Codigo

    With FeAdeuLineas
        .TextMatrix(nFilaActual, 1) = oAdeud.ObtenerCorrelativo_Paquete
        .TextMatrix(nFilaActual, 6) = "SI"
        .TextMatrix(nFilaActual, 7) = "MN"
    End With
    Set oAdeud = Nothing
End Sub
Private Sub SayInformacionAdeudados(ByVal prs As ADODB.Recordset, ByVal pcLinea As String, ByVal pnTipoCambio As Currency, ByVal pcMoneda As String, ByVal pnRequerido As Currency, Optional nPeriodo As Integer = 0)
 
    Dim fs As Scripting.FileSystemObject
 
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsArchivo2 As String
    Dim lbLibroOpen As Boolean
    Dim lsNomHoja  As String
    Dim lsMes As String
    
  
    
'    'Determinando que Archivo y hoja Excel se debe abrir de acuerdo a eleccion del usuario
'
    Select Case Mid(pcLinea, 1, 2)
        Case "01"
            lsArchivo1 = "AdeudadosCofide"
            lsNomHoja = "Cofide"
        Case "02"
            lsArchivo1 = "AdeudadosCofide"
            lsNomHoja = "Cofide"
        Case "03"
            lsArchivo1 = "AdeudadosAgrobanco"
            lsNomHoja = "Agrobanco"
        Case "04"
            lsArchivo1 = "AdeudadosFoncodes"
            lsNomHoja = "INFORME MENSUAL CARTERA"
    End Select
        

    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo1 & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    lsArchivo2 = lsArchivo1 & "_" & gsCodUser & "_" & Format$(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS")

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    Select Case Mid(pcLinea, 1, 2)
        Case "01"
            Call Say_AdeudadosCofide(prs, xlHoja1, pnTipoCambio, pcMoneda, pnRequerido)
        Case "02"
           Call Say_AdeudadosCofide(prs, xlHoja1, pnTipoCambio, pcMoneda, pnRequerido)
        Case "03"
            Call Say_AdeudadosAgrobanco(prs, xlHoja1, pnTipoCambio, pcMoneda, pnRequerido)
        Case "04"
            Call Say_AdeudadosFoncodes(prs, xlHoja1, pnTipoCambio, pcMoneda, pnRequerido, xlsLibro, nPeriodo)
    End Select
    
    

    xlHoja1.SaveAs App.path & "\Spooler\" & lsArchivo2 & ".xls"
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub
Public Sub Say_AdeudadosCofide(ByVal prs As ADODB.Recordset, ByVal xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio As Double, ByVal pcMoneda As String, ByVal pnRequerido As Currency)
Dim nFil As Integer
Dim cCuentaCod As String
Dim sSql As String
Dim nCalculado As Currency


xlHoja1.Cells(3, 11) = prs!dProceso
xlHoja1.Cells(4, 11) = prs!Pagare
xlHoja1.Cells(6, 11) = pnRequerido
nFil = 10
nCalculado = 0
Do While Not prs.EOF
    If pcMoneda = "MN" Then
        xlHoja1.Cells(5, 11) = "NUEVOS SOLES"
        If prs!cmoneda = "01" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nSaldo * pnTipoCambio, 2)
        End If
    Else
        xlHoja1.Cells(5, 11) = "US DOLARES"
        If prs!cmoneda = "02" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nSaldo / pnTipoCambio, 2)
        End If
    End If
    nFil = nFil + 1
    
    xlHoja1.Cells(nFil, 2) = prs!dProceso
    xlHoja1.Cells(nFil, 3) = prs!cIFi
    xlHoja1.Cells(nFil, 4) = prs!Paquete
    If prs!justificacion = "SI" Then
        xlHoja1.Cells(nFil, 5) = "J"
    Else
        xlHoja1.Cells(nFil, 5) = "S"
    End If
    xlHoja1.Cells(nFil, 6) = prs!dDesembolso
    xlHoja1.Cells(nFil, 7) = prs!cCliente
    xlHoja1.Cells(nFil, 8) = prs!cmoneda
    xlHoja1.Cells(nFil, 9) = pnTipoCambio
    xlHoja1.Cells(nFil, 10) = prs!nDesembolso
    xlHoja1.Cells(nFil, 11) = prs!nAFijo
    xlHoja1.Cells(nFil, 12) = prs!nKTrabajo
    xlHoja1.Cells(nFil, 13) = prs!nHipoteca
    xlHoja1.Cells(nFil, 14) = prs!cPersona
    xlHoja1.Cells(nFil, 15) = prs!cActividad
    xlHoja1.Cells(nFil, 16) = prs!cUbigeo
    xlHoja1.Cells(nFil, 17) = prs!nCuotas + prs!ngracia
    xlHoja1.Cells(nFil, 18) = prs!ngracia
    xlHoja1.Cells(nFil, 19) = prs!nCuotas
    xlHoja1.Cells(nFil, 20) = prs!nTem
    If pcMoneda = "MN" Then
        If prs!cmoneda = "01" Then
            xlHoja1.Cells(nFil, 21) = prs!nSaldo
        Else
            xlHoja1.Cells(nFil, 21) = Round(prs!nSaldo * pnTipoCambio, 2)
        End If
    Else
        If prs!cmoneda = "02" Then
            xlHoja1.Cells(nFil, 21) = prs!nSaldo
        Else
            xlHoja1.Cells(nFil, 21) = Round(prs!nSaldo / pnTipoCambio, 2)
        End If
    
    End If
    xlHoja1.Cells(nFil, 22) = prs!Pagare
    xlHoja1.Cells(nFil, 23) = prs!cLinea
    xlHoja1.Cells(nFil, 24) = prs!nCodSector
    xlHoja1.Cells(nFil, 25) = prs!nCodAct
    xlHoja1.Cells(nFil, 26) = prs!cFinan
    xlHoja1.Cells(nFil, 27) = prs!cPersCod
    xlHoja1.Cells(nFil, 28) = prs!nFrecPago
    xlHoja1.Cells(nFil, 29) = prs!cSexo
    xlHoja1.Cells(nFil, 30) = prs!cPtmoIFI
    xlHoja1.Cells(nFil, 31) = prs!cZona
    xlHoja1.Cells(nFil, 32) = prs!nDiasAtraso
    xlHoja1.Cells(nFil, 33) = prs!nCalifi
    xlHoja1.Cells(nFil, 34) = prs!cCodSbs
    cCuentaCod = prs!CuentaCod
    
    sSql = "Update ##TmpCofide Set Control='SI' Where CuentaCod='" & cCuentaCod & "'"
    oConecta.Ejecutar (sSql)
    If nCalculado > pnRequerido Then
        xlHoja1.Cells(nFil + 3, 7) = "Total Calculado Cubre lo Requerido"
        xlHoja1.Cells(nFil + 3, 21) = nCalculado
        xlHoja1.Cells(nFil + 6, 1) = "--LEYENDA--"
        xlHoja1.Cells(nFil + 7, 1) = "99 - NO DISPONIBLE PARA DATOS NUMERICOS"
        xlHoja1.Cells(nFil + 8, 1) = "ND - NO DISPONIBLE PARA DATOS NO NUMERICOS"
        Exit Do
    End If
    prs.MoveNext
Loop
If nCalculado < pnRequerido Then
   xlHoja1.Cells(nFil + 3, 7) = "¡OJO!Total Calculado no Cubre lo Requerido,Coordine con el Area de TI"
   xlHoja1.Cells(nFil + 3, 21) = nCalculado
   xlHoja1.Cells(nFil + 6, 1) = "--LEYENDA--"
   xlHoja1.Cells(nFil + 7, 1) = "99 - NO DISPONIBLE PARA DATOS NUMERICOS"
   xlHoja1.Cells(nFil + 8, 1) = "ND - NO DISPONIBLE PARA DATOS NO NUMERICOS"
End If
    
End Sub
Public Sub Say_AdeudadosAgrobanco(ByVal prs As ADODB.Recordset, ByVal xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio As Double, ByVal pcMoneda As String, ByVal pnRequerido As Currency)
Dim nFil As Integer
Dim cCuentaCod As String
Dim sSql As String
Dim nCalculado As Currency


'xlHoja1.Cells(5, 3) = Trim(plsMes)
'xlHoja1.Cells(7, 4) = Str(pnAno)
xlHoja1.Cells(5, 6) = pnTipoCambio
nFil = 8
nCalculado = 0
Do While Not prs.EOF
    If pcMoneda = "MN" Then
        If prs!cmoneda = "01" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nSaldo * pnTipoCambio, 2)
        End If
    Else
        If prs!cmoneda = "02" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nSaldo / pnTipoCambio, 2)
        End If
    End If
    nFil = nFil + 1
    
    xlHoja1.Cells(nFil, 1) = prs!dProceso
    xlHoja1.Cells(nFil, 2) = prs!prestamo
    xlHoja1.Cells(nFil, 3) = prs!prestatario
    xlHoja1.Cells(nFil, 4) = prs!Aprobacion
    xlHoja1.Cells(nFil, 5) = prs!Desembolso
    xlHoja1.Cells(nFil, 6) = prs!Linea
    xlHoja1.Cells(nFil, 7) = prs!Destino
    xlHoja1.Cells(nFil, 8) = prs!Provincia
    xlHoja1.Cells(nFil, 9) = prs!Tipo
    xlHoja1.Cells(nFil, 10) = prs!Monto
    xlHoja1.Cells(nFil, 11) = prs!Vencimiento
    xlHoja1.Cells(nFil, 12) = prs!nSaldo
    xlHoja1.Cells(nFil, 13) = prs!Actividad
    sSql = "Update ##TmpAgrobanco Set Control='SI' Where CuentaCod='" & cCuentaCod & "'"
    oConecta.Ejecutar (sSql)
    If nCalculado > pnRequerido Then
        xlHoja1.Cells(nFil + 3, 3) = "Total Calculado Cubre lo Requerido"
        xlHoja1.Cells(nFil + 3, 12) = nCalculado
        Exit Do
    End If
    prs.MoveNext
    Loop
    If nCalculado < pnRequerido Then
            'Range("C17").Select
            Selection.Font.ColorIndex = 3
            Selection.Font.Bold = True
            xlHoja1.Cells(nFil + 3, 3) = "¡OJO!Total Calculado no Cubre lo Requerido,Coordine con el Area de TI"
    End If
    xlHoja1.Cells(nFil + 6, 1) = "--LEYENDA--"
    xlHoja1.Cells(nFil + 7, 1) = "ND - NO DISPONIBLE PARA DATOS NUMERICOS"
    xlHoja1.Cells(nFil + 8, 1) = "99 - NO DISPONIBLE PARA DATOS NO NUMERICOS"
End Sub
Public Sub Say_AdeudadosCyrano(ByVal prs As ADODB.Recordset, ByVal xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio As Double, ByVal pcMoneda As String, ByVal pnRequerido As Currency)
Dim nFil As Integer
Dim cCuentaCod As String
Dim sSql As String
Dim nCalculado As Currency


'xlHoja1.Cells(5, 3) = Trim(plsMes)
'xlHoja1.Cells(7, 4) = Str(pnAno)
xlHoja1.Cells(5, 6) = pnTipoCambio
nFil = 8
nCalculado = 0
Do While Not prs.EOF
    If pcMoneda = "MN" Then
        If prs!cmoneda = "01" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nSaldo * pnTipoCambio, 2)
        End If
    Else
        If prs!cmoneda = "02" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nSaldo / pnTipoCambio, 2)
        End If
    End If
    nFil = nFil + 1
    
'    xlHoja1.Cells(nFil, 1) = prs!dProceso
'    xlHoja1.Cells(nFil, 2) = prs!prestamo
'    xlHoja1.Cells(nFil, 3) = prs!prestatario
'    xlHoja1.Cells(nFil, 4) = prs!Aprobacion
'    xlHoja1.Cells(nFil, 5) = prs!Desembolso
'    xlHoja1.Cells(nFil, 6) = prs!Linea
'    xlHoja1.Cells(nFil, 7) = prs!Destino
'    xlHoja1.Cells(nFil, 8) = prs!Provincia
'    xlHoja1.Cells(nFil, 9) = prs!Tipo
'    xlHoja1.Cells(nFil, 10) = prs!Monto
'    xlHoja1.Cells(nFil, 11) = prs!Vencimiento
'    xlHoja1.Cells(nFil, 12) = prs!nSaldo
'    xlHoja1.Cells(nFil, 13) = prs!Actividad

    sSql = "Update ##TmpFoncodes Set Control='SI' Where CuentaCod='" & cCuentaCod & "'"
    oConecta.Ejecutar (sSql)
    If nCalculado > pnRequerido Then
        xlHoja1.Cells(nFil + 3, 3) = "Total Calculado Cubre lo Requerido"
        xlHoja1.Cells(nFil + 3, 12) = nCalculado
        Exit Do
    End If
    prs.MoveNext
    Loop
    If nCalculado < pnRequerido Then
            'Range("C17").Select
            Selection.Font.ColorIndex = 3
            Selection.Font.Bold = True
            xlHoja1.Cells(nFil + 3, 3) = "¡OJO!Total Calculado no Cubre lo Requerido,Coordine con el Area de TI"
    End If
    xlHoja1.Cells(nFil + 6, 1) = "--LEYENDA--"
    xlHoja1.Cells(nFil + 7, 1) = "ND - NO DISPONIBLE PARA DATOS NUMERICOS"
    xlHoja1.Cells(nFil + 8, 1) = "99 - NO DISPONIBLE PARA DATOS NO NUMERICOS"
End Sub
'****ALPA**30/04/2008
Public Sub Say_AdeudadosFoncodes(ByVal prs As ADODB.Recordset, ByVal xlHoja1 As Excel.Worksheet, ByVal pnTipoCambio As Double, ByVal pcMoneda As String, ByVal pnRequerido As Currency, ByVal xlsLibro As Excel.Workbook, Optional nPeriodo As Integer = 0)
Dim nFil As Integer
Dim cCuentaCod As String
Dim sSql As String
Dim nCalculado As Currency
Dim bEncon As Boolean
Dim nPant As Integer
Dim nPost As Integer
Dim sPeriodo As String
Dim MatrixRota() As String
Dim MatrixTecnica() As String
Dim MatrixMeCarte() As String
Dim MatrixMeColoc() As String
Dim I As Integer
Dim J As Integer
ReDim Preserve MatrixMeColoc(10, 13)
For I = 1 To 13
    For J = 1 To 10
        MatrixMeColoc(J, I) = 0#
    Next J
Next I
I = 0
ReDim Preserve MatrixMeCarte(4, 13)
For I = 1 To 13
    MatrixMeCarte(1, I) = 0#
    MatrixMeCarte(2, I) = 0#
    MatrixMeCarte(3, I) = 0#
    MatrixMeCarte(4, I) = 0#
Next I
I = 0
ReDim Preserve MatrixRota(4, 13)
For I = 1 To 13
    MatrixRota(1, I) = 0#
    MatrixRota(2, I) = 0#
    MatrixRota(3, I) = 0#
    MatrixRota(4, I) = 0#
Next I
I = 0
ReDim Preserve MatrixTecnica(4, 13)
For I = 1 To 13
    MatrixTecnica(0, I) = 0#
    MatrixTecnica(1, I) = 0#
    MatrixTecnica(2, I) = 0#
    MatrixTecnica(3, I) = 0#
    MatrixTecnica(4, I) = 0#
Next I
nPost = 0
nPant = 0

'xlHoja1.Cells(5, 3) = Trim(plsMes)
'xlHoja1.Cells(7, 4) = Str(pnAno)
xlHoja1.Cells(5, 6) = pnTipoCambio
nFil = 8
nCalculado = 0
Do While Not prs.EOF
    If pcMoneda = "MN" Then
        If prs!cmoneda = "01" Then
            nCalculado = nCalculado + prs!nMonto
        Else
            nCalculado = Round(nCalculado + prs!nMonto * pnTipoCambio, 2)
        End If
    Else
        If prs!cmoneda = "02" Then
            nCalculado = nCalculado + prs!nSaldo
        Else
            nCalculado = Round(nCalculado + prs!nMonto / pnTipoCambio, 2)
        End If
    End If
    nFil = nFil + 1
   ' If prs!cGrupo = "REPO_FLUJO_ROTATORIO" Then
    '**ALPA**2008/04/30**************************************************************************************************
        bEncon = False
        For Each xlHoja1 In xlsLibro.Worksheets
            If UCase(xlHoja1.Name) = prs!cHoja Then
                bEncon = True
                xlHoja1.Activate
                Exit For
            End If
        Next
        If bEncon = False Then
           'ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
           MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
           Exit Sub
        End If
            If prs!cHoja = "INFORME MENSUAL DE COLOCACIONES" Then
                sPeriodo = prs!cPeriodo
                If Len(Trim(prs!cPeriodo)) = 6 Then
                    sPeriodo = "0" + Trim(prs!cPeriodo)
                ElseIf Len(Trim(prs!cPeriodo)) = 4 Then
                    sPeriodo = "00-" + Trim(prs!cPeriodo)
                End If
                nPost = 3 + CInt(Left(Trim(sPeriodo), 2))
               ' If Left(prs!cPeriodo, 2) <> "00" Then
                    If prs!cDescripcion = "FR_01_SP" Then
                     '       nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(17, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(16, nPost) = prs!nCuantos
                    End If
                    If prs!cDescripcion = "FR_02_CA" Then
                      '      nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(20, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(19, nPost) = prs!nCuantos
                    End If
                    If prs!cDescripcion = "FR_03_CD" Then
                       '     nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(23, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(22, nPost) = prs!nCuantos
                            xlHoja1.Cells(37, nPost) = prs!nCuantos
                            xlHoja1.Cells(42, nPost) = prs!nCuantos
                            xlHoja1.Cells(58, nPost) = prs!nCuantos
                            xlHoja1.Cells(25, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00") 'Nuevo
                            xlHoja1.Cells(31, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00") 'Nuevo
                            xlHoja1.Cells(33, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00") 'Nuevo
                            xlHoja1.Cells(47, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00") 'Nuevo
                            xlHoja1.Cells(53, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00") 'Nuevo
                    End If
                    If prs!cDescripcion = "FR_04_TIMP" Then
                        '    nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(24, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                    End If
                    If prs!cDescripcion = "FR_05_PL" Then
                        If prs!cGrupo = "12-24 MESES" Then
                         '   nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(29, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        End If
                        If prs!cGrupo = "6-12 MESES" Then
                          '  nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(28, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        End If
                        If prs!cGrupo = "HASTA 6 MESES" Then
                           ' nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(27, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        End If
                        If prs!cGrupo = "MAS DE 24" Then
                           ' nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(30, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        End If
                    End If
                    If prs!cDescripcion = "FR_06_CS" Then
                         If prs!cGrupo = "Hombres" Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(34, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(38, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "Mujeres" Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(35, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(39, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "Sexo_ND" Then
                         xlHoja1.Cells(36, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                         xlHoja1.Cells(40, nPost) = prs!nCuantos
                         End If
                    End If
                    'Falta
                    'Aqui
                    If prs!cDescripcion = "FR_08_SE" Then
                         If prs!cGrupo = "COMERCIO" Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(50, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(45, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "PRODUCCION" Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(48, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(43, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "SERVICIOS" Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(49, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(44, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "OTROS SECTORES" Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(51, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(46, nPost) = prs!nCuantos
                         End If
                    End If
                    If prs!cDescripcion = "FR_10_PL" Then
                         If prs!cGrupo = "DE 2001 A 5000 NS/." Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(55, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(60, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "DE 5001 A 10000 NS/." Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(56, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(61, nPost) = prs!nCuantos
                         End If
                         If prs!cGrupo = "HASTA 2000 NS/." Then
                            'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                            xlHoja1.Cells(54, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                            xlHoja1.Cells(59, nPost) = prs!nCuantos
                         End If
                        If prs!cGrupo = "MAS DE 10000 NS/." Then
                          'nPost = 2 + CInt(Left(prs!cPeriodo, 2))
                          xlHoja1.Cells(57, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                          xlHoja1.Cells(62, nPost) = prs!nCuantos
                         End If
            End If
            ElseIf prs!cHoja = "INFORME MENSUAL CARTERA" Then
                sPeriodo = prs!cPeriodo
                
                If Len(Trim(prs!cPeriodo)) = 6 Then
                    sPeriodo = "0" + Trim(prs!cPeriodo)
                End If
                If nPeriodo - 1 = Mid(Trim(prs!cPeriodo), 4, 4) Then
                'If Len(Trim(prs!cPeriodo)) = 4 Then
                    sPeriodo = "00-" + Mid(Trim(prs!cPeriodo), 4, 4)
                'End If
                End If
                nPost = 3 + CInt(Left(Trim(sPeriodo), 2))
                If prs!cGrupo = "" Then
                    If prs!cDescripcion = "C_FORMATO IMCO 005_x.9" Then
                    If nPost > 3 Then
                        xlHoja1.Cells(25, nPost - 1) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(1, nPost - 2) = prs!nMonto + MatrixMeCarte(1, nPost - 2)
                        'xlHoja1.Cells(16, nPost - 1) = MatrixMeCarte(1, nPost - 2)
                    End If
                    End If
                    If prs!cDescripcion = "D_FORMATO IMCO 005_2.9,2.8" Then
                        xlHoja1.Cells(34, nPost) = prs!nCuantos
                        xlHoja1.Cells(35, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(2, nPost - 2) = prs!nMonto + MatrixMeCarte(2, nPost - 2)
                        'xlHoja1.Cells(26, nPost) = MatrixMeCarte(2, nPost - 2)
                        'xlHoja1.Cells(27, nPost) = MatrixMeCarte(2, nPost - 1)
                    End If
                End If
                 If prs!cGrupo = "Cartera Activa" Then
                    If prs!cDescripcion = "A_FORMATO IMCO 005_x.1,x.2,x.7,x.8,x.9,x.10" Then
                        xlHoja1.Cells(32, nPost) = prs!nCuantos
                        xlHoja1.Cells(31, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(2, nPost - 2) = prs!nMonto + MatrixMeCarte(2, nPost - 2)
                        xlHoja1.Cells(26, nPost) = MatrixMeCarte(2, nPost - 2)
                        xlHoja1.Cells(27, nPost + 1) = MatrixMeCarte(2, nPost - 2)
                    End If
                    If prs!cDescripcion = "B_FORMATO IMCO 005_x.3,x.4" Then
                        xlHoja1.Cells(30, nPost) = prs!nCuantos
                        xlHoja1.Cells(29, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(2, nPost - 2) = prs!nMonto + MatrixMeCarte(2, nPost - 2)
                        'xlHoja1.Cells(26, nPost) = MatrixMeCarte(2, nPost - 3)
                        'xlHoja1.Cells(27, nPost) = MatrixMeCarte(2, nPost - 2)
                    End If
                End If
                If prs!cGrupo = "Cartera Contaminada" Then
                    If prs!cDescripcion = "A_FORMATO IMCO 005_x.1,x.2,x.7,x.8,x.9,x.10" Then
'                        xlHoja1.Cells(40, nPost) = prs!nCuantos
                        xlHoja1.Cells(41, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(3, nPost - 2) = prs!nMonto + MatrixMeCarte(3, nPost - 2)
                        'xlHoja1.Cells(37, nPost) = MatrixMeCarte(3, nPost - 2)
                        'xlHoja1.Cells(38, nPost) = MatrixMeCarte(3, nPost - 2)
                    End If
                End If
                 If prs!cGrupo = "Cartera Morosa" Then
                    If prs!cDescripcion = "A_FORMATO IMCO 005_x.1,x.2,x.7,x.8,x.9,x.10" Then
                        xlHoja1.Cells(40, nPost) = prs!nCuantos
                        xlHoja1.Cells(39, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(3, nPost - 2) = prs!nMonto + MatrixMeCarte(3, nPost - 2)
                        xlHoja1.Cells(37, nPost) = MatrixMeCarte(3, nPost - 2)
                        xlHoja1.Cells(38, nPost + 1) = MatrixMeCarte(3, nPost - 2)
                    End If
                End If
                If prs!cGrupo = "Cartera Refinanciada" Then
                    If prs!cDescripcion = "B_FORMATO IMCO 005_x.3,x.4" Then
                        xlHoja1.Cells(45, nPost) = prs!nCuantos
                        xlHoja1.Cells(46, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(4, nPost - 2) = prs!nMonto + MatrixMeCarte(4, nPost - 2)
                        'xlHoja1.Cells(43, nPost) = MatrixMeCarte(4, nPost - 2)
                        'xlHoja1.Cells(44, nPost + 1) = MatrixMeCarte(4, nPost - 2)
                    End If
                     If prs!cDescripcion = "B_FORMATO IMCO 005_x.3,x.4" Then
                        xlHoja1.Cells(47, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(4, nPost - 2) = prs!nMonto + MatrixMeCarte(4, nPost - 2)
                        'xlHoja1.Cells(43, nPost) = MatrixMeCarte(4, nPost - 2)
                        'xlHoja1.Cells(44, nPost + 1) = MatrixMeCarte(4, nPost - 2)
                    End If
                    If prs!cDescripcion = "A_FORMATO IMCO 005_x.1,x.2,x.7,x.8,x.9,x.10" Then
                        xlHoja1.Cells(49, nPost) = prs!nCuantos
                        xlHoja1.Cells(50, nPost) = Format(Round(prs!nMonto, 2), "#,###,###,##0.00")
                        MatrixMeCarte(4, nPost - 2) = prs!nMonto + MatrixMeCarte(4, nPost - 2)
                        xlHoja1.Cells(43, nPost) = MatrixMeCarte(4, nPost - 2)
                        xlHoja1.Cells(44, nPost + 1) = MatrixMeCarte(4, nPost - 2)
                    End If
                End If
                If nPost >= 4 Then
                xlHoja1.Cells(16, nPost - 2) = xlHoja1.Cells(26, nPost - 2) + xlHoja1.Cells(37, nPost - 2) + xlHoja1.Cells(43, nPost - 2)
                End If
                xlHoja1.Range("c16") = xlHoja1.Range("c26") + xlHoja1.Range("c37") + xlHoja1.Range("c43")
            End If
    '**END***************************************************************************************************************
   ' End If
'    xlHoja1.Cells(nFil, 1) = prs!dProceso
'    xlHoja1.Cells(nFil, 2) = prs!prestamo
'    xlHoja1.Cells(nFil, 3) = prs!prestatario
'    xlHoja1.Cells(nFil, 4) = prs!Aprobacion
'    xlHoja1.Cells(nFil, 5) = prs!Desembolso
'    xlHoja1.Cells(nFil, 6) = prs!Linea
'    xlHoja1.Cells(nFil, 7) = prs!Destino
'    xlHoja1.Cells(nFil, 8) = prs!Provincia
'    xlHoja1.Cells(nFil, 9) = prs!Tipo
'    xlHoja1.Cells(nFil, 10) = prs!Monto
'    xlHoja1.Cells(nFil, 11) = prs!Vencimiento
'    xlHoja1.Cells(nFil, 12) = prs!nSaldo
'    xlHoja1.Cells(nFil, 13) = prs!Actividad
'    sSql = "Update ##TmpFoncodes Set Control='SI' Where CuentaCod='" & cCuentaCod & "'"
'    oConecta.Ejecutar (sSql)
'    If nCalculado > pnRequerido Then
'        xlHoja1.Cells(nFil + 3, 3) = "Total Calculado Cubre lo Requerido"
'        xlHoja1.Cells(nFil + 3, 12) = nCalculado
'        Exit Do
'    End If
    prs.MoveNext
    Loop
    If nCalculado < pnRequerido Then
            'Range("C17").Select
            Selection.Font.ColorIndex = 3
            Selection.Font.Bold = True
            xlHoja1.Cells(nFil + 3, 3) = "¡OJO!Total Calculado no Cubre lo Requerido,Coordine con el Area de TI"
    End If
    xlHoja1.Cells(nFil + 6, 1) = "--LEYENDA--"
    xlHoja1.Cells(nFil + 7, 1) = "ND - NO DISPONIBLE PARA DATOS NUMERICOS"
    xlHoja1.Cells(nFil + 8, 1) = "99 - NO DISPONIBLE PARA DATOS NO NUMERICOS"
    sSql = "drop table ##TmpFoncodes"
    oConecta.Ejecutar (sSql)
End Sub
'****End*ALPA********



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

'Private Sub FeAdeuLineas_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
'Dim sPersCod As String
'Dim sIFTpo As String
'Dim sCtaIFCod As String
'Dim oCtaIf As DCajaCtasIF
'Dim rs As ADODB.Recordset
'
'If psDataCod = "" Then Exit Sub
'
''sPersCod = Mid(psDataCod, 4, 13)
''sIFTpo = Mid(psDataCod, 1, 2)
''sCtaIFCod = Mid(psDataCod, 18, 10)
'
''Set oCtaIf = New DCajaCtasIF
''Set rs = oCtaIf.GetSaldoCtaIFAdeudado(sPersCod, sIFTpo, sCtaIFCod)
''If Not rs.EOF Then
''    FeAdeuLineas.TextMatrix(pnRow, 5) = rs!nSaldoCap
''End If
'
'Set oCtaIf = Nothing
'End Sub

Private Sub FeAdeuLineas_RowColChange()
'If nTipoOperacion = 2 Then
'    FeAdeuLineas.Row = nFilaActual
'End If

If FeAdeuLineas.Col = 2 Then
    Dim oAdeud As New DCajaCtasIF
    FeAdeuLineas.rsTextBuscar = oAdeud.GetLineaCredito()
    Set oAdeud = Nothing
End If

'If FeAdeuLineas.Col = 3 Then
'    Dim oOpe As New DOperacion
'    FeAdeuLineas.rsTextBuscar = oOpe.GetRsOpeObj("401832", "0")
'    Set oOpe = Nothing
'End If
End Sub

Private Sub FELineas_RowColChange()
If FeAdeuLineas.Col = 1 Then
    Dim oAdeud As New DCajaCtasIF
    FeAdeuLineas.rsTextBuscar = oAdeud.GetLineaCredito()
    Set oAdeud = Nothing
End If
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    Call CargarLineasCredito
    nTipoOperacion = -1
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    FeAdeuLineas.lbEditarFlex = False
End Sub

Private Sub CargarLineasCredito()
'Dim rs As ADODB.Recordset
    Dim oAdeud As New DCajaCtasIF
    
    FeAdeuLineas.Clear
    FeAdeuLineas.FormaCabecera
    FeAdeuLineas.Rows = 2
    FeAdeuLineas.rsFlex = oAdeud.GetCredSaldosAdeudo()
    Set oAdeud = Nothing
End Sub
Private Sub CargarAdeudadosDetalle(pnTipoCambio)
    Dim oAdeud As New DCajaCtasIF
    FePagares.Clear
    FePagares.FormaCabecera
    FePagares.Rows = 2
    FePagares.rsFlex = oAdeud.GetInformacionAdeudados(pnTipoCambio, 1)
    Set oAdeud = Nothing

End Sub

Private Sub HabilitaFilaFlex(ByRef pnFilaAct As Integer, ByVal ColBloq As Variant, Optional ByVal SelecColorFlex As Long = vbYellow)
Dim I As Integer
Dim J As Integer
    
    pnFilaAct = FeAdeuLineas.Row
    Call FeAdeuLineas.BackColorRow(SelecColorFlex)
    For I = 0 To UBound(ColBloq) - 1
        FeAdeuLineas.ColumnasAEditar = Replace(FeAdeuLineas.ColumnasAEditar, Trim(Str(ColBloq(I))), "X")
    Next I
End Sub


