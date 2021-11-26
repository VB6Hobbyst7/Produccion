VERSION 5.00
Begin VB.Form frmCapRegVouDepEdi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edición de Voucher"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   Icon            =   "frmCapRegVouDepEdi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "&Detalle"
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton OptMoneda 
         Caption         =   "&Nacional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "&Extranjera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin SICMACT.FlexEdit FEVoucher 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Número-Conf-Fecha de Reg.-Cliente-Importe-Motivo-Glosa-nMovNroRVD-nId-nRealizoOperacion"
      EncabezadosAnchos=   "500-1200-500-1200-3000-1200-3500-3500-0-0-0"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-L-R-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-4-0-0-0-0-0"
      TextArray0      =   "Nro"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCapRegVouDepEdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapRegVouDepEdi
'*** Descripción : Formulario para registrar el vouchert de Depósito.
'*** Creación : ELRO el 20120703, según OYP-RFC024-2012
'********************************************************************
Option Explicit


Private Sub cargarVoucherSinOperacion()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsVouchers As ADODB.Recordset
    Set rsVouchers = New ADODB.Recordset
    Dim i As Integer
    
    Call LimpiaFlex(FEVoucher)
    
    Set rsVouchers = oNCOMCaptaGenerales.obtenerVoucherSinOperacion(Right(gsCodAge, 2), IIf(optmoneda.iTem(0), "1", "2"))
    
    If Not rsVouchers.BOF And Not rsVouchers.EOF Then
        i = 1
        FEVoucher.lbEditarFlex = True
        Do While Not rsVouchers.EOF
            FEVoucher.AdicionaFila
            FEVoucher.TextMatrix(i, 1) = rsVouchers!cNroVou
            FEVoucher.TextMatrix(i, 2) = IIf(rsVouchers!bConfirmado, "1", "0")
            FEVoucher.TextMatrix(i, 3) = rsVouchers!cFecReg
            FEVoucher.TextMatrix(i, 4) = rsVouchers!cPersNombre
            FEVoucher.TextMatrix(i, 5) = Format(rsVouchers!nMonVou, "##,###0.00")
            FEVoucher.TextMatrix(i, 6) = rsVouchers!cMotivo
            FEVoucher.TextMatrix(i, 7) = rsVouchers!cMovDesc
            FEVoucher.TextMatrix(i, 8) = rsVouchers!nMovNroRVD
            'EJVG20130911 ***
            FEVoucher.TextMatrix(i, 9) = rsVouchers!nId
            FEVoucher.TextMatrix(i, 10) = rsVouchers!nRealizoOperacion
            'END EJVG *******
            i = i + 1
            rsVouchers.MoveNext
        Loop
    End If
    
    Set rsVouchers = Nothing
    Set oNCOMCaptaGenerales = Nothing
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
'EJVG20130911 ***
Private Sub cmdDetalle_Click()
    Dim lnId As Long
    Dim foVoucher As frmCapRegVouDep_NEW
    
    If FEVoucher.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de seleccionar el Voucher primero", vbInformation, "Aviso"
        Exit Sub
    End If
    lnId = CLng(FEVoucher.TextMatrix(FEVoucher.row, 9))
    If lnId > 0 Then
        Set foVoucher = New frmCapRegVouDep_NEW
        foVoucher.Editar lnId
        cargarVoucherSinOperacion
    Else
        MsgBox "Solo se puede editar los Voucher con la Nueva Estructura", vbInformation, "Aviso"
    End If
    Set foVoucher = Nothing
End Sub
'END EJVG *******
Private Sub cmdExtornar_Click()
 
 If FEVoucher.TextMatrix(FEVoucher.row, 8) = "" Then Exit Sub
    'EJVG20130911 ***
    Dim ldFecReg As Date
    Dim lnId As Long
    Dim lnRealizoOperacion As Integer
 
    ldFecReg = CDate(FEVoucher.TextMatrix(FEVoucher.row, 3))
    lnId = CLng(FEVoucher.TextMatrix(FEVoucher.row, 9))
    lnRealizoOperacion = CInt(Trim(FEVoucher.TextMatrix(FEVoucher.row, 10)))
    
    If lnRealizoOperacion > 0 Then
        MsgBox "No se puede continuar porque ya se realizaron operaciones con el Voucher", vbInformation, "Aviso"
        Exit Sub
    End If
    If DateDiff("D", ldFecReg, gdFecSis) <> 0 Then
       MsgBox "Solo se puede extornar el mismo día de registro del Voucher", vbInformation, "Aviso"
       Exit Sub
    End If
    'END EJVG *******
 
 If MsgBox("¿Esta seguro que desea eliminar el Voucher " & FEVoucher.TextMatrix(FEVoucher.row, 1) & "?", vbYesNo + vbInformation, "Aviso") = vbYes Then
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lcMovNro As String
    Dim lnConfirmar As Long
       
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    'lnConfirmar = oNCOMCaptaGenerales.eliminarVoucherDeposito(lcMovNro, _
                                                              CLng(FEVoucher.TextMatrix(FEVoucher.row, 8)))
    lnConfirmar = oNCOMCaptaGenerales.eliminarVoucherDeposito(lcMovNro, _
                                                              CLng(FEVoucher.TextMatrix(FEVoucher.row, 8)), lnId) 'EJVG20130912
    
    If lnConfirmar > 0 Then
        MsgBox "Se realizó correctamente la eliminación.", vbInformation, "Aviso"
        FEVoucher.EliminaFila FEVoucher.row
    Else
        MsgBox "No se realizó la actualización.", vbInformation, "Aviso"
        cargarVoucherSinOperacion
    End If
    lnConfirmar = 0
    lcMovNro = ""
    Set oNCOMContFunciones = Nothing
    Set oNCOMCaptaGenerales = Nothing
    
 End If
End Sub


Private Sub FEVoucher_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
  If pnCol = 2 Then
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lcMovNro As String
    Dim lnConfirmar As Long
    Dim lnCheck As Integer
    Dim lnId As Long, lnMovNro As Long
    Dim lnRealizoOperacion As Integer
    
    lnCheck = IIf(Trim(FEVoucher.TextMatrix(pnRow, 2)) = ".", 1, 0)
    lnMovNro = CLng(FEVoucher.TextMatrix(pnRow, 8))
    lnId = CLng(Trim(FEVoucher.TextMatrix(pnRow, 9)))
    lnRealizoOperacion = CInt(Trim(FEVoucher.TextMatrix(pnRow, 10)))
    If lnRealizoOperacion > 0 Then
        MsgBox "No se puede continuar porque ya se realizaron operaciones con el Voucher", vbInformation, "Aviso"
        cargarVoucherSinOperacion 'Vuelve a mostrar los mismos datos
        Exit Sub
    End If
    
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    'lnConfirmar = oNCOMCaptaGenerales.actualizarVoucherDeposito(IIf(Trim(FEVoucher.TextMatrix(pnRow, 2)) = ".", 1, 0), _
                                                                lcMovNro, _
                                                                CLng(FEVoucher.TextMatrix(pnRow, 8)))
    lnConfirmar = oNCOMCaptaGenerales.actualizarVoucherDeposito(lnCheck, lcMovNro, lnMovNro, lnId)
    If lnConfirmar = 0 Then
        MsgBox "No se realizó la actualización.", vbInformation, "Aviso"
        cargarVoucherSinOperacion
    End If
    lnConfirmar = 0
    lcMovNro = ""
    Set oNCOMContFunciones = Nothing
    Set oNCOMCaptaGenerales = Nothing
  End If
End Sub

Private Sub Form_Activate()
    FEVoucher.SetFocus
End Sub

Private Sub Form_Load()
    cargarVoucherSinOperacion
End Sub

Private Sub OptMoneda_Click(Index As Integer)
    cargarVoucherSinOperacion
End Sub
