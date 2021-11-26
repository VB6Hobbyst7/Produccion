VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmRecEgresos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6030
   ClientLeft      =   735
   ClientTop       =   1830
   ClientWidth     =   9135
   ForeColor       =   &H00000000&
   Icon            =   "frmRecEgresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipCambio 
      Height          =   615
      Left            =   6660
      TabIndex        =   29
      Top             =   0
      Width           =   2355
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblTipCambio 
         Caption         =   "Tipo de Cambio"
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   247
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5565
      TabIndex        =   7
      Top             =   5490
      Width           =   1140
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Solicitante"
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
      Left            =   105
      TabIndex        =   27
      Top             =   660
      Width           =   8895
      Begin Sicmact.TxtBuscar txtBuscarArea 
         Height          =   330
         Left            =   1215
         TabIndex        =   2
         Top             =   240
         Width           =   1065
         _extentx        =   3069
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmRecEgresos.frx":030A
         appearance      =   1
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2355
         TabIndex        =   35
         Top             =   240
         Width           =   6315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Area Funcional"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6705
      TabIndex        =   8
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7845
      TabIndex        =   9
      Top             =   5490
      Width           =   1140
   End
   Begin VB.Frame frameRecibo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RECIBO DE A RENDIR"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   1425
      Width           =   8895
      Begin Sicmact.TxtBuscar TxtBuscarPersCod 
         Height          =   360
         Left            =   1560
         TabIndex        =   6
         Top             =   2595
         Width           =   2010
         _extentx        =   3545
         _extenty        =   423
         appearance      =   1
         appearance      =   1
         font            =   "frmRecEgresos.frx":0336
         appearance      =   1
      End
      Begin VB.TextBox txtDocNro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3750
         TabIndex        =   4
         Top             =   330
         Width           =   1605
      End
      Begin VB.TextBox txtDirPer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Tag             =   "txtDireccion"
         Top             =   3360
         Width           =   6825
      End
      Begin VB.TextBox txtNomPer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   24
         Tag             =   "txtNombre"
         Top             =   3000
         Width           =   6825
      End
      Begin VB.TextBox txtConcepto 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1725
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1620
         Width           =   6915
      End
      Begin VB.TextBox txtImpCheque 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7170
         TabIndex        =   3
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox txtLEPer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6630
         TabIndex        =   23
         Tag             =   "txtDocumento"
         Top             =   2640
         Width           =   1755
      End
      Begin VB.TextBox txtImpTexto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         TabIndex        =   15
         Top             =   1200
         Width           =   6915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   26
         Top             =   2670
         Width           =   795
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   22
         Top             =   3420
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "L.E."
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6105
         TabIndex        =   21
         Top             =   2685
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   20
         Top             =   3060
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   1275
         Left            =   180
         Top             =   2505
         Width           =   8475
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Por concepto de"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label lblImporte 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   315
         Left            =   5880
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2790
         TabIndex        =   17
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "La cantidad de"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   9
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   240
         Picture         =   "frmRecEgresos.frx":0362
         Stretch         =   -1  'True
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimiento"
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
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5760
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   4545
         TabIndex        =   0
         Top             =   195
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMovNro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1155
         TabIndex        =   34
         Top             =   187
         Width           =   2715
      End
      Begin VB.Label lblOpeCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   390
         TabIndex        =   33
         Top             =   187
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   3930
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6705
      TabIndex        =   10
      Top             =   5490
      Visible         =   0   'False
      Width           =   1140
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
      TabIndex        =   31
      Top             =   1410
      Visible         =   0   'False
      Width           =   8895
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1245
         TabIndex        =   36
         Top             =   240
         Width           =   1050
         _extentx        =   1852
         _extenty        =   609
         appearance      =   1
         appearance      =   1
         font            =   "frmRecEgresos.frx":4368
         appearance      =   1
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2310
         TabIndex        =   37
         Top             =   240
         Width           =   6360
      End
      Begin VB.Label Label2 
         Caption         =   "Area Funcional"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRecEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSalir As Boolean
Dim lMN As Boolean, sMoney As String
Dim sSql As String, rs As New ADODB.Recordset
Dim cCodUsu As String, cNomUsu As String
Dim sCodAge As String, sNomAge As String
Dim sDocNat As String, sDocTpo As String, sDocEst As String * 1
Dim lTransActiva As Boolean, lConfirma As Boolean, lArendir As Boolean
Dim sCtaOrig As String, sCtaOrigDesc As String
Dim sCtaDest As String, sCtaDestDesc As String
Dim sMovNroRef As String
Dim sObjtipo As String
Dim aObj(3, 4) As String

'*****************************  NUEVO MODELO
Dim oNContFunc As NContFunciones
'Dim oNARendir As NARendir
Dim lsDocTpo As String
Dim lsCtaContDebe As String, lsCtaContDDesc As String
Dim lsCtaContHaber As String, lsCtaContHDesc As String
Dim lnTipoARendir As ArendirTipo
Public Sub Inicio(Optional plConfirma As Boolean = False, Optional pnTipoARendir As ArendirTipo)
lConfirma = plConfirma
lnTipoARendir = pnTipoARendir
lSalir = False
Me.Show 1
End Sub
Private Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If ValFecha(txtFecha) = False Then
    ValidaInterfaz = False
    Exit Function
End If
If Len(Trim(lblMovNro)) = 0 Then
    MsgBox "Movimiento no válido ", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtFecha.SetFocus
    Exit Function
End If
If Mid(gcOpeCod, 3, 1) = gMonedaExtranjera Then
    If Val(txtTipCambio) = 0 Then
        MsgBox "Monot de Tipo de Cambio no es válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtTipCambio.SetFocus
        Exit Function
    End If
End If
If Len(Trim(txtBuscarArea)) = 0 Then
    MsgBox "Codigo de Area no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtBuscarArea.SetFocus
    Exit Function
End If
If fraCajaChica.Visible Then
    If Len(Trim(txtBuscarAreaCH)) = 0 Then
        MsgBox "Caja Chica Ingresada no válida", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtBuscarAreaCH.SetFocus
        Exit Function
    End If
End If
If frameRecibo.Enabled Then
    If Len(Trim(txtDocNro)) = 0 Then
        MsgBox "Nro de Documento no válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtDocNro.Enabled Then
            txtDocNro.SetFocus
        End If
        Exit Function
    End If
    If Val(txtImpCheque) = 0 Then
        MsgBox "Importe de a Rendir no válida", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtImpCheque.SetFocus
        Exit Function
    End If
    If Len(Trim(txtConcepto)) = 0 Then
        MsgBox "Concepto no ha se ha Ingresado o no es válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtConcepto.SetFocus
        Exit Function
    End If
    If Len(Trim(TxtBuscarPersCod)) = 0 Then
        MsgBox "Persona no se ha ingresado o código no es válido", vbInformation, "Aviso"
        ValidaInterfaz = False
        TxtBuscarPersCod.SetFocus
        Exit Function
    End If
End If

End Function
Private Sub cmdAceptar_Click()
Dim ctrControl As Control
Dim nImporte As Currency
Dim n As Integer
Dim ldFecha As Date
Dim sMsg As String
Dim sCta As String

'******************** nuevo modelo *******************************
Dim oDmov As DMov
Dim lsCodAgeArea As String
Dim lnPos As String
Dim lsCadenaPrint    As String
On Error GoTo ErrSql

If ValidaInterfaz = False Then Exit Sub
lnPos = InStr(1, txtBuscarArea.Text, gsCodCMAC)
If lnPos > 0 Then
    lsCodAgeArea = Mid(txtBuscarArea.Text, Len(gsCodCMAC) + 1, Len(txtBuscarArea.Text))
Else
    lsCodAgeArea = txtBuscarArea.Text
End If

Set oDmov = New DMov
If Not lArendir Then
   'sSql = "SELECT cValorVar FROM varsistema " _
        & "WHERE  cCodProd = 'CON' and cNomVar = 'cCCH" & txtCajaCod & "'"
   'Set rs = CargaRecord(sSql)
   'If rs.EOF Then
   '   MsgBox "Agencia no fue Habilitada como Caja Chica...!", vbCritical, "Error"
   '   Exit Sub
   'End If
   
   'GetSaldoCtaObj GetCtaObjFiltro(sCtaOrig, txtCajaCod), txtCajaCod, gdFecSis
   'If gnSaldo = 0 Then
   '   MsgBox "Caja Chica sin Saldo. Es necesario solicitar Apertura o Reembolso", vbCritical, "Error"
   '   Exit Sub
   'End If
   'If gnSaldo < Val(Format(txtImpCheque, gcFormDato)) Then
   '   MsgBox "Egreso no puede ser mayor que " & gcMN & " " & Format(gnSaldo, gcFormView), vbCritical, "Error"
   '   Exit Sub
   'End If

   'sSql = "SELECT cTipDat, cValorVar FROM VarSistema WHERE cNomVar = 'cCCH" & txtCajaCod & "'"
   'If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
   'rs.Open sSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
   'If Not rs.EOF Then
   '   If rs!cTipDat = "A" Then
   '      If gnSaldo < Val(Format(rs!cValorVar, gcFormDato)) * gnTasaCajaCh Then
   '         If MsgBox("No puede realizar Egreso porque el saldo de esta Caja Chica es menor que el permitido." & Chr(10) _
   '              & "Se recomienda realizar Rendición. ¿ Desea Confirmar Recibo de Arendir ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Error") = vbNo Then
   '            If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
   '            Exit Sub
   '         End If
   '      End If
   '   End If
   'End If
   'If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
End If
'Volvemos a generar Nro. de Mov. en caso de que el generado ya exista
lblMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
If lArendir Then
   sMsg = " ¿ Seguro de Grabar Recibo de A rendir Cuenta ? "
Else
   sMsg = " ¿ Seguro de Grabar Recibo de A rendir de Caja Chica ? "
End If
If MsgBox(sMsg, vbQuestion + vbYesNo, "Aviso de Confirmación") = vbNo Then
   Exit Sub
End If
nImporte = Val(Format(txtImpCheque, gcFormDato))
'Grabamos Mov
oDmov.InsertaMov lblMovNro, gcOpeCod, txtConcepto, gMovEstContabMovContable, gMovFlagVigente, True
If Not lConfirma Then
   'Grabamos MovCabObj según el Orden
   oDmov.InsertaMovARendir lblMovNro, lnTipoARendir, lsCodAgeArea, TxtBuscarPersCod.Text, 0, True
   oDmov.InsertaMovCont lblMovNro, CCur(txtImpCheque), "0", "0", True
Else
   'sMsg = " ¿ Seguro de aceptar egreso de Caja Chica ? "
   'sSql = "INSERT INTO  movobj VALUES ('" & Me.lblMovNro & "', '001', '1', '" & sObjtipo & "') "
   'dbCmact.Execute sSql
   'sCta = GetCtaObjFiltro(sCtaDest, sObjtipo)
  '
   'sSql = "INSERT INTO  movobj VALUES ('" & txtMovNro & "', '001', '2', '" & txtAgeCod & "')"
   'dbCmact.Execute sSql
'   sCta = sCta & GetCtaObjFiltro(sCtaDest, txtAgeCod, False)
'
'   sSql = "INSERT INTO  movobj VALUES ('" & txtMovNro & "', '001', '3', '" & txtPerCod & "')"
'   dbCmact.Execute sSql
'
'   sSql = "INSERT INTO  movcta VALUES ('" & txtMovNro & "', '001', '" & sCta & "', " & nImporte & ")"
'   dbCmact.Execute sSql
'
'   sSql = "INSERT INTO  movobj VALUES ('" & txtMovNro & "', '002', '1', '" & txtCajaCod & "')"
'   dbCmact.Execute sSql
'   sCta = GetCtaObjFiltro(sCtaOrig, txtCajaCod)
'
'   sSql = "INSERT INTO  movcta VALUES ('" & txtMovNro & "', '002', '" & sCta & "', " & nImporte * -1 & ")"
'   dbCmact.Execute sSql
'
'   sSql = "INSERT INTO  movref VALUES ('" & txtMovNro & "', '" & sMovNroRef & "')"
'   dbCmact.Execute sSql
'
'   sSql = "INSERT INTO MovCajaChica VALUES ('" & txtMovNro & "','" & txtCajaCod & "')"
'   dbCmact.Execute sSql
'
'   ldFecha = CDate(GetFechaMov(sMovNroRef, True))
End If
oDmov.Inicio gsFormatoFecha
oDmov.InsertaMovDoc lblMovNro, lsDocTpo, txtDocNro, CDate(txtFecha), True
If Mid(gcOpeCod, 3, 1) = gMonedaExtranjera Then
    oDmov.GeneraMovME lblMovNro, True
End If
If oDmov.EjecutaBatch = 1 Then Exit Sub
If lConfirma Then
    Unload Me
    Exit Sub
Else
    lsCadenaPrint = oNARendir.EmiteReciboARendir(lblMovNro, gnColPage, gcEmpresaLogo, gcEmpresa, gcEmpresaRUC)
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    oPrevio.Show lsCadenaPrint, Me.Caption, , 66
    Set oPrevio = Nothing
    If MsgBox(" ¿ Desea continuar registrando Recibos de A rendir ? ", vbQuestion + vbYesNo, "Recibo de Egresos") = vbYes Then
        lblMovNro = oNContFunc.GeneraMovNro(gdFecSis, , gsCodUser)
        txtImpCheque = "0.00": txtImpTexto = "0.00"
        txtDocNro = oNContFunc.GeneraDocNro(lsDocTpo, Mid(gcOpeCod, 3, 1), gsCodUser)
        txtBuscarArea.SetFocus
    Else
        Unload Me
    End If
End If
   
Set oDmov = Nothing
Exit Sub
ErrSql:
   MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdExaCaja_Click()
sSql = "SELECT a.cObjetoCod, a.cObjetoDesc, a.nObjetoNiv " _
     & "FROM   " & gcCentralCom & "Objeto a, varsistema b " _
     & "WHERE  b.cCodProd = 'CON' and substring(b.cNomVar,1,4) = 'cCCH' and " _
           & "       substring(b.cNomVar,5,5) = a.cObjetoCod "
Set rs = CargaRecord(sSql)
If RSVacio(rs) Then
   MsgBox "No se definieron Areas funcionales como Caja Chica", vbCritical, "Error"
   rs.Close
   Exit Sub
End If
frmDescObjeto.Inicio rs, txtCajaCod, 3, "Caja Chica"
If frmDescObjeto.lOk Then
   txtCajaCod = gaObj(0, 0, 0)
   txtCajaDes = gaObj(0, 1, 0)
   frameRecibo.Enabled = True
   txtImpCheque.SetFocus
Else
   txtCajaCod.SetFocus
End If
rs.Close
End Sub

Private Sub cmdExaminar_Click()
Dim sSqlObj As String
Dim rsObj As New ADODB.Recordset
sSqlObj = gcCentralCom & "spGetTreeObj '" & aObj(1, 0) & "', " _
        & Val(aObj(1, 1)) + Val(aObj(1, 2)) & ", '" _
        & aObj(1, 3) & "'"
Set rsObj = dbCmact.Execute(sSqlObj)
If RSVacio(rsObj) Then
   MsgBox "No se definieron Areas funcionales " & IIf(lArendir, "", "como Caja Chica"), vbCritical, "Error"
   rsObj.Close
   Exit Sub
End If
frmDescObjeto.Inicio rsObj, txtAgeCod, Val(aObj(1, 1)) + Val(aObj(1, 2))
If frmDescObjeto.lOk Then
   txtAgeCod = gaObj(0, 0, 1)
   txtAgeDesc = gaObj(0, 1, 1)
   frameRecibo.Enabled = True
   txtImpCheque.SetFocus
Else
   txtAgeCod.SetFocus
End If
rsObj.Close
End Sub

Private Sub cmdImprimir_Click()

frmPrevio.Previo rtxtAsiento, "Recibo de A rendir", False, 66
End Sub
Private Sub cmdRechazar_Click()
If txtMovNro = "" Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro de Rechazar Recibo de Egresos ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   sSql = "UPDATE Mov SET cMovEstado = '2', cMovFlag = 'X' WHERE cMovNro = '" & txtMovNro & "'"
   dbCmact.Execute sSql
   LimpiaControles
   txtDocNro.SetFocus
End If
End Sub
Private Sub LimpiaControles()
If lConfirma Then
   txtMovNro = "": txtFecha = "": txtOpeCod = ""
Else
   txtOpeCod = gcOpeCod
   lblMovNro = oNContFunc.GeneraMovNro(gdFecSis, , gsCodUser)
   txtFecha = Format(gdFecSis, "dd/mm/yyyy")
End If
txtAgeCod = "": txtAgeDesc = ""
txtDocNro = "": txtImpCheque = "0.00"
txtImpTexto = "": txtConcepto = ""
txtPerCod = "": txtLEPer = ""
txtNomPer = "": txtDirPer = ""
End Sub
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdSeekPer_Click()
Dim CadNom As String
    Call frmBuscaCli.Inicia(Me, True)
    If Len(Trim(CodGrid)) > 0 Then
       txtNomPer = ShowNombre(txtNomPer.Text, True)
       cmdAceptar.SetFocus
    Else
       txtPerCod.SetFocus
    End If
End Sub


Private Sub Form_Activate()
If lSalir Then
   Unload Me
   Exit Sub
End If
If lConfirma Then
   frameRecibo.Enabled = True
   txtImpCheque.Enabled = False
   txtConcepto.Enabled = False
Else
   cmdRechazar.Visible = False
   txtDocNro.Enabled = False
End If

End Sub
Private Sub Form_Load()
Dim sTxt As String
Dim n As Integer

Me.Caption = gcOpeDesc

CentraForm Me
lSalir = False
fraTipCambio.Visible = False
If Mid(gcOpeCod, 3, 1) = gMonedaExtranjera Then    'Identificación de Tipo de Moneda
   If gnTipCambio = 0 Then
      If Not GetTipCambio(gdFecSis) Then
         lSalir = True
         Exit Sub
      End If
   End If
   txtTipCambio = Format(gnTipCambio, gcFormView)
   fraTipCambio.Visible = True
End If

Set oNContFunc = New NContFunciones
Set oNARendir = New NARendir
' Defino el Nro de Movimiento
lblOpeCod = gcOpeCod
lblMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
txtFecha = gdFecSis

lsDocTpo = oNContFunc.EmiteDocOpe(gcOpeCod)
If lsDocTpo <> "" Then
    txtDocNro = oNContFunc.GeneraDocNro(lsDocTpo, Mid(gcOpeCod, 3, 1), gsCodUser)
End If

TxtBuscarPersCod.TipoBusqueda = BuscaPersona
txtBuscarArea.psRaiz = "UNIDADES ORGANIZACIONALES"
txtBuscarArea.rs = oNARendir.CargaAgenciasAreas(gsCodCMAC)
txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
txtBuscarAreaCH.rs = oNARendir.EmiteCajasChicas
If lnTipoARendir = gArendirTipoCajaChica Then
      frameRecibo.Caption = frameRecibo.Caption & " CAJA CHICA"
      fraCajaChica.Visible = True
      frameRecibo.Top = frameRecibo.Top + fraCajaChica.Height + 60
      cmdAceptar.Top = cmdAceptar.Top + fraCajaChica.Height + 60
      cmdSalir.Top = cmdSalir.Top + fraCajaChica.Height + 60
      cmdRechazar.Top = cmdRechazar.Top + fraCajaChica.Height + 60
      Me.Height = Me.Height + fraCajaChica.Height + 60
End If
lsCtaContDebe = oNContFunc.EmiteOpeCta(gcOpeCod, "D")
lsCtaContHaber = oNContFunc.EmiteOpeCta(gcOpeCod, "H")
lsCtaContDDesc = oNContFunc.EmiteCtaContDesc(lsCtaContDebe)
lsCtaContHDesc = oNContFunc.EmiteCtaContDesc(lsCtaContHaber)

If lConfirma Then
   txtDocNro.TabIndex = 0
   
   'lsCtaContDebe = oNContFunc.EmiteOpeCta(gcOpeCod, "D")
   'lsCtaContHaber = oNContFunc.EmiteOpeCta(gcOpeCod, "D")
   'lsCtaContDDesc = oNContFunc.EmiteCtaContDesc(lsCtaContDebe)
   'lsCtaContHDesc = oNContFunc.EmiteCtaContDesc(lsCtaContHaber)
   
   'sSql = "SELECT a.cObjetoCod, b.nObjetoNiv, a.nCtaObjNiv, a.cCtaObjFiltro, a.cCtaObjImpre " _
        & "FROM    " & gcCentralCom & "CtaObj a,  " & gcCentralCom & "Objeto b " _
        & "WHERE  a.cCtaContCod = '" & sCtaDest & "' and a.cObjetoCod = b.cObjetoCod "
  ' Set rs = CargaRecord(sSql)
  ' If rs.RecordCount <> 3 Then
  '    MsgBox "Error en asignación de Objetos a Cuenta de A rendir", vbCritical, "Error"
  '    lSalir = True
  '    Exit Sub
  ' End If
  ' N = 0
  ' Do While Not rs.EOF
  '    aObj(N, 0) = rs!cObjetoCod:  aObj(N, 1) = rs!nCtaObjNiv
  '    aObj(N, 2) = rs!nObjetoNiv:  aObj(N, 3) = rs!cCtaObjFiltro
  '    aObj(N, 4) = rs!cCtaObjImpre
  '    rs.MoveNext
   '   N = N + 1
   'Loop
   
   sSql = "SELECT cObjetoCod FROM  " & gcCentralCom & "OpeObj WHERE cOpeCod = '" & gcOpeCod & "'"
   'Set rs = CargaRecord(sSql)
   If RSVacio(rs) Then
      MsgBox "No es asignó el Objeto de Caja Chica a Operación"
      lSalir = True
      Exit Sub
   End If
   sObjtipo = rs!cObjetoCod
Else
    
    
   'sSql = "Select a.cOpeCod, b.cObjetoCod, a.nOpeObjNiv, " _
        & "b.nObjetoNiv, a.cOpeObjFiltro from  " & gcCentralCom & "OpeObj as a,  " & gcCentralCom & "Objeto as b " _
        & "where cOpeCod = '" & gcOpeCod & "' and a.cobjetocod = b.cobjetocod"
   ''Set rs = CargaRecord(sSql)
   'If RSVacio(rs) Then
   '   MsgBox "Falta difinir Objetos de Operación...!", vbCritical, "Error"
   '   lSalir = True
   '   Exit Sub
   'End If
   n = 0
 '  If rs.RecordCount > 3 Then
 '     MsgBox " Se definieron demasiados Objetos en Operación ", vbCritical, "Error"
 '     lSalir = True
 '     Exit Sub
 '  End If
 '  Do While Not rs.EOF
 '     aObj(N, 0) = rs!cObjetoCod:  aObj(N, 1) = rs!nOpeObjNiv
 '     aObj(N, 2) = rs!nObjetoNiv:  aObj(N, 3) = rs!cOpeObjFiltro
 '     rs.MoveNext
 '     N = N + 1
 '  Loop
End If


'rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub txtAgeCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidarAgencia(txtAgeCod, txtAgeDesc) Then
      frameRecibo.Enabled = True
      txtImpCheque.SetFocus
   End If
End If
End Sub

Private Sub txtAgeCod_Validate(Cancel As Boolean)
If Not ValidarAgencia(txtAgeCod.Text, txtAgeDesc) Then
   Cancel = True
Else
   frameRecibo.Enabled = True
End If
End Sub

Private Sub txtCajaCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidarAgencia(txtCajaCod, txtCajaDes) Then
      frameRecibo.Enabled = True
      txtImpCheque.SetFocus
   End If
End If
End Sub

Private Sub txtCajaCod_Validate(Cancel As Boolean)
If Not ValidarAgencia(txtCajaCod.Text, txtCajaDes) Then
   Cancel = True
Else
   frameRecibo.Enabled = True
End If
End Sub
Private Sub txtBuscarArea_EmiteDatos()
lblAreaDesc = txtBuscarArea.psDescripcion
If lblAreaDesc <> "" Then
    frameRecibo.Enabled = True
    txtImpCheque.SetFocus
    txtImpCheque.SetFocus
Else
    frameRecibo.Enabled = False
End If
End Sub
Private Sub txtBuscarAreaCH_EmiteDatos()
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
If lblCajaChicaDesc <> "" Then
   frameRecibo.Enabled = True
   txtImpCheque.SetFocus
Else
    frameRecibo.Enabled = False
End If
End Sub
Private Sub TxtBuscarPersCod_EmiteDatos()
If TxtBuscarPersCod.psDescripcion <> "" Then
    txtNomPer = TxtBuscarPersCod.psDescripcion
    txtDirPer = TxtBuscarPersCod.sPersDireccion
    txtLEPer = TxtBuscarPersCod.sPersNroDoc
    cmdAceptar.SetFocus
End If
End Sub
Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   TxtBuscarPersCod.SetFocus
End If
End Sub

Private Sub txtDocNro_GotFocus()
txtDocNro.SelStart = 0
txtDocNro.SelLength = Len(txtDocNro)
End Sub

Private Sub txtDocNro_KeyPress(KeyAscii As Integer)
Dim nPos As Integer
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   nPos = InStr(1, txtDocNro, "-")
   If nPos > 0 Then
      txtDocNro = Mid(txtDocNro, 1, nPos) & Right(String(8, "0") & Mid(txtDocNro, nPos + 1, 8), 8)
   Else
      txtDocNro = Right(String(8, "0") & txtDocNro, 8)
   End If
   If ValidaRecibo Then
      cmdAceptar.SetFocus
   Else
      txtDocNro.SelStart = 0
      txtDocNro.SelLength = Len(txtDocNro)
   End If
End If
End Sub
Private Function ValidaRecibo() As Boolean
Dim rsVer As New ADODB.Recordset
ValidaRecibo = False
sSql = "SELECT a.cMovNro, a.dDocFecha, b.cMovCabObjOrden, b.cObjetoCod, c.cObjetoDesc, 0 as nMovMonto, '' as cDirPers, '' as cMovDesc, '' as cNuDocI, '' as cMovEstado " _
     & "FROM   MovDoc a,  MovCabObj b,  " & gcCentralCom & "Objeto c " _
     & "WHERE  a.cDocTpo='" & gcDocTpo & "' and a.cDocNro = '" & txtDocNro & "' and b.cMovNro = a.cMovNro AND " _
     & "       b.cObjetoCod <> '" & sObjtipo & "' and c.cObjetoCod = b.cObjetoCod and " _
     & "       EXISTS (SELECT cMovNro from  MovCabObj d Where a.cMovNro = d.cMovNro and " _
     & "                      d.cObjetoCod = '" & sObjtipo & "') " _
     & "UNION " _
     & "SELECT a.cMovNro, a.dDocFecha, b.cMovCabObjOrden, b.cObjetoCod, c.cNomPers as cObjetoDesc, e.nMovMonto, c.cDirPers, e.cMovDesc, isnull(c.cNuDocI,'') as cNuDocI, e.cMovEstado " _
     & "FROM   MovDoc a,  MovCabObj b, " & gcCentralPers & "Persona c,  Mov e " _
     & "WHERE  e.cMovFlag <> 'X' and a.cDocTpo='" & gcDocTpo & "' and a.cDocNro = '" & txtDocNro & "' and b.cMovNro = a.cMovNro AND " _
     & "       b.cObjetoCod <> '" & sObjtipo & "' and substring(b.cObjetoCod,3,10) = c.cCodPers and " _
     & "       e.cMovNro = a.cMovNro and " _
     & "       EXISTS (SELECT cMovNro FROM  MovCabObj d Where a.cMovNro = d.cMovNro and " _
     & "                     d.cObjetoCod = '" & sObjtipo & "') ORDER BY b.cMovCabObjOrden "
   If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
   rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If rs.RecordCount <> 3 Then
   MsgBox "Recibo no existe. Por favor reintentar", vbInformation, "Aviso"
   Exit Function
End If
If rs!cMovEstado = "1" Then
   MsgBox " Recibo está ANULADO...! ", vbCritical, "Error"
   Exit Function
End If
If rs!cMovEstado = "2" Then
   MsgBox " Recibo fue RECHAZADO...! ", vbCritical, "Error"
   Exit Function
End If

sSql = "SELECT cMovNro FROM  MovRef WHERE cMovNroRef = '" & rs!cMovNro & "'"
Set rsVer = CargaRecord(sSql)
If Not RSVacio(rsVer) Then
   MsgBox "Recibo de Egreso ya fue confirmado...", vbCritical, "Error"
   Exit Function
End If
sMovNroRef = rs!cMovNro
txtAgeCod = rs!cObjetoCod
txtAgeDesc = rs!cObjetoDesc
rs.MoveNext
txtPerCod = rs!cObjetoCod
txtNomPer = rs!cObjetoDesc
txtDirPer = rs!cDirPers
txtLEPer = rs!cNudoci
txtImpCheque = Format(rs!nMovMonto, gcFormView)
txtImpTexto = ConvNumLet(rs!nMovMonto)
txtConcepto = rs!cMovdesc
rs.MoveNext
txtCajaCod = rs!cObjetoCod
txtCajaDes = rs!cObjetoDesc
ValidaRecibo = True
rs.Close
rsVer.Close
Set rs = Nothing
Set rsVer = Nothing
End Function

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = False Then Exit Sub
    lblMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    If fraTipCambio.Visible Then
        txtTipCambio.SetFocus
    Else
        txtBuscarArea.SetFocus
    End If
End If
End Sub
Private Sub txtImpCheque_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImpCheque, KeyAscii, 15, 2)
If KeyAscii = 13 Then
   txtConcepto.SetFocus
End If
End Sub
Private Sub txtImpCheque_LostFocus()
Dim nImporte As Currency
If Mid(gcOpeCod, 3, 1) = gMonedaNacional Then
   nImporte = gnArendirImporte  '* IIf(txtBuscarArea.Text = "11508", 1, 1)
Else
   nImporte = 0
End If
txtImpCheque = Format(txtImpCheque, gcFormView)
   If lnTipoARendir = gArendirTipoCajaGeneral Then
      If Val(Format(txtImpCheque, gcFormDato)) <= gnArendirImporte Then
         MsgBox "El Importe para solicitar A rendir cuenta debe ser mayor a " & Format(nImporte, gcFormView) & ". " & Chr(10) & "En caso contrario solicite A rendir de Caja Chica", vbInformation, "Error"
         txtImpCheque = "0.00"
         txtImpTexto = ""
         Me.txtImpCheque.SetFocus
         Exit Sub
      End If
   Else
      If Val(Format(txtImpCheque, gcFormDato)) > gnArendirImporte Then
         MsgBox "El Importe para solicitar a Caja Chica no puede ser mayor a " & Format(nImporte, gcFormView) & ". " & Chr(10) & "En caso contrario solicite A rendir Cuenta con Caja General", vbInformation, "Error"
         txtImpCheque = ""
         txtImpTexto = ""
         txtImpCheque.SetFocus
         Exit Sub
      End If
   End If
   txtImpTexto = ConvNumLet(Val(Format(txtImpCheque, gcFormDato)))
End Sub
Private Sub txtPerCod_KeyPress(KeyAscii As Integer)
Dim rsPer As New ADODB.Recordset
KeyAscii = intfNumEnt(KeyAscii)
If KeyAscii = 13 Then
   sSql = "Select cCodPers, cNomPers, cDirPers, isnull(cNudoci,'') as cNuDocI from " & gcCentralPers & "Persona where cCodPers = '" & txtPerCod & "'"
   Set rsPer = CargaRecord(sSql)
   If rsPer.RecordCount > 0 Then
      txtLEPer = rsPer!cNudoci
      txtNomPer = ShowNombre(rsPer!cNomPers, True)
      txtDirPer = rsPer!cDirPers
      cmdAceptar.SetFocus
   Else
      MsgBox " Persona no registrada ....! ", vbCritical, "Error"
   End If
   rsPer.Close
End If
End Sub
Private Sub txtTipCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipCambio, KeyAscii)
If KeyAscii = 13 Then
    txtBuscarArea.SetFocus
End If
End Sub
