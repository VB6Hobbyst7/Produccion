VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdeudCalMnt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  CALENDARIO DE ADEUDADOS"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmAdeudCalMnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "4. &Grabar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   7410
      TabIndex        =   5
      Top             =   7650
      Width           =   1320
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "3. &Actualizar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5985
      TabIndex        =   4
      Top             =   7650
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Height          =   6435
      Left            =   75
      TabIndex        =   11
      Top             =   1155
      Width           =   10275
      Begin VB.TextBox txtcMovNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3075
         Visible         =   0   'False
         Width           =   2430
      End
      Begin MSComctlLib.ListView lstCabecera 
         Height          =   2730
         Left            =   75
         TabIndex        =   0
         Top             =   315
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   4815
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo Cuota"
            Object.Width           =   2090
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cuota"
            Object.Width           =   1085
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Capital"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Interés"
            Object.Width           =   1402
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Comisión"
            Object.Width           =   1429
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "cMovNro"
            Object.Width           =   4577
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cTpoCuota"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "cEstado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "nDiasPago"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "nInteresPagado"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstDetalle 
         Height          =   2730
         Left            =   60
         TabIndex        =   1
         Top             =   3390
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   4815
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo Cuota"
            Object.Width           =   2090
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cuota"
            Object.Width           =   1085
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Capital"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Interés"
            Object.Width           =   1402
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Comisión"
            Object.Width           =   1429
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Saldo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "cMovNro"
            Object.Width           =   4577
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cTpoCuota"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "cEstado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "nDiasPago"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "nInteresPagado"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   135
         Left            =   1620
         TabIndex        =   17
         Top             =   3225
         Visible         =   0   'False
         Width           =   8550
         _ExtentX        =   15081
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calendario EXCEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   75
         TabIndex        =   15
         Top             =   3165
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calendario SICMACT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   135
         Width           =   1710
      End
      Begin VB.Label lblImportar 
         AutoSize        =   -1  'True
         Caption         =   "Importando datos... Por favor espere un momento ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   6135
         Width           =   4440
      End
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "1. &Cargar Excel"
      Height          =   330
      Left            =   3120
      TabIndex        =   2
      Top             =   7650
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adeudado"
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
      Height          =   1050
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   8775
      Begin VB.Label lblDescEntidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   915
         TabIndex        =   10
         Top             =   630
         Width           =   4695
      End
      Begin VB.Label lblEntidad 
         AutoSize        =   -1  'True
         Caption         =   "Entidad :"
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
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   330
         Width           =   780
      End
      Begin VB.Label lblCodAdeudado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   915
         TabIndex        =   8
         Top             =   285
         Width           =   3765
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   9045
      TabIndex        =   6
      Top             =   7650
      Width           =   1230
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "2. &Validar Datos"
      Height          =   330
      Left            =   4545
      TabIndex        =   3
      Top             =   7650
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Siga la secuencia para importar y/o actualizar un calendario                ==>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   150
      TabIndex        =   16
      Top             =   7620
      Width           =   2805
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   9285
      Picture         =   "frmAdeudCalMnt.frx":08CA
      Stretch         =   -1  'True
      Top             =   210
      Width           =   675
   End
End
Attribute VB_Name = "frmAdeudCalMnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nBandera As Integer
Dim nCuotaPagada As Integer
Dim nCuotaPagadaXLS As Integer

Private Sub cmdActualizar_Click()
Dim i As Integer
Dim j As Integer

'Dim nArreglo() As Integer      ' COMENTADO POR ANGC 20210205
'Dim nCant As Integer
'
'Dim nCantTemp As Integer
'Dim nCAntTemp1 As Integer
'
'
'nCantTemp = 0
'nCAntTemp1 = 0
'For i = 1 To lstDetalle.ListItems.Count
'    If lstDetalle.ListItems(i).SubItems(9) = "6" Then
'        nCantTemp = nCantTemp + 1
'    End If
'Next
'
'If nCantTemp > 0 Then
'    For i = 1 To lstCabecera.ListItems.Count
'        If lstCabecera.ListItems(i).SubItems(9) = "6" Then
'            nCAntTemp1 = nCAntTemp1 + 1
'        End If
'    Next
'End If
'
'
'If lstCabecera.ListItems.Count > 0 Then
'    For i = 1 To lstCabecera.ListItems.Count
'        For j = 1 To lstDetalle.ListItems.Count
'            If lstCabecera.ListItems(i).SubItems(2) = lstDetalle.ListItems(j).SubItems(2) Then 'nNroCuota
'                If lstCabecera.ListItems(i).SubItems(9) = lstDetalle.ListItems(j).SubItems(9) Then 'Tipo de Cuota
'                    nCant = nCant + 1
'                    ReDim Preserve nArreglo(nCant) As Integer
'                    nArreglo(nCant) = i
'                    Exit For
'                End If
'            End If
'        Next
'    Next
'
'    If nCantTemp > 0 And nCAntTemp1 = 0 Then
'        For i = nCantTemp + 1 To lstDetalle.ListItems.Count
'            For j = 1 To 12
'                If j = 8 Or j = 10 Or j = 12 Then
'
'                    '1=Fecha NO
'                    '2=NroCuota NO
'                    '3=nCapital NO
'                    '4=nInteres NO
'                    '5=Comision NO
'                    '6=Total NO
'                    '7=Saldo Capital NO
'                    '8 = cMovNro SI
'                    '9 =cTpoCuota NO
'                    '10 = cEstado SI
'                    '11 = nDiasPago NO
'                    '12 = nInteresPagado SI
'
'
'                    lstDetalle.ListItems(i).SubItems(j) = lstCabecera.ListItems(i - nCantTemp).SubItems(j)
'                    If j = 10 Then
'                        If Trim(lstCabecera.ListItems(i - nCantTemp).SubItems(10)) = "1" Then
'                            lstDetalle.ListItems.Item(i).ListSubItems.Item(2).Bold = True
'                            lstDetalle.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbRed
'                        End If
'                    ElseIf j = 8 Then
'                        If Len(Trim(lstCabecera.ListItems(i - nCantTemp).SubItems(8))) > 0 Then
'                            lstDetalle.ListItems.Item(i).ListSubItems.Item(8).ForeColor = vbBlue
'                        End If
'                    End If
'                End If
'            Next
'        Next
'    Else
'
'        For i = 1 To lstDetalle.ListItems.Count
'            For j = 1 To 12
'                If j = 8 Or j = 10 Or j = 12 Then
'
'                    '1=Fecha NO
'                    '2=NroCuota NO
'                    '3=nCapital NO
'                    '4=nInteres NO
'                    '5=Comision NO
'                    '6=Total NO
'                    '7=Saldo Capital NO
'                    '8 = cMovNro SI
'                    '9 =cTpoCuota NO
'                    '10 = cEstado SI
'                    '11 = nDiasPago NO
'                    '12 = nInteresPagado SI
'
'
'                    lstDetalle.ListItems(i).SubItems(j) = lstCabecera.ListItems(i).SubItems(j)
'                    If j = 10 Then
'                        If Trim(lstCabecera.ListItems(i).SubItems(10)) = "1" Then
'                            lstDetalle.ListItems.Item(i).ListSubItems.Item(2).Bold = True
'                            lstDetalle.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbRed
'                        End If
'                    ElseIf j = 8 Then
'                        If Len(Trim(lstCabecera.ListItems(i).SubItems(8))) > 0 Then
'                            lstDetalle.ListItems.Item(i).ListSubItems.Item(8).ForeColor = vbBlue
'                        End If
'                    End If
'                End If
'            Next
'        Next
'    End If
'End If
lstDetalle.BackColor = &H80000018
cmdGrabar.Enabled = True
cmdGrabar.SetFocus
End Sub

 

Private Sub cmdValidar_Click()
Dim nCant1 As Integer
Dim ncant2 As Integer
Dim nCantA1 As Integer
Dim nCantA2 As Integer
Dim i As Integer
Dim nSaldoSist As Double
Dim nSaldoExcel As Double
Dim sMensaje As String
Dim nPagados As Integer
Dim nProvisionados As Integer

If gsOpeCod = "401803" Or gsOpeCod = "402803" Then 'Nuevo
    If lstCabecera.ListItems.Count > 0 Then
        MsgBox "Existen datos en el calendario del sistema" & Chr(10) & "Verifique", vbInformation, "Aviso"
        Exit Sub
    Else
        cmdActualizar.Enabled = True
        cmdActualizar.SetFocus
        Exit Sub
    End If
End If

nPagados = 0
nProvisionados = 0
nCant1 = 0
ncant2 = 0
cmdActualizar.Enabled = False
cmdGrabar.Enabled = False
'nSaldoSist = 0      'ANGC 20210205
'nSaldoExcel = 0

For i = 1 To lstCabecera.ListItems.Count
    If lstCabecera.ListItems(i).SubItems(9) = "2" Then
        nCant1 = nCant1 + 1
    ElseIf lstCabecera.ListItems(i).SubItems(9) = "6" Then
        ncant2 = ncant2 + 1
    End If
    
'    If lstCabecera.ListItems(i).SubItems(10) <> 1 Then
'        nSaldoSist = nSaldoSist + nSaldoSist
'    End If
Next

For i = 1 To lstDetalle.ListItems.Count
    If lstDetalle.ListItems(i).SubItems(9) = "2" Then
        nCantA1 = nCantA1 + 1
    ElseIf lstDetalle.ListItems(i).SubItems(9) = "6" Then
        nCantA2 = nCantA2 + 1
    End If
    
'    If lstDetalle.ListItems(i).SubItems(10) <> 1 Then
'        nSaldoExcel = nSaldoExcel + nSaldoExcel
'    End If
Next

If Len(Trim(sMensaje)) > 0 Then
    MsgBox "Existen estas observaciones " & Chr(10) & Chr(10) & sMensaje, vbInformation, "Aviso"
    cmdActualizar.Enabled = False
    cmdGrabar.Enabled = False
    Exit Sub
End If

For i = 1 To nCant1
    If lstCabecera.ListItems(i).SubItems(10) = "1" Then  'pagado
        nPagados = nPagados + 1
    Else
        If Val(lstCabecera.ListItems(i).SubItems(12)) > 0 Then 'Interes Pagado
            nProvisionados = nProvisionados + 1
        End If
    End If
Next

If nPagados > 0 Then        'ANGC CUOTAS PAGADAS NO SERAN MODIFICADAS
    sMensaje = sMensaje & "Existen " & nPagados & " Cuotas Pagadas que no podrán ser modificadas" & Chr(10)
End If
If nProvisionados > 0 Then
    sMensaje = sMensaje & "Existen " & nProvisionados & " Cuotas Provisionadas " & Chr(10)
End If

If Len(Trim(sMensaje)) > 0 Then
    MsgBox "Existen estas observaciones " & Chr(10) & Chr(10) & sMensaje, vbInformation, "Aviso"
End If
    cmdActualizar.Enabled = True
    cmdActualizar.SetFocus
 
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
 
Private Sub cmdImportar_Click()

Dim oArch As Object
Dim lsArchivo As String
Dim i As Integer, nCol As Integer
Dim lbExcel As Boolean

Dim oCon As DConecta

Dim pcIFTpo As String
Dim pcPersCod As String
Dim pcCtaIFCod As String

Dim sSql As String
Dim reg As New ADODB.Recordset
Dim nCuoSig As Integer
Dim nTipCuoSig As Integer
Dim nCuoSigNC As Integer
Dim nTipCuoSigNC As Integer

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Dim mySheet As Excel.Worksheet


Dim L As ListItem

On Error GoTo ImportarErr
nCuotaPagadaXLS = 0
cmdActualizar.Enabled = False
cmdGrabar.Enabled = False
lstDetalle.ListItems.Clear
lstDetalle.BackColor = lstCabecera.BackColor

pcIFTpo = Left(lblCodAdeudado, 2)
pcPersCod = Mid(lblCodAdeudado, 4, 13)
pcCtaIFCod = Mid(lblCodAdeudado, 18, 7)
     
lsArchivo = ""
Dialog.FileName = ""
Dialog.DialogTitle = "Adeudados: Importar Calendario"
Dialog.Flags = 21
Dialog.Filter = "*.xls"
Dialog.DefaultExt = "*.xls"
Dialog.ShowOpen
lsArchivo = Dialog.FileName
If lsArchivo = "" Then
    Screen.MousePointer = vbDefault
    MsgBox "Debe seleccionar un archivo Excel para Importar datos", vbInformation, "¡Aviso!"
Else

    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbExcel Then
        'ExcelAddHoja "Calendario", xlLibro, xlHoja1, True  '? para que agregar hoja?
         
        lstDetalle.ListItems.Clear
        lblImportar.Caption = "Importando Registro 0 - Por favor espere un momento ..."
        lblImportar.Visible = True
        Me.Enabled = False
        
        i = 1
        Set mySheet = xlAplicacion.Worksheets(1) 'Hoja 1
        With mySheet
            i = i + 1
            Do While .Cells(i, 1) <> ""
                'IIf(rs!cTpoCuota = gCGTipoCuotCalIFCuota, "Cuota", "No Concesional")
                'Nueva Fila
                Set L = lstDetalle.ListItems.Add(, , IIf(IIf(.Cells(i, 2) = "", False, True), "No Concesional", "Cuota"))
                    
                L.SubItems(2) = .Cells(i, 3)                            '"CUOTA N°"
                    
                If IIf(Trim(.Cells(i, 13)) = "" Or .Cells(i, 13) = "0", "0", "1") = "1" Then   ' SI LA CUOTA ES PAGADA
                    lstDetalle.ListItems.Item(i - 1).ListSubItems.Item(2).Bold = True
                    lstDetalle.ListItems.Item(i - 1).ListSubItems.Item(2).ForeColor = vbRed
                    nCuotaPagadaXLS = .Cells(i, 3)
                End If
                    
                lblImportar.Caption = "Importando Cuota " & .Cells(i, 3) & " - Por favor espere un momento ..."
                    
                If IIf(.Cells(i, 2) = "", 0, 1) = 1 Then                            '"Secuencia"
                    L.SubItems(9) = "6" '   NO CONSECIONAL
                Else
                    L.SubItems(9) = "2" '  CUOTA
                End If
                    
                L.SubItems(1) = Format(.Cells(i, 4), "dd/MM/YYYY")      '"Fecha de Vencimiento"
                    
                L.SubItems(3) = Format(.Cells(i, 7), "#0.00")           'Principal
                L.SubItems(4) = Format(.Cells(i, 8), "#0.00")           'Interes
                L.SubItems(5) = Format(.Cells(i, 9), "#0.00")           'Comisiones
                L.SubItems(6) = Format(.Cells(i, 10), "#0.00")          'Monto a Cobrar
                L.SubItems(11) = Format(.Cells(i, 5), "#0.00")          'Dias
                L.SubItems(7) = Format(.Cells(i, 11), "#0.00")          'Principal X Vencer
                L.SubItems(10) = IIf(Trim(.Cells(i, 13)) = "" Or .Cells(i, 13) = "0", "0", "1")    'nEstado
                L.SubItems(8) = ""
                i = i + 1
            Loop
        End With
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
        
        lblImportar.Visible = False
        Me.Enabled = True
        MsgBox "Archivo " & lsArchivo & Chr(10) & Chr(10) & " Importado Satisfactoriamente", vbInformation, "Aviso"
        
        cmdValidar.SetFocus
    End If
End If
 
Screen.MousePointer = vbDefault
Exit Sub
ImportarErr:
    lblImportar.Visible = True
    Me.Enabled = False
    MsgBox "Existieron problemas al Importar Archivo" & Chr(10) & TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

 
Private Sub cmdGrabar_Click()
Dim i As Integer

    If MsgBox("Se actualizará el sistema con los datos del calendario de Excel" & Chr(10) & Chr(10) & "Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        PB.Visible = True
        PB.Min = 0
        PB.Max = lstDetalle.ListItems.Count
        PB.value = 0
        i = 1
        frmAdeudCal.fgCronograma.Clear
        frmAdeudCal.fgCronograma.Rows = 2
        frmAdeudCal.fgCronograma.FormaCabecera
        If nCuotaPagada > 0 Then
            With frmAdeudCal.fgCronograma
                
                For i = 1 To nCuotaPagada 'lstCabecera.ListItems.Count
                    PB.value = PB.value + 1
                    If .Rows = 2 And Len(Trim(.TextMatrix(1, 2))) = 0 Then
                    Else
                        .AdicionaFila
                        '.Rows = .Rows + 1
                    End If
                    '#
                    .TextMatrix(.Rows - 1, 0) = "-"
                    'Fecha Pago
                    .TextMatrix(.Rows - 1, 1) = lstCabecera.ListItems(i).SubItems(1)
                    'Cuota
                    .TextMatrix(.Rows - 1, 2) = lstCabecera.ListItems(i).SubItems(2)
                    'Capital
                    .TextMatrix(.Rows - 1, 3) = lstCabecera.ListItems(i).SubItems(3)
                    'Interes
                    .TextMatrix(.Rows - 1, 4) = lstCabecera.ListItems(i).SubItems(4)
                    'Comision
                    .TextMatrix(.Rows - 1, 5) = lstCabecera.ListItems(i).SubItems(5)
                    'Total
                    .TextMatrix(.Rows - 1, 16) = lstCabecera.ListItems(i).SubItems(6)
                    'Saldo
                    .TextMatrix(.Rows - 1, 17) = lstCabecera.ListItems(i).SubItems(7)
                    'cTpoCuota
                    .TextMatrix(.Rows - 1, 8) = lstCabecera.ListItems(i).SubItems(9)
                    'cEstado
                    .TextMatrix(.Rows - 1, 9) = lstCabecera.ListItems(i).SubItems(10)
                    'nDiasPago
                    .TextMatrix(.Rows - 1, 10) = lstCabecera.ListItems(i).SubItems(11)
                    'nInteresPagado
                    .TextMatrix(.Rows - 1, 11) = lstCabecera.ListItems(i).SubItems(12)
                    'cMovNro
                    .TextMatrix(.Rows - 1, 12) = lstCabecera.ListItems(i).SubItems(8)
                    
                    .TextMatrix(.Rows - 1, 13) = Format("0", "#0.00")
                    .TextMatrix(.Rows - 1, 14) = Format("0", "#0.00")
                    .TextMatrix(.Rows - 1, 15) = Format("0", "#0.00")
                    
                    '1=Fecha NO
                    '2=NroCuota NO
                    '3=nCapital NO
                    '4=nInteres NO
                    '5=Comision NO
                    '6=Total NO
                    '7=Saldo Capital NO
                    '8 = cMovNro SI
                    '9 =cTpoCuota NO
                    '10 = cEstado SI
                    '11 = nDiasPago NO
                    '12 = nInteresPagado SI
                Next
            End With
        End If
        
        With frmAdeudCal.fgCronograma
            '.Clear
            '.Rows = 2
            '.FormaCabecera
            
            For i = i To lstDetalle.ListItems.Count
                PB.value = PB.value + 1
                If .Rows = 2 And Len(Trim(.TextMatrix(1, 2))) = 0 Then
                Else
                    .AdicionaFila
                    '.Rows = .Rows + 1
                End If
                '#
                .TextMatrix(.Rows - 1, 0) = "-"
                'Fecha Pago
                .TextMatrix(.Rows - 1, 1) = lstDetalle.ListItems(i).SubItems(1)
                'Cuota
                .TextMatrix(.Rows - 1, 2) = lstDetalle.ListItems(i).SubItems(2)
                'Capital
                .TextMatrix(.Rows - 1, 3) = lstDetalle.ListItems(i).SubItems(3)
                'Interes
                .TextMatrix(.Rows - 1, 4) = lstDetalle.ListItems(i).SubItems(4)
                'Comision
                .TextMatrix(.Rows - 1, 5) = lstDetalle.ListItems(i).SubItems(5)
                'Total
                .TextMatrix(.Rows - 1, 16) = lstDetalle.ListItems(i).SubItems(6)
                'Saldo
                .TextMatrix(.Rows - 1, 17) = lstDetalle.ListItems(i).SubItems(7)
                'cTpoCuota
                .TextMatrix(.Rows - 1, 8) = lstDetalle.ListItems(i).SubItems(9)
                'cEstado
                .TextMatrix(.Rows - 1, 9) = lstDetalle.ListItems(i).SubItems(10)
                'nDiasPago
                .TextMatrix(.Rows - 1, 10) = lstDetalle.ListItems(i).SubItems(11)
                'nInteresPagado
                .TextMatrix(.Rows - 1, 11) = lstDetalle.ListItems(i).SubItems(12)
                'cMovNro
                .TextMatrix(.Rows - 1, 12) = lstDetalle.ListItems(i).SubItems(8)
                
                .TextMatrix(.Rows - 1, 13) = Format("0", "#0.00")
                .TextMatrix(.Rows - 1, 14) = Format("0", "#0.00")
                .TextMatrix(.Rows - 1, 15) = Format("0", "#0.00")
                
                '1=Fecha NO
                '2=NroCuota NO
                '3=nCapital NO
                '4=nInteres NO
                '5=Comision NO
                '6=Total NO
                '7=Saldo Capital NO
                '8 = cMovNro SI
                '9 =cTpoCuota NO
                '10 = cEstado SI
                '11 = nDiasPago NO
                '12 = nInteresPagado SI
            Next
            
            frmAdeudCal.txtTotalInteres = Format(.SumaRow(4), gsFormatoNumeroView)
            frmAdeudCal.txtTotalcapital = Format(.SumaRow(3), gsFormatoNumeroView)
            frmAdeudCal.txtTotalGeneral = Format(.SumaRow(6), gsFormatoNumeroView)
        End With
        
        frmAdeudCal.Sumatoria
        MsgBox "Datos Transferidos ... No olvide Grabar", vbInformation, "Aviso"
        Unload Me
    End If
End Sub

Private Sub Form_Load()

Dim oCal As New nAdeudCal
Dim rs As New ADODB.Recordset
Dim L As ListItem
Dim i As Integer

Dim nTemp As Integer
Dim pnCapital As Double
Dim lnCapital As Double
Dim pnTramo As Integer
Dim lnCapitalCuota As Double
Dim lnNumCuotas As Integer
Dim nMontoOrigen As Double

CentraForm frmAdeudCalMnt

lblCodAdeudado.Caption = frmAdeudCal.lblCodAdeudado
lblDescEntidad.Caption = frmAdeudCal.lblDescEntidad
pnCapital = frmAdeudCal.txtCapital
pnTramo = frmAdeudCal.txtTramo
''''
If gsOpeCod = "401803" Or gsOpeCod = "402803" Then 'Nuevo
    Me.Caption = " REGISTRO DE CALENDARIO"
    nBandera = 1
ElseIf gsOpeCod = "401832" Or gsOpeCod = "402832" Then 'Nuevo
    Me.Caption = " MANTENIMIENTO DE CALENDARIO"
    nBandera = 2
End If

'Select c.cPersCod, c.cCtaIFCod, c.cTpoCuota, c.nNroCuota, C.dVencimiento,
'c.nInteresPagado , nCapital, nInteres, cEstado, nComision, nDiasPago, cMovNro
'From    CtaIfCalendario C
'Where  cIFTpo = '05' and cPersCod = '1129800000370' and cCtaIfcod = '0520183'
'Order BY C.cPersCod,C.cCtaIFCod, C.cTpoCuota,C.nNroCuota

'''
i = 0
lnCapital = pnCapital
nMontoOrigen = 0

If pnTramo <> 0 Then
    lnCapital = Round(pnCapital * pnTramo / 100, 2)
End If

lstCabecera.ListItems.Clear
Set rs = oCal.GetCalendarioDatos(Mid(lblCodAdeudado, 4, 13), Left(lblCodAdeudado, 2), Mid(lblCodAdeudado, 18, 7), , , "DESC")
lnNumCuotas = rs.RecordCount - 1
If rs.BOF Then
Else
    nTemp = rs!ctpocuota
    Do While Not rs.EOF
        
        i = i + 1
        
        
        
        Set L = lstCabecera.ListItems.Add(, , IIf(rs!ctpocuota = gCGTipoCuotCalIFCuota, "Cuota", "No Concesional"))
        If rs!ctpocuota = gCGTipoCuotCalIFCuota Or rs!ctpocuota = 3 Then '2
            'L.SubItems(1) = "CUOTA"
            lstCabecera.ListItems(i).ForeColor = vbBlue
        ElseIf rs!ctpocuota = gCGTipoCuotCalIFNoConcesional Then '6
            'L.SubItems(1) = "NO CONC."
            lstCabecera.ListItems(i).ForeColor = vbRed
        End If
                
        L.SubItems(1) = Format(rs!dVencimiento, "dd/MM/YYYY")
        L.SubItems(2) = rs!nNroCuota
        If rs!cEstado = "1" Then
            lstCabecera.ListItems.Item(i).ListSubItems.Item(2).Bold = True
            lstCabecera.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbRed
            nCuotaPagada = rs!nNroCuota
        End If
        L.SubItems(3) = Format(rs!nCapital, "0.00")
        lnCapitalCuota = Format(rs!nCapital, "0.00")
        L.SubItems(4) = Format(rs!nInteres, "0.00")
        L.SubItems(5) = Format(rs!nComision, "0.00")
        L.SubItems(6) = Format(rs!nCapital + rs!nInteres + rs!nComision, "0.00")
        L.SubItems(8) = Trim(rs!cMovNro)
        
        If Len(rs!cMovNro) > 0 Then
            lstCabecera.ListItems.Item(i).ListSubItems.Item(8).ForeColor = vbBlue
        End If
        
        If pnTramo <> 0 And rs!ctpocuota = gCGTipoCuotCalIFNoConcesional Then
            If rs!nNroCuota = 1 Then
                lnCapital = pnCapital - Round(pnCapital * pnTramo / 100, 2)
            End If
        End If
             
        If rs!nCapital > 0 Then
            lnCapital = Format(lnCapital - lnCapitalCuota, "0.00")
        End If
                        
        If nMontoOrigen = 0 Then
            nMontoOrigen = lnCapital + lnCapitalCuota
        End If
        
        If nTemp <> rs!ctpocuota And lnCapital < 0 Then
            nTemp = rs!ctpocuota
            lnCapital = pnCapital - nMontoOrigen - lnCapitalCuota
        End If
        
        L.SubItems(7) = Format(lnCapital, "0.00")
        L.SubItems(9) = rs!ctpocuota
        L.SubItems(10) = rs!cEstado
        L.SubItems(11) = rs!nDiasPago
        L.SubItems(12) = rs!nInteresPagado
            
        rs.MoveNext
    Loop
End If
Set rs = Nothing
Set oCal = Nothing

End Sub
 
Private Sub lstCabecera_DblClick()
    If lstCabecera.ListItems.Count > 0 Then
        txtcMovNro.Text = ""
        txtcMovNro.Text = lstCabecera.SelectedItem.SubItems(8)
        txtcMovNro.Visible = True
        txtcMovNro.SetFocus
    End If
End Sub
Private Sub txtcMovNro_GotFocus()
    Call fEnfoque(txtcMovNro)
End Sub

Private Sub txtcMovNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    lstCabecera.SetFocus
End If
End Sub

Private Sub txtcMovNro_LostFocus()
    txtcMovNro.Text = ""
    txtcMovNro.Visible = False
End Sub
