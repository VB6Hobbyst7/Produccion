VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCCESaldosAFavorChe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos a Favor CCE - CHEQUE"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16785
   Icon            =   "frmCCESaldosAFavorChe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   16785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMensaje 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   13815
      Begin VB.Label lblMensaje 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SELECCIONE Y GUARDE LOS REGISTROS QUE SE ENVIARAN EN LA TRAMA DE SALDOS A FAVOR."
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
         Top             =   105
         Width           =   8895
      End
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   15360
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   13980
      TabIndex        =   3
      Top             =   5160
      Width           =   1335
   End
   Begin TabDlg.SSTab sstSaldoAFavor 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Presentados"
      TabPicture(0)   =   "frmCCESaldosAFavorChe.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSelPreTotalMN"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSelPreTotalME"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFechaSaldos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFechaSaldo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fePresentado"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Ajustes"
      TabPicture(1)   =   "frmCCESaldosAFavorChe.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "feAjuste"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "lblSelAjuTotalME"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "lblSelAjuTotalMN"
      Tab(1).ControlCount=   5
      Begin SICMACT.FlexEdit fePresentado 
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   6800
         Cols0           =   12
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro. Cheque-Entidad Financ.-Oficina-Cuenta-Monto-Moneda-Agencia-Fecha Registro-Fecha Presentación--nId"
         EncabezadosAnchos=   "300-1200-3500-700-2000-1200-1200-1800-1500-1800-500-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-10-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feAjuste 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   16335
         _ExtentX        =   28813
         _ExtentY        =   6800
         Cols0           =   12
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro. Cheque-Entidad Financ.-Oficina-Cuenta-Monto-Moneda-Agencia-Fecha Registro-Fecha Presentación--nId"
         EncabezadosAnchos=   "300-1200-3500-700-2000-1200-1200-1800-1500-1800-500-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-10-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-4-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSMask.MaskEdBox txtFechaSaldo 
         Height          =   315
         Left            =   14400
         TabIndex        =   15
         Top             =   4440
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFechaSaldos 
         Caption         =   "Fecha de Saldos a Favor :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12240
         TabIndex        =   16
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Total Sel. ME:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72240
         TabIndex        =   14
         Top             =   4470
         Width           =   1095
      End
      Begin VB.Label lblSelAjuTotalME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   -71040
         TabIndex        =   13
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Total Sel. MN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   4470
         Width           =   1215
      End
      Begin VB.Label lblSelAjuTotalMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   -73665
         TabIndex        =   11
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Total Sel. ME:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   4470
         Width           =   1095
      End
      Begin VB.Label lblSelPreTotalME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3960
         TabIndex        =   9
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Total Sel. MN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4470
         Width           =   1215
      End
      Begin VB.Label lblSelPreTotalMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1340
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCCESaldosAFavorChe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cancelar_Click()
    CargaDatos
End Sub
Private Sub cmdAceptar_Click()
    Dim oCCE As COMNCajaGeneral.NCOMCCE
    Dim pbSeleccion As Boolean
    Dim i As Integer
    Dim nRegSelPre(), nRegSelAju() As Long
    Set oCCE = New COMNCajaGeneral.NCOMCCE
    If Len(fePresentado.TextMatrix(1, 1)) = 0 And Len(feAjuste.TextMatrix(1, 1)) = 0 Then
        MsgBox "No existen registros para guardar.", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    pbSeleccion = False
    ReDim Preserve nRegSelPre(0)
    If Not Len(fePresentado.TextMatrix(1, 1)) = 0 Then
        For i = 1 To fePresentado.Rows - 1
           If fePresentado.TextMatrix(i, 10) = "." Then
                ReDim Preserve nRegSelPre(UBound(nRegSelPre) + 1)
                nRegSelPre(UBound(nRegSelPre)) = fePresentado.TextMatrix(i, 11)
                pbSeleccion = True
           End If
        Next i
    End If
    ReDim Preserve nRegSelAju(0)
    If Not Len(feAjuste.TextMatrix(1, 1)) = 0 Then
        For i = 1 To feAjuste.Rows - 1
           If feAjuste.TextMatrix(i, 10) = "." Then
                ReDim Preserve nRegSelAju(UBound(nRegSelAju) + 1)
                nRegSelAju(UBound(nRegSelAju)) = feAjuste.TextMatrix(i, 11)
                pbSeleccion = True
           End If
         Next i
    End If
    'VAPA20170323 CCE AGREGO VALIDACION DE FECHA DE SALDO
If Not UBound(nRegSelPre) = 0 Then
    If Not IsDate(txtFechaSaldo) Then
            MsgBox "Formato no Valido Fecha No Válida", vbInformation, "SICMACM - Aviso"
            txtFechaSaldo.SetFocus
            Exit Sub
        End If
      If CDate(txtFechaSaldo) < 0 Then
            MsgBox "Debe Ingresar una fecha para Saldos a Favor", vbInformation, "SICMACM - Aviso"
            txtFechaSaldo.SetFocus
            Exit Sub
        End If
End If
    'END VAPA20170323
    If Not pbSeleccion Then
        MsgBox "No se ha seleccionado ningún registro. Verifique.", vbInformation, "¡Aviso!"
        Exit Sub
    Else
        If MsgBox("Se han seleccionado " & UBound(nRegSelPre) & " registro(s) de presentados y " & Chr(10) _
                & UBound(nRegSelAju) & " registro(s) de ajustes, ¿Está seguro de guardar la " & Chr(10) _
                & "información? ", vbYesNo + vbInformation, "¡Aviso!") = vbYes Then
            If Not UBound(nRegSelPre) = 0 Then
                For i = 1 To UBound(nRegSelPre)
                    oCCE.CCE_DocRecSaldoAFavor nRegSelPre(i), gsCodUser, "PRE", CDate(txtFechaSaldo) 'VAPA20170323 AGREGO txtFechaSaldo
                Next i
            End If
            If Not UBound(nRegSelAju) = 0 Then
                For i = 1 To UBound(nRegSelAju)
                    oCCE.CCE_DocRecSaldoAFavor nRegSelAju(i), gsCodUser, "AJU", gdFecSis 'VAPA20170323 AGREGO gdFecSis
                Next i
            End If
            MsgBox "Se ha realizado el registro satisfactoriamente.", vbInformation, "¡Aviso!"
            CargaDatos
        Else
            Exit Sub
        End If
    End If
End Sub
Private Sub feAjuste_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If feAjuste.TextMatrix(pnRow, 10) = "." Then
        If feAjuste.TextMatrix(pnRow, 6) = "SOLES" Then
            lblSelAjuTotalMN.Caption = Format(nVal(lblSelAjuTotalMN) + nVal(feAjuste.TextMatrix(pnRow, 5)), "#,##0.00")
        Else
            lblSelAjuTotalME.Caption = Format(nVal(lblSelAjuTotalME) + nVal(feAjuste.TextMatrix(pnRow, 5)), "#,##0.00")
        End If
    Else
        If feAjuste.TextMatrix(pnRow, 6) = "SOLES" Then
            lblSelAjuTotalMN.Caption = Format(nVal(lblSelAjuTotalMN) - nVal(feAjuste.TextMatrix(pnRow, 5)), "#,##0.00")
        Else
            lblSelAjuTotalME.Caption = Format(nVal(lblSelAjuTotalME) - nVal(feAjuste.TextMatrix(pnRow, 5)), "#,##0.00")
        End If
    End If
End Sub
Private Sub fePresentado_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If fePresentado.TextMatrix(pnRow, 10) = "." Then
        If fePresentado.TextMatrix(pnRow, 6) = "SOLES" Then
            lblSelPreTotalMN.Caption = Format(nVal(lblSelPreTotalMN) + nVal(fePresentado.TextMatrix(pnRow, 5)), "#,##0.00")
        Else
            lblSelPreTotalME.Caption = Format(nVal(lblSelPreTotalME) + nVal(fePresentado.TextMatrix(pnRow, 5)), "#,##0.00")
        End If
    Else
        If fePresentado.TextMatrix(pnRow, 6) = "SOLES" Then
            lblSelPreTotalMN.Caption = Format(nVal(lblSelPreTotalMN) - nVal(fePresentado.TextMatrix(pnRow, 5)), "#,##0.00")
        Else
            lblSelPreTotalME.Caption = Format(nVal(lblSelPreTotalME) - nVal(fePresentado.TextMatrix(pnRow, 5)), "#,##0.00")
        End If
    End If
End Sub
'VAPA20170323 CCE
Private Sub optTipoBus_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtFechaSaldo.Visible Then
        txtFechaSaldo.SetFocus
    End If
End If
End Sub
Private Sub txtFechaSaldo_GotFocus()
With txtFechaSaldo
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

'END VAPA20170323 CCE
Private Sub Form_Load()
    CargaDatos
End Sub
Private Sub CargaDatos()
Dim rs As ADODB.Recordset
Dim oCCE As COMNCajaGeneral.NCOMCCE

    Set oCCE = New COMNCajaGeneral.NCOMCCE
    LimpiaFlex fePresentado
    LimpiaFlex feAjuste
    Set rs = oCCE.CCE_SaldosAFavorPresentado
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            fePresentado.AdicionaFila
            fePresentado.TextMatrix(fePresentado.row, 1) = rs!cNroCheque
            fePresentado.TextMatrix(fePresentado.row, 2) = rs!cCodEntDebitar
            fePresentado.TextMatrix(fePresentado.row, 3) = rs!cCodOficina
            fePresentado.TextMatrix(fePresentado.row, 4) = rs!cCtaDebitar
            fePresentado.TextMatrix(fePresentado.row, 5) = Format(rs!cImporte, "#,##0.00")
            fePresentado.TextMatrix(fePresentado.row, 6) = rs!cCodMoneda
            fePresentado.TextMatrix(fePresentado.row, 7) = rs!cAgeDescripcion
            fePresentado.TextMatrix(fePresentado.row, 8) = rs!dRegistro
            fePresentado.TextMatrix(fePresentado.row, 9) = rs!dFechaArchivo
            fePresentado.TextMatrix(fePresentado.row, 11) = rs!nId
            rs.MoveNext
        Loop
    End If
    Set rs = oCCE.CCE_SaldosAFavorAjuste
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            feAjuste.AdicionaFila
            feAjuste.TextMatrix(feAjuste.row, 1) = rs!cNroCheque
            feAjuste.TextMatrix(feAjuste.row, 2) = rs!cCodEntDebitar
            feAjuste.TextMatrix(feAjuste.row, 3) = rs!cCodOficina
            feAjuste.TextMatrix(feAjuste.row, 4) = rs!cCtaDebitar
            feAjuste.TextMatrix(feAjuste.row, 5) = Format(rs!cImporte, "#,##0.00")
            feAjuste.TextMatrix(feAjuste.row, 6) = rs!cCodMoneda
            feAjuste.TextMatrix(feAjuste.row, 7) = rs!cAgeDescripcion
            feAjuste.TextMatrix(feAjuste.row, 8) = rs!dRegistro
            feAjuste.TextMatrix(feAjuste.row, 9) = rs!dFechaArchivo
            feAjuste.TextMatrix(feAjuste.row, 11) = rs!nId
            rs.MoveNext
        Loop
    End If
End Sub
