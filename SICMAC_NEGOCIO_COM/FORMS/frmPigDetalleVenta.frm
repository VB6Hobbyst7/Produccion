VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPigDetalleVenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Venta"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   Icon            =   "frmPigDetalleVenta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnular 
      Caption         =   "&Anular"
      Height          =   375
      Left            =   7335
      TabIndex        =   28
      Top             =   5685
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8655
      TabIndex        =   26
      Top             =   5685
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   5625
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   9930
      Begin VB.Frame frCliente 
         Caption         =   "Datos del Cliente"
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
         Height          =   945
         Left            =   90
         TabIndex        =   11
         Top             =   1005
         Width           =   9750
         Begin MSComctlLib.ListView lstCliente 
            Height          =   705
            Left            =   60
            TabIndex        =   12
            Top             =   165
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1244
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   10485760
            BackColor       =   16777215
            BorderStyle     =   1
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre o Razón Social"
               Object.Width           =   5822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Documento"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Dirección"
               Object.Width           =   5115
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Telefono"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Frame frDocumento 
         Caption         =   "Documento"
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
         Height          =   885
         Left            =   2490
         TabIndex        =   5
         Top             =   120
         Width           =   4260
         Begin VB.OptionButton OptDocumento 
            Caption         =   "Factura"
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   7
            Top             =   555
            Width           =   855
         End
         Begin VB.OptionButton OptDocumento 
            Caption         =   "Boleta"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   285
            Width           =   855
         End
         Begin MSMask.MaskEdBox mskNroDocumento 
            Height          =   300
            Left            =   2580
            TabIndex        =   8
            Top             =   345
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSerDocumento 
            Height          =   300
            Left            =   1995
            TabIndex        =   9
            Top             =   345
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "Número :"
            Height          =   195
            Left            =   1260
            TabIndex        =   10
            Top             =   405
            Width           =   645
         End
      End
      Begin VB.Frame frTipoVenta 
         Caption         =   "Tipo Venta"
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
         Height          =   885
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   2340
         Begin VB.OptionButton optTipoVenta 
            Caption         =   "A Terceros"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   4
            Top             =   225
            Width           =   1470
         End
         Begin VB.OptionButton optTipoVenta 
            Caption         =   "Por Responsabilidad"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   3
            Top             =   420
            Width           =   1890
         End
         Begin VB.OptionButton optTipoVenta 
            Caption         =   "Al Titular"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   2
            Top             =   615
            Width           =   1470
         End
      End
      Begin VB.Frame frJoyas 
         Caption         =   "Joyas"
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
         Height          =   3585
         Left            =   90
         TabIndex        =   13
         Top             =   1935
         Width           =   9750
         Begin MSComctlLib.ListView lstJoyas 
            Height          =   2475
            Left            =   45
            TabIndex        =   27
            Top             =   195
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4366
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sec"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Nro.Contrato"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Pza"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Descripción de la Joya"
               Object.Width           =   9949
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Monto Venta"
               Object.Width           =   1941
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Igv"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "NroRemate"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblIGV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   8415
            TabIndex        =   23
            Top             =   2745
            Width           =   1200
         End
         Begin VB.Label lblDescuento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   3720
            TabIndex        =   22
            Top             =   2745
            Width           =   1200
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6075
            TabIndex        =   21
            Top             =   2745
            Width           =   1200
         End
         Begin VB.Label Label4 
            Caption         =   "% Dcto.     :"
            Height          =   195
            Left            =   2775
            TabIndex        =   20
            Top             =   2805
            Width           =   840
         End
         Begin VB.Label Label5 
            Caption         =   "SubTotal   :"
            Height          =   195
            Left            =   5115
            TabIndex        =   19
            Top             =   2805
            Width           =   840
         End
         Begin VB.Label Label6 
            Caption         =   "I.G.V.        :"
            Height          =   195
            Left            =   7470
            TabIndex        =   18
            Top             =   2805
            Width           =   840
         End
         Begin VB.Label Label7 
            Caption         =   "Total US$ :"
            Height          =   195
            Left            =   5115
            TabIndex        =   17
            Top             =   3210
            Width           =   840
         End
         Begin VB.Label lblTotalD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6075
            TabIndex        =   16
            Top             =   3150
            Width           =   1200
         End
         Begin VB.Label Label9 
            Caption         =   "Total S/.   :"
            Height          =   195
            Left            =   7470
            TabIndex        =   15
            Top             =   3210
            Width           =   840
         End
         Begin VB.Label lblTotalS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00404080&
            Height          =   315
            Left            =   8415
            TabIndex        =   14
            Top             =   3150
            Width           =   1200
         End
      End
      Begin VB.Label lblSituacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   6840
         TabIndex        =   29
         Top             =   165
         Width           =   2955
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Cambio    :"
         Height          =   255
         Left            =   6960
         TabIndex        =   25
         Top             =   735
         Width           =   1410
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   8490
         TabIndex        =   24
         Top             =   690
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmPigDetalleVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'* frmPigDetalleVenta - Venta de Joyas en Tienda
'* EAFA - 15/10/2002
'****************************************************
Dim lnMovNroAnt As Long
Dim lmJoyas(19, 5) As String
Dim nFilas As Integer

Public Sub Inicio(ByVal lnTipDoc As Integer, ByVal lnNroDoc As String, ByVal lsFecVta As String)
GetTipCambio (lsFecVta)
lblTipoCambio.Caption = Format(gnTipCambioC, "##0.000 ")
MuestraDetalleDocumento lnTipDoc, lnNroDoc
Me.Show 1
End Sub

Private Sub cmdAnular_Click()
Dim lrPigContrato As NPigContrato
Dim loContFunct As NContFunciones
Dim lsMovNro As String

 If MsgBox(" Esta Ud seguro de Extornar dicha Operación ? ", vbQuestion + vbYesNo + vbDefaultButton2, " Aviso ") = vbNo Then
     Exit Sub
 Else
     Set loContFunct = New NContFunciones
     lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
     Set loContFunct = Nothing

     Set lrPigContrato = New NPigContrato
            Call lrPigContrato.nAnularVentaJoyas(lnMovNroAnt, lmJoyas, lsMovNro, nFilas)
     Set lrPigContrato = Nothing
     lblSituacion.Caption = "ANULADA"
     cmdAnular.Enabled = False
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub MuestraDatosCliente(ByVal psCliente As String)
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato
Dim lstTmpCliente As ListItem

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
      Set lrDatos = lrPigContrato.dObtieneDatosClienteVentaJoyas(psCliente)
Set lrPigContrato = Nothing

If lrDatos.EOF And lrDatos.BOF Then
    Exit Sub
Else
     lstCliente.ListItems.Clear
     Set lstTmpCliente = lstCliente.ListItems.Add(, , Format(lrDatos!cPersCod, "##0"))
            lstTmpCliente.SubItems(1) = PstaNombre(Trim(lrDatos!cPersNombre))
            If lrDatos!nPersPersoneria = gPersonaNat Then
                lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrDatos!NroDNI), "", lrDatos!NroDNI))
            Else
                lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(lrDatos!NroRUC), "", lrDatos!NroRUC))
            End If
            lstTmpCliente.SubItems(3) = Format(lrDatos!cPersDireccDomicilio, "##0")
            lstTmpCliente.SubItems(4) = Format(lrDatos!cPersTelefono)
     Set lrDatos = Nothing
End If
End Sub

Private Sub MuestraDetalleDocumento(ByVal pnTipDoc As Integer, ByVal psNroDoc As String)
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato
Dim lstTmpJoyas As ListItem
Dim fnVarSubTotal As Currency
Dim fnVarIGV As Currency
Dim fnVarTotalS As Currency
Dim fnVarTotalD As Currency
Dim f As Integer

fnVarSubTotal = 0: fnVarIGV = 0: fnVarTotalS = 0: fnVarTotalD = 0: f = 0

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
      Set lrDatos = lrPigContrato.dObtieneDetDocumentoVentasJoyas(pnTipDoc, Mid(psNroDoc, 1, 4) & Mid(psNroDoc, 6, 8))
Set lrPigContrato = Nothing

If lrDatos.EOF And lrDatos.BOF Then
    Exit Sub
Else
     Select Case lrDatos!nCodTipo             '****** Tipo de Documento
                 Case gPigTipoBoleta
                           OptDocumento(0).value = True
                 Case gPigTipoFactura
                           OptDocumento(1).value = True
     End Select
     Select Case lrDatos!nMotivo                '****** Tipo de Venta (Motivo)
                 Case gPigTipoVentaATerceros
                           optTipoVenta(0).value = True
                 Case gPigTipoVentaPorResponsabilidad
                           optTipoVenta(1).value = True
                 Case gPigTipoVentaAlTitular
                           optTipoVenta(2).value = True
     End Select
     If lrDatos!nFlag = gMovFlagExtornado Then
         lblSituacion.Caption = "ANULADA"
         cmdAnular.Enabled = False
     End If
     mskSerDocumento.Text = Mid(lrDatos!cDocumento, 1, 4)          '**** Serie de Boleta/Factura
     mskNroDocumento.Text = Mid(lrDatos!cDocumento, 6, 8)         '**** Número de Boleta/Factura
     lnMovNroAnt = lrDatos!nNroMov
     psCliente = lrDatos!cPersCod
     lstJoyas.ListItems.Clear
     Do While Not lrDatos.EOF
            Set lstTmpJoyas = lstJoyas.ListItems.Add(, , Format(lrDatos!nItemDoc, "##0"))
                   lstTmpJoyas.SubItems(1) = Trim(lrDatos!cCtaCod)
                   lstTmpJoyas.SubItems(2) = Format(lrDatos!nItemPieza, "##0")
                   lstTmpJoyas.SubItems(3) = Trim(lrDatos!Tipo) & " DE " & Trim(lrDatos!Material) & " DE " & Format(lrDatos!pNeto, "#,##0.00") & " grs. " & Trim(lrDatos!Observacion) & " " & Trim(lrDatos!ObservAdic)
                   lstTmpJoyas.SubItems(4) = Format(lrDatos!SubVenta, "###,##0.00")
                   lstTmpJoyas.SubItems(5) = Format(lrDatos!IGV, "##,##0.00")
                   lstTmpJoyas.SubItems(6) = lrDatos!nRema
                   
                   fnVarSubTotal = fnVarSubTotal + Val(lrDatos!SubVenta)
                   fnVarIGV = fnVarIGV + Val(lrDatos!IGV)
                   
                   lmJoyas(f, 0) = Trim(lrDatos!cCtaCod)           '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 1) = Trim(lrDatos!nItemPieza)        '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 2) = Trim(lrDatos!nRema)               '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 3) = Trim(lrDatos!SubVenta)          '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   lmJoyas(f, 4) = Trim(lrDatos!IGV)          '*** Carga a una Matriz, para luego pasarlos como paramtros para la Anulación
                   f = f + 1
            lrDatos.MoveNext
     Loop
     nFilas = f - 1
     Set lrDatos = Nothing
                   
     fnVarTotalS = Round(fnVarSubTotal + fnVarIGV, 2)
     fnVarTotalD = Round(fnVarTotalS / Val(lblTipoCambio.Caption), 2)
     lblSubTotal.Caption = Format(fnVarSubTotal, "###,##0.00 ")
     lblIGV.Caption = Format(Round(fnVarIGV, 2), "###,##0.00 ")
     lblTotalS.Caption = Format(fnVarTotalS, "###,##0.00 ")
     lblTotalD.Caption = Format(fnVarTotalD, "###,##0.00 ")
     
     MuestraDatosCliente (psCliente)
End If
End Sub
