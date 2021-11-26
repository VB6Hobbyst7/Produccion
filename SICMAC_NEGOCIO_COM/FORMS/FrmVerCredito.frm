VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmVerCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Credito"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   3330
      Width           =   8955
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8985
      Begin MSComctlLib.ListView Lst 
         Height          =   3060
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   5398
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Linea de Credito"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Desembolso"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Analista"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Atraso"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmVerCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbLeasing As Boolean 'ALPA 20130108
Dim fbHistorial As Boolean 'JUEZ 20121215
Dim fbOtros As Boolean 'JUEZ 20130411
Dim CtlCtaCod As ActXCodCta

Public Sub Inicio(ByVal psPersCod As String, Optional pbLeasing As Boolean = False, Optional ByVal pbHistorial As Boolean = False, Optional ByVal pbOtros As Boolean = False, Optional ByVal pCtlCtaCod As ActXCodCta, _
                  Optional ByVal pbTransferido As Boolean = False)
    'ALPA 20130108 se agrego Optional pbLeasing As Boolean
    'JUEZ 20121215 se agregó fbHistorial
    'JUEZ 20130411 se agregó pbOtros y pCtlCtaCod para hacer dinamico el focus del ActXCodCta
    'FRHU 20150415 ERS022-2015 Se agrego pbTransferido
    Dim rs As ADODB.Recordset
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion
    Dim iTem As ListItem
    
    Set oVisualizacion = New COMNCredito.NCOMVisualizacion
    Set rs = oVisualizacion.VerCreditoByPersona(psPersCod, pbHistorial, pbTransferido)
    Set oVisualizacion = Nothing
    Lst.ListItems.Clear
    Do Until rs.EOF
        Set iTem = Lst.ListItems.Add(, , rs!cCtaCod)
        iTem.SubItems(1) = rs!cLineaCredDes
        iTem.SubItems(2) = Format(rs!nMontoDes, "#0.00")
        iTem.SubItems(3) = rs!Moneda
        iTem.SubItems(4) = IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre)
        iTem.SubItems(5) = rs!nDiasAtraso
        rs.MoveNext
    Loop
    Set rs = Nothing
    lbLeasing = pbLeasing 'ALPA 20130108
    fbHistorial = pbHistorial 'JUEZ 20121215
    'JUEZ 20130411 ************
    fbOtros = pbOtros
    Set CtlCtaCod = pCtlCtaCod
    'END JUEZ *****************
    Me.Show vbModal
'    Lst.SetFocus
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Lst.SetFocus
End Sub

Private Sub Lst_DblClick()
   If Not Lst.SelectedItem Is Nothing Then
        'JUEZ 20130411 ********************************************************************
        If fbOtros Then
            CtlCtaCod.NroCuenta = Lst.ListItems(Lst.SelectedItem.Index).Text
            CtlCtaCod.SetFocusCuenta
        Else
            'JUEZ 20121215 *********************************************
            If fbHistorial Then
                frmCredNewNivAprHist.ActxCta.NroCuenta = Lst.ListItems(Lst.SelectedItem.Index).Text
                frmCredNewNivAprHist.ActxCta.SetFocusCuenta
            Else
                'ALPA 20130108******************************************************
                If lbLeasing Then
                    frmCredpagoCuotasLeasingDetalle.ActxCta.SetFocusCuenta
                Else
                    'frmCredPagoCuotas.ActxCta.NroCuenta = Lst.ListItems(Lst.SelectedItem.Index).Text
                    frmCredPagoCuotas.ActxCta.SetFocusCuenta
                End If
                '*******************************************************************
            End If
            'END JUEZ **************************************************
        End If
        'END JUEZ **************************************************************************
        Unload Me
  End If
   
End Sub

Private Sub Lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lst_DblClick
    End If
End Sub
