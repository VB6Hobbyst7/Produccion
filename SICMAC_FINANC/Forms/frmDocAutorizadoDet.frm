VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDocAutorizadoDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Comprobante de Pago Autorizado"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "frmDocAutorizadoDet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   9050
      TabIndex        =   25
      Top             =   4680
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   10440
      TabIndex        =   26
      Top             =   4680
      Width           =   1365
   End
   Begin TabDlg.SSTab sstCompPag 
      Height          =   4455
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Comprobante de Pago"
      TabPicture(0)   =   "frmDocAutorizadoDet.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTipoCambio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPvtaMN"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblVvtaTotal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPvtaTotal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraInformGen"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraComprobante"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTpoCambio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtPvtaMN"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtVvtaTotal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtIGV"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkIGV"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPvtaTotal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox txtPvtaTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   10080
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   3960
         Width           =   1365
      End
      Begin VB.CheckBox chkIGV 
         Caption         =   "I.G.V.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9000
         TabIndex        =   21
         Top             =   3650
         Width           =   900
      End
      Begin VB.TextBox txtIGV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   345
         Left            =   10080
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   3600
         Width           =   1365
      End
      Begin VB.TextBox txtVvtaTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   345
         Left            =   10080
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   3240
         Width           =   1365
      End
      Begin VB.TextBox txtPvtaMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   345
         Left            =   10080
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   2760
         Width           =   1365
      End
      Begin VB.TextBox txtTpoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   345
         Left            =   10080
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Frame fraComprobante 
         Caption         =   "Comprobante"
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
         Height          =   1935
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   4695
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmDocAutorizadoDet.frx":0326
            Left            =   3240
            List            =   "frmDocAutorizadoDet.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1450
            Width           =   1365
         End
         Begin MSMask.MaskEdBox mebFechaEm 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   1450
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cboOpeTpo 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   3285
         End
         Begin VB.Label lblNroVal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2520
            TabIndex        =   33
            Top             =   1080
            Width           =   2085
         End
         Begin VB.Label lblSerieVal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1320
            TabIndex        =   32
            Top             =   1080
            Width           =   1125
         End
         Begin VB.Label lblTipoVal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1320
            TabIndex        =   31
            Top             =   720
            Width           =   3285
         End
         Begin VB.Label lblMoneda 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label lblNro 
            Caption         =   "Nro:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblTipo 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   750
            Width           =   975
         End
         Begin VB.Label lblOperacion 
            Caption         =   "Operación:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblEmision 
            Caption         =   "Emisión:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.Frame fraInformGen 
         Caption         =   "Información General"
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
         Height          =   3975
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   6615
         Begin VB.TextBox txtConcepto 
            Appearance      =   0  'Flat
            Height          =   2745
            Left            =   1320
            MaxLength       =   800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1100
            Width           =   5200
         End
         Begin Sicmact.TxtBuscar txtProvCod 
            Height          =   345
            Left            =   1320
            TabIndex        =   2
            Top             =   360
            Width           =   1695
            _extentx        =   2990
            _extenty        =   609
            appearance      =   0
            appearance      =   0
            font            =   "frmDocAutorizadoDet.frx":03B4
            appearance      =   0
            tipobusqueda    =   3
            stitulo         =   ""
            tipobuspers     =   1
         End
         Begin VB.Label lblAreaAgeDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   3120
            TabIndex        =   30
            Top             =   720
            Width           =   3405
         End
         Begin VB.Label lblAreaAge 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1320
            TabIndex        =   29
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblProvNom 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   3120
            TabIndex        =   28
            Top             =   360
            Width           =   3405
         End
         Begin VB.Label lblConcepto 
            Caption         =   "Concepto:"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblAreaUsu 
            Caption         =   "Área Usuaria:"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   750
            Width           =   975
         End
         Begin VB.Label lblCliente 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Label lblPvtaTotal 
         Caption         =   "Precio Venta:"
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
         Left            =   8700
         TabIndex        =   23
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Label lblVvtaTotal 
         Caption         =   "Valor Venta:"
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
         Left            =   8760
         TabIndex        =   19
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label lblPvtaMN 
         Caption         =   "Precio de Venta(MN):"
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
         Left            =   8000
         TabIndex        =   17
         Top             =   2775
         Width           =   1935
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Tipo de Cambio:"
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
         Left            =   8490
         TabIndex        =   15
         Top             =   2430
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDocAutorizadoDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'** Nombre : frmDocAutorizado
'** Descripción : Para el registro de Documentos Autorizados, TI-ERS0532016
'** Creación : PASI, 20161104
'**********************************************************************
Option Explicit
Dim rs As New ADODB.Recordset
Dim oReg As DRegVenta
Dim nTasaIGV As Currency
Dim gsDocNro As String
Dim gsDocNroAnt As String 'NAGL ERS 012-2017
Dim gdFechaAnt As Date
Public Sub inicio(Optional ByVal psDocNro As String = "")
    gsDocNro = psDocNro
    Me.Show 1
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If Len(txtProvCod.Text) = 0 Then MsgBox "Falta indicar el cliente. Verifique.", vbInformation, "¡Aviso!": txtProvCod.SetFocus: Exit Function
    If Len(lblProvNom.Caption) = 0 Then MsgBox "Falta indicar el cliente. Verifique.", vbInformation, "¡Aviso!": txtProvCod.SetFocus: Exit Function
    If Len(lblAreaAge.Caption) = 0 Then MsgBox "No se indica el código de área, consulte al Dpto. de TI.", vbInformation, "¡Aviso!":  Exit Function
    If Len(lblAreaAgeDesc.Caption) = 0 Then MsgBox "No se indica la descripción de área, consulte al Dpto. de TI.", vbInformation, "¡Aviso!": Exit Function
    If Len(LTrim(RTrim(Replace(Replace(txtConcepto.Text, Chr(10), ""), Chr(13), "")))) = 0 Then MsgBox "Falta indicar el concepto. Verifique.", vbInformation, "¡Aviso!": txtConcepto.SetFocus: Exit Function
    If cboOpeTpo.ListIndex = -1 Then MsgBox "Falta indicar la operación. Verifique.", vbInformation, "¡Aviso!": cboOpeTpo.SetFocus: Exit Function
    If Len(lblTipoVal.Caption) = 0 Then MsgBox "No se indica el tipo de comprobante, consulte al Dpto. de TI.", vbInformation, "¡Aviso!": Exit Function
    If Len(lblSerieVal.Caption) = 0 Then MsgBox "No se indica la serie del comprobante, consulte al Dpto. de TI.", vbInformation, "¡Aviso!": Exit Function
    If Len(lblNroVal.Caption) = 0 Then MsgBox "No se indica el número de comprobante, consulte al Dpto. de TI.", vbInformation, "¡Aviso!": Exit Function
    If Not IsDate(mebFechaEm.Text) Then MsgBox "La fecha de emisión no es correcta. Verifique.", vbInformation, "¡Aviso!": mebFechaEm.SetFocus: Exit Function
    If cboMoneda.ListIndex = -1 Then MsgBox "No se ha seleccionado la moneda. Verifique.": cboMoneda.SetFocus: Exit Function
    If Right(cboMoneda.Text, 1) = 2 And nVal(txtTpoCambio.Text) = 0 Then MsgBox "El tipo de cambio no puede ser cero. Verifique", vbInformation, "¡Aviso!": txtTpoCambio.SetFocus: Exit Function
    'If nVal(txtVvtaTotal.Text) = 0 Then MsgBox "El valor de venta no puede ser cero. Verifique", vbInformation, "¡Aviso!": txtTpoCambio.SetFocus: Exit Function /**Comentado PASI20170420**/
    'If nVal(txtPvtaTotal.Text) = 0 Then MsgBox "El precio de venta no puede ser cero. Verifique", vbInformation, "¡Aviso!": txtPvtaTotal.SetFocus: Exit Function /**Comentado PASI20170420**/
    ValidaDatos = True
End Function
Private Sub cmdAceptar_Click()
Dim oMov As DMov
Dim sMovNro As String
Dim nMovNro As Long
Dim RSBusca As ADODB.Recordset
Dim lsSerieVal As String
Dim lsNroVal As String
On Error GoTo errAcepta
    If Not ValidaDatos Then Exit Sub
    If Len(gsDocNro) = 0 Then
        If MsgBox(" ¿ Seguro de grabar datos ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Confirmación") = vbNo Then Exit Sub
        'Set RSBusca = oReg.VerificaVentaExistente(Left(lblTipoVal.Caption, 2), Trim(lblSerieVal.Caption & lblNroVal.Caption), mebFechaEm.Text, Right(cboOpeTpo, 1)) 'Comments PASI20170421
        'If Not (RSBusca.EOF And RSBusca.BOF) Then MsgBox "Este registro ya fue ingresado.", vbOKOnly + vbExclamation, "Atención": Exit Sub 'Comments PASI20170421
        Set oMov = New DMov
        sMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov sMovNro, gsOpeCod, "REGISTRO DE VENTAS", gMovEstContabNoContable, gMovFlagVigente
        nMovNro = oMov.GetnMovNro(sMovNro)
        Set oMov = Nothing
        'PASI20170421 ****************
        Set rs = oReg.DocAutDevuelveNroComprobante(gsCodAge)
        lsSerieVal = rs!cSerie
        lsNroVal = rs!cNro
        'PASI END***
        'If Right(cboMoneda, 1) = 1 Then oReg.InsertaVenta Right(cboOpeTpo.Text, 1), Left(lblTipoVal.Caption, 2), Trim(lblSerieVal.Caption & lblNroVal.Caption), mebFechaEm.Text, txtProvCod.Tag, "", txtConcepto.Text, nVal(txtVvtaTotal.Text), nVal(txtIGV.Text), nVal(txtPvtaTotal.Text), "", gdFecSis, 0, nMovNro, Right(cboMoneda.Text, 1), nVal(txtTpoCambio.Text), Trim(lblAreaAge.Caption) /**Comments PASI20170421**/
        If Right(cboMoneda, 1) = 1 Then oReg.InsertaVenta Right(cboOpeTpo.Text, 1), Left(lblTipoVal.Caption, 2), Trim(lsSerieVal & lsNroVal), mebFechaEm.Text, txtProvCod.Tag, "", txtConcepto.Text, nVal(txtVvtaTotal.Text), nVal(txtIGV.Text), nVal(txtPvtaTotal.Text), "", gdFecSis, 0, nMovNro, Right(cboMoneda.Text, 1), nVal(txtTpoCambio.Text), Trim(lblAreaAge.Caption) 'PASI 20170421
        'If Right(cboMoneda, 1) = 2 Then oReg.InsertaVenta Right(cboOpeTpo.Text, 1), Left(lblTipoVal.Caption, 2), Trim(lblSerieVal.Caption & lblNroVal.Caption), mebFechaEm.Text, txtProvCod.Tag, "", txtConcepto.Text, nVal(txtVvtaTotal.Text) * nVal(txtTpoCambio.Text), nVal(txtIGV.Text) * nVal(txtTpoCambio.Text), nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), "", gdFecSis, 0, nMovNro, Right(cboMoneda.Text, 1), txtTpoCambio.Text, Trim(lblAreaAge.Caption) /**Comments PASI20170421**/
        If Right(cboMoneda, 1) = 2 Then oReg.InsertaVenta Right(cboOpeTpo.Text, 1), Left(lblTipoVal.Caption, 2), Trim(lsSerieVal & lsNroVal), mebFechaEm.Text, txtProvCod.Tag, "", txtConcepto.Text, nVal(txtVvtaTotal.Text) * nVal(txtTpoCambio.Text), nVal(txtIGV.Text) * nVal(txtTpoCambio.Text), nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), "", gdFecSis, 0, nMovNro, Right(cboMoneda.Text, 1), txtTpoCambio.Text, Trim(lblAreaAge.Caption) 'PASI20170421
        MsgBox "El registro se ha realizado correctamente.", vbInformation, "¡Aviso!"
    Else
        If MsgBox(" ¿ Seguro de modificar los datos ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Confirmación") = vbNo Then Exit Sub
        'PASI20170421 ****
        lsSerieVal = lblSerieVal.Caption
        lsNroVal = lblNroVal.Caption
        'PASI END***
        If Right(cboMoneda, 1) = 1 Then oReg.ActualizaDocAutorizado Right(cboOpeTpo.Text, 1), Left(lblTipoVal.Caption, 2), Trim(lblSerieVal.Caption & lblNroVal.Caption), mebFechaEm.Text, txtProvCod.Tag, "", txtConcepto.Text, nVal(txtVvtaTotal.Text), nVal(txtIGV.Text), nVal(txtPvtaTotal.Text), Left(lblTipoVal.Caption, 2), gsDocNroAnt, Format(gdFechaAnt, "YYYYmmdd"), "", gdFecSis, 0, Right(cboMoneda.Text, 1), nVal(txtTpoCambio.Text), Trim(lblAreaAge.Caption)
        If Right(cboMoneda, 2) = 2 Then oReg.ActualizaDocAutorizado Right(cboOpeTpo.Text, 1), Left(lblTipoVal.Caption, 2), Trim(lblSerieVal.Caption & lblNroVal.Caption), mebFechaEm.Text, txtProvCod.Tag, "", txtConcepto.Text, nVal(txtVvtaTotal.Text) * nVal(txtTpoCambio.Text), nVal(txtIGV.Text) * nVal(txtTpoCambio.Text), nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), Left(lblTipoVal.Caption, 2), gsDocNroAnt, Format(gdFechaAnt, "YYYYmmdd"), "", gdFecSis, 0, Right(cboMoneda.Text, 1), nVal(txtTpoCambio.Text), Trim(lblAreaAge.Caption)
        MsgBox "La actualización se ha realizado correctamente.", vbInformation, "¡Aviso!"
    End If 'Se adecuó al cDocNro Ant. -> gsDocNroAnt 'NAGL ERS 012-2017
    'ImprimeComprobanteAutorizado Trim(lblSerieVal.Caption & lblNroVal.Caption) '/**Comments PASI20170421**/
    ImprimeComprobanteAutorizado Trim(lsSerieVal & lsNroVal) 'PASI20170421
    Unload Me
    Exit Sub
errAcepta:
   MsgBox TextErr(Err.Description), vbInformation, "! Aviso !"
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim oPer As New UPersona
    Dim oCons As NConstSistemas: Set oCons = New NConstSistemas
    Dim oOpe As New DOperacion: Dim oImp As New DImpuesto
    Dim sSerie, sNro, sCtaIGV As String
    
    Set oReg = New DRegVenta
    Set rs = oReg.CargaRegOperacion()
    RSLlenaCombo rs, cboOpeTpo
    mebFechaEm.Text = gdFecSis
    lblAreaAge.Caption = gsCodArea & gsCodAge
    lblAreaAgeDesc.Caption = gsNomAge
    lblTipoVal.Caption = oCons.LeeConstSistema(535)
    Set rs = oReg.DocAutDevuelveNroComprobante(gsCodAge)
    lblSerieVal.Caption = rs!cSerie
    lblNroVal.Caption = rs!cNro
    Set rs = oOpe.CargaOpeCta(gsOpeCod)
    If Not rs.EOF Then sCtaIGV = rs!cCtaContCod
    Set rs = oImp.CargaImpuesto(sCtaIGV)
    If rs.EOF Then MsgBox "No se definio tasa de IGV. Consultar con Sistemas!", vbInformation, "Aviso": Exit Sub
    nTasaIGV = (rs!nImpTasa / 100)
    Me.Caption = "Registro de Comprobante de Pago Autorizado"
    If Not Len(gsDocNro) = 0 Then
        Set rs = oReg.ObtieneDocAutorizadoxNroDoc(gsDocNro)
        If Not rs.EOF Then
            cboOpeTpo.ListIndex = BuscaCombo(rs!cOpeTpo, cboOpeTpo)
            lblSerieVal.Caption = Mid(gsDocNro, 1, 4) 'NAGL ERS 012-2017 Se cambió de Long.Nro Serie de 3 a 4 Dígitos
            lblNroVal.Caption = Mid(gsDocNro, 5, 12) 'NAGL ERS 012-2017 Se cambió de Long.NroRest, a partir del 5to dígito
            gsDocNroAnt = rs!cDocNro 'Agregado by NAGL ERS012-2017
            mebFechaEm.Text = Format(rs!dDocFecha, "dd/mm/yyyy")
            gdFechaAnt = Format(rs!dDocFecha, "dd/mm/yyyy")
            txtConcepto.Text = rs!cDescrip
            cboMoneda.ListIndex = BuscaCombo(rs!nMoneda, cboMoneda)
            txtVvtaTotal.Text = Format(rs!nVVenta / IIf(nVal(rs!nTipoCambio) = 0, 1, rs!nTipoCambio), gsFormatoNumeroView)
            txtIGV.Text = Format(rs!nIGV / IIf(nVal(rs!nTipoCambio) = 0, 1, rs!nTipoCambio), gsFormatoNumeroView)
            txtPvtaTotal.Text = Format(rs!nPVenta / IIf(nVal(rs!nTipoCambio) = 0, 1, rs!nTipoCambio), gsFormatoNumeroView)
            chkIGV.value = IIf(nVal(rs!nIGV) = 0, 0, 1)
            txtTpoCambio = Format(IIf(IsNull(rs!nTipoCambio), 0, rs!nTipoCambio), "#0.000")
            If Len(rs!cPersCod) > 0 Then
                oPer.ObtieneClientexCodigo rs!cPersCod
                txtProvCod.Tag = oPer.sPersCod
                lblProvNom.Caption = oPer.sPersNombre
                txtProvCod = oPer.sPersIdnroRUC
            End If
            If txtProvCod = "" Then txtProvCod = oPer.sPersIdnroDNI
            Me.Caption = "Modificación de Comprobante de Pago Autorizado"
        End If
    End If
End Sub
Private Sub mebFechaEm_GotFocus()
    fEnfoque mebFechaEm
End Sub
Private Sub mebFechaEm_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 And Not IsDate(mebFechaEm.Text) Then MsgBox "Ingrese una fecha válida", vbInformation, "¡Aviso!": Exit Sub
    If KeyAscii = 13 And Not Len(mebFechaEm.Text) = 0 Then cboMoneda.SetFocus
End Sub
Private Sub txtProvCod_EmiteDatos()
    If Len(txtProvCod.Text) = 0 Then lblProvNom.Caption = "": Exit Sub
    txtProvCod.Tag = txtProvCod.Text
    lblProvNom.Caption = txtProvCod.psDescripcion
    txtProvCod.Text = txtProvCod.sPersNroDoc
    txtConcepto.SetFocus
End Sub
Private Sub txtProvCod_GotFocus()
    fEnfoque txtProvCod
End Sub
Private Sub txtProvCod_LostFocus()
    If Len(txtProvCod.Text) = 0 Then lblProvNom.Caption = ""
End Sub
Private Sub txtPvtaTotal_GotFocus()
    fEnfoque txtPvtaTotal
End Sub
Private Sub txtPvtaTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 And Not Len(txtPvtaTotal.Text) = 0 Then cmdAceptar.SetFocus
End Sub
Private Sub txtPvtaTotal_KeyUp(KeyCode As Integer, Shift As Integer)
    txtVvtaTotal.Text = Format(Round(nVal(txtPvtaTotal.Text) / (1 + IIf(chkIGV.value = 0, 0, nTasaIGV)), 2), gsFormatoNumeroView)
    txtIGV = Format(nVal(txtPvtaTotal.Text) - nVal(txtVvtaTotal.Text), gsFormatoNumeroView)
    txtPvtaMN.Text = Format(nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), gsFormatoNumeroView)
End Sub
Private Sub txtPvtaTotal_LostFocus()
    If Trim(txtPvtaTotal.Text) = "" Then txtPvtaTotal.Text = "0.00"
    txtPvtaTotal.Text = Format(txtPvtaTotal.Text, "#0.00")
End Sub
Private Sub txtTpoCambio_Change()
    txtPvtaMN.Text = Format(nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), gsFormatoNumeroView)
End Sub
Private Sub txtTpoCambio_GotFocus()
    fEnfoque txtTpoCambio
End Sub
Private Sub txtTpoCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 And Not Len(txtTpoCambio.Text) = 0 Then txtVvtaTotal.SetFocus
End Sub
Private Sub txtTpoCambio_LostFocus()
    If Trim(txtTpoCambio.Text) = "" Then txtTpoCambio.Text = "0.000"
    txtTpoCambio.Text = Format(txtTpoCambio.Text, "#0.000")
End Sub
Private Sub txtVvtaTotal_GotFocus()
    fEnfoque txtVvtaTotal
End Sub
Private Sub txtVvtaTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = TextBox_SoloNumerosDecimales(KeyAscii)
    If KeyAscii = 13 And Not Len(txtVvtaTotal.Text) = 0 Then txtPvtaTotal.SetFocus
End Sub
Private Sub txtVvtaTotal_KeyUp(KeyCode As Integer, Shift As Integer)
    txtIGV = Format(IIf(chkIGV.value = 0, 0, Round(nVal(txtVvtaTotal.Text) * nTasaIGV, 2)), gsFormatoNumeroView)
    txtPvtaTotal.Text = Format(nVal(txtVvtaTotal.Text) + nVal(txtIGV.Text), gsFormatoNumeroView)
    txtPvtaMN.Text = Format(nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), gsFormatoNumeroView)
End Sub
Private Sub txtVvtaTotal_LostFocus()
    If Trim(txtVvtaTotal.Text) = "" Then txtVvtaTotal.Text = "0.00"
    txtVvtaTotal.Text = Format(txtVvtaTotal.Text, "#0.00")
End Sub
Private Sub chkIGV_Click()
    txtIGV.Text = Format(IIf(chkIGV.value = 0, 0, Round(nVal(txtVvtaTotal.Text) * nTasaIGV, 2)), gsFormatoNumeroView)
    txtPvtaTotal.Text = Format(nVal(txtVvtaTotal.Text) + nVal(txtIGV.Text), gsFormatoNumeroView)
    txtPvtaMN.Text = Format(nVal(txtPvtaTotal.Text) * nVal(txtTpoCambio.Text), gsFormatoNumeroView)
End Sub
Private Sub cboMoneda_Click()
    txtTpoCambio.Text = "0.00": txtPvtaMN.Text = "0.00"
    txtVvtaTotal.Text = "0.00"
    chkIGV.value = 0: txtIGV.Text = "0.00"
    txtPvtaTotal.Text = "0.00"
    txtTpoCambio.Enabled = IIf(Right(cboMoneda, 1) = 1, False, True)
End Sub
