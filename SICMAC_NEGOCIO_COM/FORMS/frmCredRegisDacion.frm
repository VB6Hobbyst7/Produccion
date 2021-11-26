VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmCredRegisDacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Dacion de Pago"
   ClientHeight    =   6705
   ClientLeft      =   2670
   ClientTop       =   1170
   ClientWidth     =   7110
   Icon            =   "frmCredRegisDacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   7035
      Begin VB.Frame Frame4 
         Caption         =   "Bienes de Dacion"
         Height          =   3780
         Left            =   150
         TabIndex        =   8
         Top             =   2010
         Width           =   6765
         Begin MSComctlLib.ListView LstDatosDet 
            Height          =   2865
            Left            =   135
            TabIndex        =   30
            Top             =   270
            Width           =   6420
            _ExtentX        =   11324
            _ExtentY        =   5054
            View            =   3
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
            Enabled         =   0   'False
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Bien"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cant"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Comercial"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Realizacion"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.TextBox TxtPrecRealiza 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   4875
            TabIndex        =   29
            Top             =   2640
            Width           =   1065
         End
         Begin VB.TextBox TxtPrecTasa 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1860
            TabIndex        =   27
            Top             =   2640
            Width           =   1065
         End
         Begin VB.TextBox TxtCant 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            TabIndex        =   25
            Top             =   2190
            Width           =   660
         End
         Begin VB.TextBox TxtBien 
            Enabled         =   0   'False
            Height          =   285
            Left            =   765
            TabIndex        =   23
            Top             =   2190
            Width           =   3960
         End
         Begin VB.CommandButton CmdAceptarDet 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   4350
            TabIndex        =   21
            Top             =   3300
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton CmdCancelarDet 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   5460
            TabIndex        =   20
            Top             =   3300
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   150
            TabIndex        =   19
            Top             =   3300
            Width           =   1080
         End
         Begin VB.CommandButton CmdEliminarDet 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1260
            TabIndex        =   18
            Top             =   3300
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Precio de Realizacion :"
            Height          =   195
            Left            =   3135
            TabIndex        =   28
            Top             =   2685
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Precio de Tasacion :"
            Height          =   195
            Left            =   315
            TabIndex        =   26
            Top             =   2685
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad :"
            Height          =   195
            Left            =   4860
            TabIndex        =   24
            Top             =   2235
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bien :"
            Height          =   195
            Left            =   285
            TabIndex        =   22
            Top             =   2220
            Width           =   405
         End
         Begin VB.Shape Shape1 
            Height          =   960
            Left            =   195
            Top             =   2070
            Width           =   6180
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de Dacion"
         Height          =   1755
         Left            =   120
         TabIndex        =   7
         Top             =   135
         Width           =   6795
         Begin VB.CommandButton CmdBuscarDacion 
            Caption         =   "&Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2640
            TabIndex        =   17
            Top             =   210
            Width           =   1215
         End
         Begin OcxLabelX.LabelX LblxCodCli 
            Height          =   420
            Left            =   675
            TabIndex        =   15
            Top             =   705
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   741
            FondoBlanco     =   0   'False
            Resalte         =   16711680
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX LblXNrodac 
            Height          =   465
            Left            =   1335
            TabIndex        =   14
            Top             =   210
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   820
            FondoBlanco     =   0   'False
            Resalte         =   16711680
            Bold            =   -1  'True
            Alignment       =   2
         End
         Begin VB.ComboBox CmbCredito 
            Enabled         =   0   'False
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1170
            Width           =   2580
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   315
            Left            =   6330
            TabIndex        =   10
            Top             =   740
            Width           =   375
         End
         Begin OcxLabelX.LabelX LblxNomCli 
            Height          =   420
            Left            =   2490
            TabIndex        =   16
            Top             =   705
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   741
            FondoBlanco     =   0   'False
            Resalte         =   16711680
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Dacion :"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   315
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Credito :"
            Height          =   195
            Left            =   75
            TabIndex        =   11
            Top             =   1215
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Codigo :"
            Height          =   195
            Left            =   75
            TabIndex        =   9
            Top             =   780
            Width           =   585
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   30
      TabIndex        =   0
      Top             =   5970
      Width           =   6855
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5700
         TabIndex        =   4
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4590
         TabIndex        =   3
         Top             =   225
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Anular"
         Height          =   375
         Left            =   1245
         TabIndex        =   2
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5700
         TabIndex        =   5
         Top             =   225
         Visible         =   0   'False
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmCredRegisDacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nAccion As Integer

Private Sub LimpiaDatosDet()
    TxtBien.Text = ""
    TxtCant.Text = ""
    TxtPrecRealiza.Text = ""
    TxtPrecTasa.Text = ""
End Sub

Private Sub LimpiaDatos()
    LstDatosDet.ListItems.Clear
    LblxCodCli.Caption = ""
    LblXNrodac.Caption = ""
    LblxNomCli.Caption = ""
    CmbCredito.Clear
    TxtBien.Text = ""
    TxtCant.Text = ""
    TxtPrecRealiza.Text = ""
    TxtPrecTasa.Text = ""
End Sub
Private Sub HabilitaDetalle(ByVal pbHabilita As Boolean)
    If pbHabilita Then
        LstDatosDet.Height = 1755
    Else
        LstDatosDet.Height = 2835
    End If
    cmdAgregar.Enabled = Not pbHabilita
    CmdEliminarDet.Enabled = Not pbHabilita
    CmdAceptarDet.Visible = pbHabilita
    CmdCancelarDet.Visible = pbHabilita
    TxtBien.Enabled = pbHabilita
    TxtCant.Enabled = pbHabilita
    TxtPrecTasa.Enabled = pbHabilita
    TxtPrecRealiza.Enabled = pbHabilita
    LstDatosDet.Enabled = Not pbHabilita
End Sub

Private Sub HabilitaNuevoEditar(ByVal pbHabilita As Boolean)
    CmdBuscarDacion.Enabled = Not pbHabilita
    CmbCredito.Enabled = pbHabilita
    CmdBuscar.Enabled = pbHabilita
    LstDatosDet.Enabled = pbHabilita
    TxtBien.Enabled = Not pbHabilita
    TxtCant.Enabled = Not pbHabilita
    TxtPrecTasa.Enabled = Not pbHabilita
    TxtPrecRealiza.Enabled = Not pbHabilita
    cmdAgregar.Enabled = pbHabilita
    CmdEliminarDet.Enabled = pbHabilita
    CmdNuevo.Enabled = Not pbHabilita
    CmdEditar.Enabled = Not pbHabilita

    CmdAceptar.Visible = pbHabilita
    CmdCancelar.Visible = pbHabilita
    CmdSalir.Visible = Not pbHabilita
End Sub

Private Sub CmbCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregar.SetFocus
    End If
End Sub

Private Sub CmdAceptar_Click()

Dim oBase As COMDCredito.DCOMCredActBD
'Dim nCodNew As Long
Dim nValorTotal As Double
Dim i As Integer
'Dim oFun As COMNContabilidad.NCOMContFunciones
'Dim sMovNro As String
'Dim nmovnro As Long

Dim MatDatos() As String
Dim nNroGarantRec As Long

    If Len(Trim(LblxCodCli.Caption)) <= 0 Then
        MsgBox "Seleccione un Titular para la Dacion", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(Trim(CmbCredito.Text)) <= 0 Then
        MsgBox "Seleccione el Credito para la Dacion en Pago", vbInformation, "Aviso"
        Exit Sub
    End If
    If LstDatosDet.ListItems.Count <= 0 Then
        MsgBox "Ingrese los Bienes para la Dacion en Pago", vbInformation, "Aviso"
        Exit Sub
    End If
    
 '   Set oBase = New COMDCredito.DCOMCredActBD
    ReDim MatDatos(LstDatosDet.ListItems.Count, 4)
    
    nValorTotal = 0#
    For i = 1 To LstDatosDet.ListItems.Count
        nValorTotal = nValorTotal + CDbl(LstDatosDet.ListItems(i).SubItems(3))
        
        MatDatos(i - 1, 0) = LstDatosDet.ListItems(i).Text
        MatDatos(i - 1, 1) = LstDatosDet.ListItems(i).SubItems(1)
        MatDatos(i - 1, 2) = LstDatosDet.ListItems(i).SubItems(2)
        MatDatos(i - 1, 3) = LstDatosDet.ListItems(i).SubItems(3)
    Next i
    nValorTotal = CDbl(Format(nValorTotal, "#0.00"))
    
    If LblXNrodac.Caption = "" Then
        nNroGarantRec = 0
    Else
        nNroGarantRec = CLng(LblXNrodac.Caption)
    End If
    
    Set oBase = New COMDCredito.DCOMCredActBD
    Call oBase.RegistrarDacionEnPago(nValorTotal, nAccion, gdFecSis, CmbCredito.Text, MatDatos, _
                                     nNroGarantRec, gsCodAge, gsCodUser)
    Set oBase = Nothing
    'Si es Nuevo
'    If nAccion = 1 Then
'        oBase.dBeginTrans
'        nCodNew = oBase.dNuevoCodigoColocGarantRec
'        Call oBase.dInsertColocGarantRec(nCodNew, gColocGarantRecDacion, CmbCredito.Text, gdFecSis, nValorTotal, gColocGarantRecEstadoRegistrado, False)
'        'Ingresa el Detalle
'        For I = 1 To LstDatosDet.ListItems.Count
'            Call oBase.dInsertColocGarantRecDet(nCodNew, CInt(LstDatosDet.ListItems(I).SubItems(1)), CDbl(LstDatosDet.ListItems(I).SubItems(2)), CDbl(LstDatosDet.ListItems(I).SubItems(3)), LstDatosDet.ListItems(I).Text, False)
'        Next I
'        Set oFun = New COMNContabilidad.NCOMContFunciones
'        sMovNro = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        Set oFun = Nothing
'        Call oBase.dInsertMov(sMovNro, gCredRegisDacion, "Registro de dacion de Pago", gMovEstContabMovContable, gMovFlagVigente, False)
'        nmovnro = oBase.dGetnMovNro(sMovNro)
'        Call oBase.dInsertMovCol(nmovnro, gCredRegisDacion, CmbCredito.Text, nCodNew, nValorTotal, 0, "", 0, 0, 0, False)
'        oBase.dCommitTrans
'    Else
'        'Si Es Modificacion
'        oBase.dBeginTrans
'        Call oBase.dUpdateColocGarantRec(CLng(LblXNrodac.Caption), CmbCredito.Text, nValorTotal, False)
'        Call oBase.dDeleteColocGarantRecDet(CLng(LblXNrodac.Caption), False)
'        For I = 1 To LstDatosDet.ListItems.Count
'            Call oBase.dInsertColocGarantRecDet(CLng(LblXNrodac.Caption), CInt(LstDatosDet.ListItems(I).SubItems(1)), CDbl(LstDatosDet.ListItems(I).SubItems(2)), CDbl(LstDatosDet.ListItems(I).SubItems(3)), LstDatosDet.ListItems(I).Text, False)
'        Next I
'        oBase.dCommitTrans
'    End If
'    Set oBase = Nothing
    
    HabilitaNuevoEditar False
    If nAccion = 1 Then
        LimpiaDatos
    End If
End Sub

Private Sub CmdAceptarDet_Click()
Dim L As ListItem
    Call TxtCant_LostFocus
    Call TxtPrecTasa_LostFocus
    Call TxtPrecRealiza_LostFocus
    If Len(Trim(TxtBien.Text)) <= 0 Or Len(Trim(TxtCant.Text)) <= 0 Or Len(Trim(TxtPrecTasa.Text)) <= 0 Or Len(Trim(TxtPrecRealiza.Text)) <= 0 Then
        MsgBox "Falta Ingresar un Dato", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If CDbl(TxtCant.Text) <= 0 Then
        MsgBox "Cantidad debe ser mayor a Cero", vbInformation, "Aviso"
        TxtCant.SetFocus
        Exit Sub
    End If
    If CDbl(TxtPrecTasa.Text) <= 0 Then
        MsgBox "Monto de Tasacion debe ser Mayor a Cero", vbInformation, "Aviso"
        TxtPrecTasa.SetFocus
        Exit Sub
    End If
    If CDbl(TxtPrecRealiza.Text) <= 0 Then
        MsgBox "Monto de Realiacion debe ser Mayor a Cero", vbInformation, "Aviso"
        TxtPrecRealiza.SetFocus
        Exit Sub
    End If
    'Valida Monto de Tasacion sea mayor a Monto Realizacion
    If CDbl(TxtPrecTasa.Text) < CDbl(TxtPrecRealiza.Text) Then
        MsgBox "Monto de Tasacion debe ser Mayor a Monto de Realizacion ", vbInformation, "Aviso"
        TxtPrecRealiza.SetFocus
        Exit Sub
    End If
    
    Set L = LstDatosDet.ListItems.Add(, , TxtBien.Text)
    L.SubItems(1) = TxtCant.Text
    L.SubItems(2) = TxtPrecTasa.Text
    L.SubItems(3) = TxtPrecRealiza.Text
    HabilitaDetalle False
End Sub

Private Sub CmdAgregar_Click()
    If CmbCredito.ListCount <= 0 Then
        MsgBox "No Existen Creditos para la Dacion", vbInformation, "Aviso"
        Exit Sub
    End If
    HabilitaDetalle True
    LimpiaDatosDet
    TxtBien.SetFocus
End Sub

Private Sub cmdBuscar_Click()
Dim oPers As COMDPersona.UCOMPersona
Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset

    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblxCodCli.Caption = oPers.sPerscod
        LblxNomCli.Caption = oPers.sPersNombre
        Set oDCred = New COMDCredito.DCOMCredito
        Set R = oDCred.RecuperaCreditosVigentesCtas(LblxCodCli.Caption)
        CmbCredito.Clear
        Do While Not R.EOF
            If Mid(R!cCtaCod, 6, 3) <> "402" Then
                CmbCredito.AddItem R!cCtaCod
            End If
            R.MoveNext
        Loop
        R.Close
        CmbCredito.SetFocus
    Else
        MsgBox "Cliente No Valido", vbInformation, "Aviso"
        
        Exit Sub
    End If
End Sub

Private Sub CargaDatos(ByVal pnDacion As Long)

Dim R As ADODB.Recordset
Dim RCred As ADODB.Recordset
Dim RDet As ADODB.Recordset
Dim oDCred As COMDCredito.DCOMCredito
Dim L As ListItem

    Set oDCred = New COMDCredito.DCOMCredito
    'Set R = oDCred.RecuperaDacionPago(pnDacion)
    Call oDCred.CargarDatosDacion(pnDacion, R, RCred, RDet)
    Set oDCred = Nothing
    
    If Not R.BOF And Not R.EOF Then
        LblXNrodac.Caption = pnDacion
        LblxCodCli.Caption = R!cPersCod
        LblxNomCli.Caption = R!cPersNombre
        'Set RCred = oDCred.RecuperaCreditosVigentesCtas(LblxCodCli.Caption)
        CmbCredito.Clear
        Do While Not RCred.EOF
            CmbCredito.AddItem RCred!cCtaCod
            RCred.MoveNext
        Loop
        RCred.Close
        Set RCred = Nothing
        CmbCredito.ListIndex = IndiceListaCombo(CmbCredito, R!cCtaCod, 3)
        
        'Set RDet = oDCred.RecuperaDacionPagoDetalle(pnDacion)
        LstDatosDet.ListItems.Clear
        Do While Not RDet.EOF
            Set L = LstDatosDet.ListItems.Add(, , RDet!cComentario)
            L.SubItems(1) = Trim(Str(RDet!nCantidad))
            L.SubItems(2) = Format(RDet!nPreTasacion, "#0.00")
            L.SubItems(3) = Format(RDet!nPreRealizacion, "#0.00")
            RDet.MoveNext
        Loop
        RDet.Close
        Set RDet = Nothing
    End If
    R.Close
    Set R = Nothing
End Sub

Private Sub CmdBuscarDacion_Click()
Dim nDacion As Long
    nDacion = frmCredPersEstado.BuscaDacionesPago("Daciones de Pago")
    If nDacion <> -1 Then
       Call CargaDatos(nDacion)
    Else
        HabilitaNuevoEditar False
        LimpiaDatos
    End If
End Sub

Private Sub cmdCancelar_Click()
    HabilitaNuevoEditar False
    If nAccion = 1 Then
        LimpiaDatos
    End If
End Sub

Private Sub CmdCancelarDet_Click()
    HabilitaDetalle False
End Sub

Private Sub cmdEditar_Click()
Dim oNCred As COMNCredito.NCOMCredito

    If Len(Trim(LblXNrodac.Caption)) <= 0 Or Len(Trim(CmbCredito.Text)) <= 0 Or LstDatosDet.ListItems.Count <= 0 Then
        MsgBox "Dacion no puede ser Eliminada", vbInformation, "Aviso"
        Exit Sub
    End If
    
    nAccion = 2
    If MsgBox("Se va a Anular el Registro de Dacion, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.AnularRegistroDacionEnPago(CLng(LblXNrodac.Caption), CmbCredito.Text)
    Set oNCred = Nothing
    Call LimpiaDatos
End Sub

Private Sub CmdEliminarDet_Click()
    If LstDatosDet.ListItems.Count > 0 Then
        If MsgBox("Se va a Eliminar el Bien : " & LstDatosDet.SelectedItem.Text & Chr(10) & " Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Call LstDatosDet.ListItems.Remove(LstDatosDet.SelectedItem.Index)
        End If
    Else
        MsgBox "No Existen Datos, para este Proceso", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdNuevo_Click()
    nAccion = 1
    HabilitaNuevoEditar True
    Call LimpiaDatos
    CmdBuscar.SetFocus
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub TxtBien_GotFocus()
    fEnfoque TxtBien
End Sub

Private Sub TxtBien_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtCant.SetFocus
    End If
End Sub

Private Sub TxtCant_GotFocus()
    fEnfoque TxtCant
End Sub

Private Sub TxtCant_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        TxtPrecTasa.SetFocus
    End If
End Sub

Private Sub TxtCant_LostFocus()
    If Trim(TxtCant.Text) = "" Then
        TxtCant.Text = "0"
    End If
End Sub

Private Sub TxtPrecRealiza_GotFocus()
    fEnfoque TxtPrecRealiza
End Sub

Private Sub TxtPrecRealiza_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPrecRealiza, KeyAscii)
    If KeyAscii = 13 Then
        CmdAceptarDet.SetFocus
    End If
End Sub

Private Sub TxtPrecRealiza_LostFocus()
    If Trim(TxtPrecRealiza.Text) = "" Then
        TxtPrecRealiza.Text = "0.00"
    End If
    TxtPrecRealiza.Text = Format(TxtPrecRealiza.Text, "#0.00")
End Sub

Private Sub TxtPrecTasa_GotFocus()
    fEnfoque TxtPrecTasa
End Sub

Private Sub TxtPrecTasa_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPrecTasa, KeyAscii)
    If KeyAscii = 13 Then
        TxtPrecRealiza.SetFocus
    End If
End Sub

Private Sub TxtPrecTasa_LostFocus()
    If Trim(TxtPrecTasa.Text) = "" Then
        TxtPrecTasa.Text = "0.00"
    End If
    TxtPrecTasa.Text = Format(TxtPrecTasa.Text, "#0.00")
End Sub
