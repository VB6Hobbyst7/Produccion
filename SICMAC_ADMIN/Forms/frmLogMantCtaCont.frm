VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogMantCtaCont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MANTENIMIENTO DE CUENTAS CONTABLES"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmLogMantCtaCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   4260
      Left            =   60
      TabIndex        =   6
      Top             =   645
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   7514
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Cta Debe"
      TabPicture(0)   =   "frmLogMantCtaCont.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCon"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cta - Cta"
      TabPicture(1)   =   "frmLogMantCtaCont.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Objetos"
      TabPicture(2)   =   "frmLogMantCtaCont.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   -74925
         TabIndex        =   16
         Top             =   375
         Width           =   8130
         Begin VB.CommandButton cmdEliminarObj 
            Caption         =   "&Eliminar"
            Height          =   300
            Left            =   6930
            TabIndex        =   18
            Top             =   3360
            Width           =   1080
         End
         Begin VB.CommandButton cmdNuevoObj 
            Caption         =   "&Nuevo"
            Height          =   300
            Left            =   5820
            TabIndex        =   17
            Top             =   3360
            Width           =   1080
         End
         Begin Sicmact.FlexEdit flexObj 
            Height          =   3105
            Left            =   75
            TabIndex        =   19
            Top             =   195
            Width           =   7965
            _extentx        =   14049
            _extenty        =   5477
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Cod Bien-Descrip-Objeto-CtaCont"
            encabezadosanchos=   "380-1300-3000-1600-1200"
            font            =   "frmLogMantCtaCont.frx":091E
            font            =   "frmLogMantCtaCont.frx":0946
            font            =   "frmLogMantCtaCont.frx":096E
            font            =   "frmLogMantCtaCont.frx":0996
            font            =   "frmLogMantCtaCont.frx":09BE
            fontfixed       =   "frmLogMantCtaCont.frx":09E6
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            columnasaeditar =   "X-1-X-3-4"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-3-0"
            encabezadosalineacion=   "L-L-L-L-L"
            formatosedit    =   "0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            appearance      =   0
            colwidth0       =   375
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   -74925
         TabIndex        =   10
         Top             =   375
         Width           =   8130
         Begin VB.CommandButton cmdNuevoCta 
            Caption         =   "&Nuevo"
            Height          =   300
            Left            =   5850
            TabIndex        =   12
            Top             =   3360
            Width           =   1080
         End
         Begin VB.CommandButton cmdEliminarCta 
            Caption         =   "&Eliminar"
            Height          =   300
            Left            =   6960
            TabIndex        =   11
            Top             =   3360
            Width           =   1080
         End
         Begin Sicmact.FlexEdit flexCtaCta 
            Height          =   3105
            Left            =   75
            TabIndex        =   13
            Top             =   195
            Width           =   7980
            _extentx        =   13097
            _extenty        =   5477
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Cod Bien-Descrip-Objeto-Cta Cont"
            encabezadosanchos=   "400-1400-2900-1400-1400"
            font            =   "frmLogMantCtaCont.frx":0A0C
            font            =   "frmLogMantCtaCont.frx":0A34
            font            =   "frmLogMantCtaCont.frx":0A5C
            font            =   "frmLogMantCtaCont.frx":0A84
            font            =   "frmLogMantCtaCont.frx":0AAC
            fontfixed       =   "frmLogMantCtaCont.frx":0AD4
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            columnasaeditar =   "X-1-2-3-4"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-3-0"
            encabezadosalineacion=   "L-L-L-L-L"
            formatosedit    =   "0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            appearance      =   0
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
      Begin VB.Frame fraCon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   75
         TabIndex        =   7
         Top             =   375
         Width           =   8130
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   300
            Left            =   6930
            TabIndex        =   9
            Top             =   3360
            Width           =   1080
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   300
            Left            =   5820
            TabIndex        =   8
            Top             =   3360
            Width           =   1080
         End
         Begin Sicmact.FlexEdit flex 
            Height          =   3105
            Left            =   75
            TabIndex        =   14
            Top             =   195
            Width           =   7965
            _extentx        =   13097
            _extenty        =   5477
            cols0           =   4
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Cod Bien-Descrip-Cta Debe"
            encabezadosanchos=   "400-1200-4100-1200"
            font            =   "frmLogMantCtaCont.frx":0AFA
            font            =   "frmLogMantCtaCont.frx":0B22
            font            =   "frmLogMantCtaCont.frx":0B4A
            font            =   "frmLogMantCtaCont.frx":0B72
            font            =   "frmLogMantCtaCont.frx":0B9A
            fontfixed       =   "frmLogMantCtaCont.frx":0BC2
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            columnasaeditar =   "X-1-X-3"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-0"
            encabezadosalineacion=   "L-L-L-R"
            formatosedit    =   "0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbpuntero       =   -1
            lbbuscaduplicadotext=   -1
            appearance      =   0
            colwidth0       =   405
            rowheight0      =   300
            forecolorfixed  =   -2147483630
         End
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   1275
      TabIndex        =   3
      Top             =   4950
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7260
      TabIndex        =   2
      Top             =   4950
      Width           =   1110
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   105
      TabIndex        =   1
      Top             =   4950
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   105
      TabIndex        =   0
      Top             =   4950
      Width           =   1110
   End
   Begin Sicmact.TxtBuscar txOperacion 
      Height          =   330
      Left            =   60
      TabIndex        =   4
      Top             =   255
      Width           =   1605
      _extentx        =   2831
      _extenty        =   582
      appearance      =   0
      appearance      =   0
      backcolor       =   12648447
      font            =   "frmLogMantCtaCont.frx":0BE8
      appearance      =   0
      stitulo         =   ""
   End
   Begin VB.Label lblOpe 
      Caption         =   "Operacion :"
      Height          =   225
      Left            =   90
      TabIndex        =   15
      Top             =   30
      Width           =   1230
   End
   Begin VB.Label lblOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   270
      Width           =   6645
   End
End
Attribute VB_Name = "frmLogMantCtaCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdCancelar_Click()
    Activa False
    txOperacion.Text = ""
    txOperacion_EmiteDatos
End Sub

Private Sub cmdEditar_Click()
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro numero : " & Me.flex.row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    flex.EliminaFila flex.row
End Sub

Private Sub cmdEliminarCta_Click()
    If MsgBox("Desea Eliminar el registro numero : " & Me.flexCtaCta.row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    flexCtaCta.EliminaFila flexCtaCta.row
End Sub

Private Sub cmdEliminarObj_Click()
    If MsgBox("Desea Eliminar el registro numero : " & Me.flexObj.row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    flexObj.EliminaFila flexObj.row
End Sub

Private Sub cmdGrabar_Click()
    Dim oCta As DLogCtaCont
    Set oCta = New DLogCtaCont
    
    If Not Valida Then Exit Sub
    
    oCta.SetCtaBS txOperacion.Text, flex.GetRsNew, Me.flexCtaCta.GetRsNew, Me.flexObj.GetRsNew, GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        
        'ARLO 20160126 ***
        Dim nNum As Integer
        nNum = Me.flex.Rows - 1
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Grabo el Codigo de Operacion : " & txOperacion & " de la Cuenta Debe " & Me.flex.TextMatrix(nNum, 3)
        Set objPista = Nothing
        '**************
    
    cmdCancelar_Click
End Sub

Private Sub cmdNuevo_Click()
    flex.AdicionaFila
End Sub

Private Sub cmdNuevoCta_Click()
    Me.flexCtaCta.AdicionaFila
End Sub

Private Sub cmdNuevoObj_Click()
    flexObj.AdicionaFila
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    Dim rsBS As ADODB.Recordset
    Set rsBS = New ADODB.Recordset
    Dim oObj As DObjeto
    Set oObj = New DObjeto
    Dim rsObj As ADODB.Recordset
    Set rsObj = New ADODB.Recordset
    
    Set rsBS = oALmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "", True)
    Set rsObj = oObj.CargaObjetoCombo("BAF", True)
    Me.flex.rsTextBuscar = rsBS
    Me.flexObj.rsTextBuscar = rsBS
    Me.txOperacion.rs = oOpe.GetOperaciones("5%")
    flexObj.CargaCombo rsObj
    
    Set oOpe = Nothing
    Set oALmacen = Nothing
    Set rsBS = Nothing
    Set rsObj = Nothing
    Set rsObj = Nothing
    Activa False
    SSTab.TabVisible(2) = False
End Sub


Private Sub Activa(pbEditar As Boolean)
    Me.txOperacion.Enabled = Not pbEditar
    Me.flex.lbEditarFlex = pbEditar
    Me.flexCtaCta.lbEditarFlex = pbEditar
    Me.cmdCancelar.Visible = pbEditar
    Me.cmdEditar.Visible = Not pbEditar
    Me.cmdGrabar.Enabled = pbEditar
    Me.SSTab.Tab = 0
End Sub

Private Function Valida() As Boolean
    Dim lnI As Integer
    
    If Me.txOperacion.Text = "" Then
        Valida = False
        MsgBox "Debe hacer referencia a una planilla valida.", vbInformation, "Aviso"
        Exit Function
    End If
    
    For lnI = 1 To Me.flex.Rows - 1
        flex.row = lnI
        If flex.TextMatrix(lnI, 1) = "" And flex.TextMatrix(lnI, 0) <> "" Then
            flex.col = 1
            MsgBox "Debe ingresar un codigo valido.", vbInformation, "Aviso"
            flex.SetFocus
            Exit Function
        ElseIf flex.TextMatrix(lnI, 3) = "" And flex.TextMatrix(lnI, 0) <> "" Then
            flex.col = 3
            MsgBox "Debe ingresar una cuenta contable valida, para el debe o el haber.", vbInformation, "Aviso"
            flex.SetFocus
            Exit Function
        End If
    Next lnI
    
    For lnI = 1 To Me.flexCtaCta.Rows - 1
        flexCtaCta.row = lnI
        If flexCtaCta.TextMatrix(lnI, 1) = "" And flexCtaCta.TextMatrix(lnI, 0) <> "" Then
            flexCtaCta.col = 1
            MsgBox "Debe ingresar una cuenta contable valida.", vbInformation, "Aviso"
            flexCtaCta.SetFocus
            Exit Function
        ElseIf flexCtaCta.TextMatrix(lnI, 2) = "" And flexCtaCta.TextMatrix(lnI, 0) <> "" Then
            flexCtaCta.col = 3
            MsgBox "Debe ingresar una cuenta contable valida.", vbInformation, "Aviso"
            flexCtaCta.SetFocus
            Exit Function
        End If
    Next lnI
    
    Valida = True
End Function

Private Sub txOperacion_EmiteDatos()
    Dim oCta As DLogCtaCont
    Set oCta = New DLogCtaCont
        
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim rsO As ADODB.Recordset
    Set rsO = New ADODB.Recordset

    Me.lblOperacion.Caption = txOperacion.psDescripcion
    
    If Me.lblOperacion.Caption = "" Then
        flex.Clear
        flex.Rows = 2
        flex.FormaCabecera
    
        flexCtaCta.Clear
        flexCtaCta.Rows = 2
        flexCtaCta.FormaCabecera
    Else
        Set rs = oCta.GetConceptoCtaDeb(Me.txOperacion.Text)
        Set rsC = oCta.GetConceptoCtaCta(Me.txOperacion.Text)
       ' Set rsO = oCta.GetConceptoObj(Me.txOperacion.Text)
        
        If rs.EOF And rs.BOF Then
            flex.Clear
            flex.Rows = 2
            flex.FormaCabecera
        Else
            flex.rsFlex = rs
        End If
        
        If rsC.EOF And rsC.BOF Then
            flexCtaCta.Clear
            flexCtaCta.Rows = 2
            flexCtaCta.FormaCabecera
        Else
            flexCtaCta.rsFlex = rsC
        End If
        
'        If rsO.EOF And rsO.BOF Then
'            flexObj.Clear
'            flexObj.Rows = 2
'            flexObj.FormaCabecera
'        Else
'            flexObj.rsFlex = rsO
'        End If
    
    
    
    End If
End Sub
