VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.OCX"
Begin VB.Form frmLogMantCtaCont 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmLogMantCtaCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   4260
      Left            =   60
      TabIndex        =   6
      Top             =   405
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   7514
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Debe"
      TabPicture(0)   =   "frmLogMantCtaCont.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCon"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cta - Cta"
      TabPicture(1)   =   "frmLogMantCtaCont.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   75
         TabIndex        =   10
         Top             =   375
         Width           =   7560
         Begin VB.CommandButton cmdNuevoCta 
            Caption         =   "&Nuevo"
            Height          =   300
            Left            =   5310
            TabIndex        =   12
            Top             =   3360
            Width           =   1080
         End
         Begin VB.CommandButton cmdEliminarCta 
            Caption         =   "&Eliminar"
            Height          =   300
            Left            =   6420
            TabIndex        =   11
            Top             =   3360
            Width           =   1080
         End
         Begin Sicmact.FlexEdit flexCtaCta 
            Height          =   3105
            Left            =   75
            TabIndex        =   13
            Top             =   195
            Width           =   7425
            _extentx        =   13097
            _extenty        =   5477
            cols0           =   5
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-CtaCont Debe-CtaCont Haber-CtaCont Debe Otro-CtaCont Haber Otro"
            encabezadosanchos=   "400-1600-1600-1600-1600"
            font            =   "frmLogMantCtaCont.frx":0342
            font            =   "frmLogMantCtaCont.frx":036A
            font            =   "frmLogMantCtaCont.frx":0392
            font            =   "frmLogMantCtaCont.frx":03BA
            font            =   "frmLogMantCtaCont.frx":03E2
            fontfixed       =   "frmLogMantCtaCont.frx":040A
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1  'True
            columnasaeditar =   "X-1-2-3-4"
            textstylefixed  =   3
            listacontroles  =   "0-0-0-0-0"
            encabezadosalineacion=   "L-L-L-L-L"
            formatosedit    =   "0-0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbbuscaduplicadotext=   -1  'True
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
         Left            =   -74925
         TabIndex        =   7
         Top             =   375
         Width           =   7560
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   300
            Left            =   6420
            TabIndex        =   9
            Top             =   3360
            Width           =   1080
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   300
            Left            =   5310
            TabIndex        =   8
            Top             =   3360
            Width           =   1080
         End
         Begin Sicmact.FlexEdit flex 
            Height          =   3105
            Left            =   75
            TabIndex        =   14
            Top             =   195
            Width           =   7425
            _extentx        =   13097
            _extenty        =   5477
            cols0           =   4
            highlight       =   1
            allowuserresizing=   3
            rowsizingmode   =   1
            encabezadosnombres=   "#-Cod Concepto-Concepto Remunerativo-Debe"
            encabezadosanchos=   "400-1200-4100-1200"
            font            =   "frmLogMantCtaCont.frx":0430
            font            =   "frmLogMantCtaCont.frx":0458
            font            =   "frmLogMantCtaCont.frx":0480
            font            =   "frmLogMantCtaCont.frx":04A8
            font            =   "frmLogMantCtaCont.frx":04D0
            fontfixed       =   "frmLogMantCtaCont.frx":04F8
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            columnasaeditar =   "X-1-X-3"
            textstylefixed  =   3
            listacontroles  =   "0-1-0-0"
            encabezadosalineacion=   "L-L-L-R"
            formatosedit    =   "0-0-0-0"
            textarray0      =   "#"
            lbeditarflex    =   -1  'True
            lbpuntero       =   -1  'True
            lbbuscaduplicadotext=   -1  'True
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
      Left            =   1230
      TabIndex        =   3
      Top             =   4725
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   6645
      TabIndex        =   2
      Top             =   4725
      Width           =   1110
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   345
      Left            =   60
      TabIndex        =   1
      Top             =   4725
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   4725
      Width           =   1110
   End
   Begin Sicmact.TxtBuscar txOperacion 
      Height          =   330
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   1605
      _extentx        =   2831
      _extenty        =   582
      appearance      =   0
      appearance      =   0
      backcolor       =   12648447
      font            =   "frmLogMantCtaCont.frx":051E
      appearance      =   0
      stitulo         =   ""
   End
   Begin VB.Label lblOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   45
      Width           =   6030
   End
End
Attribute VB_Name = "frmLogMantCtaCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
    Activa False
    txOperacion.Text = ""
    txOperacion_EmiteDatos
End Sub

Private Sub cmdEditar_Click()
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro numero : " & Me.Flex.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Flex.EliminaFila Flex.Row
End Sub

Private Sub cmdEliminarCta_Click()
    If MsgBox("Desea Eliminar el registro numero : " & Me.flexCtaCta.Row, vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    flexCtaCta.EliminaFila flexCtaCta.Row
End Sub

Private Sub cmdGrabar_Click()
    Dim oCta As DLogCtaCont
    Set oCta = New DLogCtaCont
    
    If Not Valida Then Exit Sub
    
    oCta.SetCtaBS txOperacion.Text, Flex.GetRsNew, Me.flexCtaCta.GetRsNew
    
    cmdCancelar_Click
End Sub

Private Sub cmdNuevo_Click()
    Flex.AdicionaFila
End Sub

Private Sub cmdNuevoCta_Click()
    Me.flexCtaCta.AdicionaFila
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oAlmacen As DLogAlmacen
    Set oAlmacen = New DLogAlmacen
    Dim oOpe As DOperaciones
    Set oOpe = New DOperaciones
    
    Me.Flex.rsTextBuscar = oAlmacen.GetBienesAlmacen(, "" & gnLogBSTpoBienConsumo & "','" & gnLogBSTpoBienFijo & "','" & gnLogBSTpoBienNoDepreciable & "")
    Me.txOperacion.rs = oOpe.GetOperaciones("5%")
    
    Set oOpe = Nothing
    Set oAlmacen = Nothing
    Activa False
End Sub


Private Sub Activa(pbEditar As Boolean)
    Me.txOperacion.Enabled = Not pbEditar
    Me.Flex.lbEditarFlex = pbEditar
    Me.flexCtaCta.lbEditarFlex = pbEditar
    Me.CmdCancelar.Visible = pbEditar
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
    
    For lnI = 1 To Me.Flex.Rows - 1
        Flex.Row = lnI
        If Flex.TextMatrix(lnI, 1) = "" And Flex.TextMatrix(lnI, 0) <> "" Then
            Flex.Col = 1
            MsgBox "Debe ingresar un codigo valido.", vbInformation, "Aviso"
            Flex.SetFocus
            Exit Function
        ElseIf Flex.TextMatrix(lnI, 3) = "" And Flex.TextMatrix(lnI, 0) <> "" Then
            Flex.Col = 3
            MsgBox "Debe ingresar una cuenta contable valida, para el debe o el haber.", vbInformation, "Aviso"
            Flex.SetFocus
            Exit Function
        End If
    Next lnI
    
    For lnI = 1 To Me.flexCtaCta.Rows - 1
        flexCtaCta.Row = lnI
        If flexCtaCta.TextMatrix(lnI, 1) = "" And flexCtaCta.TextMatrix(lnI, 0) <> "" Then
            flexCtaCta.Col = 1
            MsgBox "Debe ingresar una cuenta contable valida.", vbInformation, "Aviso"
            flexCtaCta.SetFocus
            Exit Function
        ElseIf flexCtaCta.TextMatrix(lnI, 2) = "" And flexCtaCta.TextMatrix(lnI, 0) <> "" Then
            flexCtaCta.Col = 3
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

    Me.lblOperacion.Caption = txOperacion.psDescripcion
    
    If Me.lblOperacion.Caption = "" Then
        Flex.Clear
        Flex.Rows = 2
        Flex.FormaCabecera
    
        flexCtaCta.Clear
        flexCtaCta.Rows = 2
        flexCtaCta.FormaCabecera
    Else
        Set rs = oCta.GetConceptoCtaDeb(Me.txOperacion.Text)
        Set rsC = oCta.GetConceptoCtaCta(Me.txOperacion.Text)
        
        If rs.EOF And rs.BOF Then
            Flex.Clear
            Flex.Rows = 2
            Flex.FormaCabecera
        Else
            Flex.rsFlex = rs
        End If
        
        If rsC.EOF And rsC.BOF Then
            flexCtaCta.Clear
            flexCtaCta.Rows = 2
            flexCtaCta.FormaCabecera
        Else
            flexCtaCta.rsFlex = rsC
        End If
    End If
End Sub
