VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCredGastosXProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro de Gastos por Tipo de Crédito"
   ClientHeight    =   6225
   ClientLeft      =   2415
   ClientTop       =   1590
   ClientWidth     =   7425
   Icon            =   "frmCredGastosXProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   5970
      TabIndex        =   14
      Top             =   5715
      Width           =   1350
   End
   Begin TabDlg.SSTab SSFiltro 
      Height          =   4680
      Left            =   75
      TabIndex        =   6
      Top             =   990
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   8255
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Opciones"
      TabPicture(0)   =   "frmCredGastosXProd.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Resultado"
      TabPicture(1)   =   "frmCredGastosXProd.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LstResul"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CmdEliminar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "CmdGrabar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar Resultados"
         Height          =   435
         Left            =   1770
         TabIndex        =   13
         Top             =   4005
         Width           =   1725
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   435
         Left            =   165
         TabIndex        =   12
         Top             =   4005
         Width           =   1575
      End
      Begin MSComctlLib.ListView LstResul 
         Height          =   3330
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   5874
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod.TipCred"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo Crédito"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "VMon"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "cCodAge"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Agencia"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodIntitucion"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Institucion"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Frame Frame2 
         Height          =   3990
         Left            =   -74880
         TabIndex        =   7
         Top             =   435
         Width           =   7080
         Begin VB.CheckBox ChkIntitucion 
            Caption         =   "Intitucion"
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
            Left            =   5370
            TabIndex        =   18
            Top             =   3030
            Width           =   1455
         End
         Begin VB.CommandButton CmdInstitucion 
            Caption         =   "Intitucion"
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
            Height          =   375
            Left            =   3510
            TabIndex        =   17
            Top             =   2970
            Width           =   1635
         End
         Begin VB.CheckBox Chkprod 
            Caption         =   "Todos"
            Height          =   240
            Left            =   2505
            TabIndex        =   15
            Top             =   255
            Width           =   795
         End
         Begin VB.Frame Frame3 
            Height          =   900
            Left            =   3465
            TabIndex        =   11
            Top             =   1950
            Width           =   3450
            Begin VB.CommandButton CmdAplicar 
               Caption         =   "&Aplicar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   660
               TabIndex        =   16
               Top             =   285
               Width           =   2190
            End
         End
         Begin VB.CheckBox ChkAgeTotal 
            Caption         =   "Todas"
            Height          =   240
            Left            =   6045
            TabIndex        =   10
            Top             =   255
            Width           =   795
         End
         Begin VB.ListBox LstAgencias 
            Height          =   1410
            Left            =   3465
            Style           =   1  'Checkbox
            TabIndex        =   1
            Top             =   525
            Width           =   3435
         End
         Begin VB.ListBox LstProd 
            Height          =   3210
            Left            =   165
            Style           =   1  'Checkbox
            TabIndex        =   0
            Top             =   525
            Width           =   3120
         End
         Begin VB.Label Label3 
            Caption         =   "Agencias:"
            Height          =   225
            Left            =   3495
            TabIndex        =   9
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Tipos de Créditos :"
            Height          =   225
            Left            =   120
            TabIndex        =   8
            Top             =   225
            Width           =   1755
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   450
      TabIndex        =   3
      Top             =   60
      Width           =   6435
      Begin VB.Label lblConcepto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C16A0B&
         Height          =   315
         Left            =   1095
         TabIndex        =   5
         Top             =   210
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmCredGastosXProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nConcepto As Long
Dim vnMoneda As Integer
Dim vbCargaGar As Boolean
Dim MatInstitucion As Variant

Public Sub Inicio(ByVal pnConcepto As Long, ByVal psDescripcion As String, _
    ByVal pnMoneda As Moneda, Optional ByVal pbGarantias As Boolean = False)
    vbCargaGar = pbGarantias
    
    nConcepto = pnConcepto
    lblConcepto.Caption = Space(2) & psDescripcion
    vnMoneda = pnMoneda
    Me.Show 1
End Sub

Public Sub CargaResultados()
Dim R As ADODB.Recordset
Dim oGasto As COMDCredito.DCOMGasto
Dim L As ListItem
Dim i, J As Integer
Dim rs As ADODB.Recordset

    Set oGasto = New COMDCredito.DCOMGasto
    Set R = oGasto.RecuperaFiltros(nConcepto, IIf(vbCargaGar, "G", "P"))
    Set oGasto = Nothing
    LstResul.ListItems.Clear
    Do While Not R.EOF
        Set L = LstResul.ListItems.Add(, , R!nProdCod)
        L.SubItems(1) = Trim(R!cConsDescripcion)
        L.SubItems(2) = vnMoneda
        L.SubItems(3) = IIf(vnMoneda = gMonedaNacional, "SOLES", "DOLARES")
        L.SubItems(4) = Trim(R!cAgeCod)
        L.SubItems(5) = Trim(R!cAgeDescripcion)
        'ARCV 01-02-2007
        'Set oGasto = New COMDCredito.DCOMGasto
        'Set rs = oGasto.RecupCodInstitucion(nConcepto, R!nProdCod, R!cAgeCod)
        'If Not rs.EOF Then
        '    L.SubItems(6) = rs!cPersCod
        '    L.SubItems(7) = rs!cPersNombre
        'End If
        If R!cIntitucion <> "" Then
            L.SubItems(6) = R!cIntitucion
            L.SubItems(7) = R!cPersNombre
        End If
        R.MoveNext
        '---------------
    Loop
    'Checkea los producto y agencias que ya estan grabados
    If R.RecordCount > 0 Then
        R.MoveFirst
        Do While Not R.EOF
            For J = 0 To LstProd.ListCount - 1
                If Trim(R!nProdCod) = Trim(Right(LstProd.List(J), 6)) Then
                    LstProd.Selected(J) = True
                End If
            Next J
            For J = 0 To LstAgencias.ListCount - 1
                If Trim(R!cAgeCod) = Trim(Right(LstAgencias.List(J), 6)) Then
                    LstAgencias.Selected(J) = True
                End If
            Next J
            R.MoveNext
        Loop
    End If
    
End Sub

Private Sub CargaControles()
Dim oCred As COMDCredito.DCOMCredito
Dim oGasto As COMDCredito.DCOMGasto
Dim oConst As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    If Not vbCargaGar Then
        'Carga Productos
        Set oCred = New COMDCredito.DCOMCredito
        Set R = oCred.RecuperaProductosDeCredito
        Set oCred = Nothing
        LstProd.Clear
        Do While Not R.EOF
            LstProd.AddItem Trim(R!cConsDescripcion) & Space(100) & Trim(R!nConsValor)
            R.MoveNext
        Loop
        R.Close
    Else
        Set oConst = New COMDConstantes.DCOMConstantes
        Set R = oConst.RecuperaConstantes(gPersGarantia)
        Set oConst = Nothing
        LstProd.Clear
        Do While Not R.EOF
            LstProd.AddItem Trim(R!cConsDescripcion) & Space(100) & Trim(R!nConsValor)
            R.MoveNext
        Loop
        R.Close
    End If
    
    'Carga Agencias
    LstAgencias.Clear
    Set oGasto = New COMDCredito.DCOMGasto
    Set R = oGasto.RecuperaAgencias
    Do While Not R.EOF
        LstAgencias.AddItem Trim(R!cAgeDescripcion) & Space(100) & Trim(R!cAgeCod)
        R.MoveNext
    Loop
    Set oGasto = Nothing

    
    
End Sub

Private Function ExisteFiltro(ByVal psProdCod As String, ByVal psMoneda As String, ByVal psAgeCod As String, _
Optional ByVal pbIntitucion As Boolean = False, Optional ByVal psInstitucion As String = "") As Boolean
Dim i As Integer
    
        ExisteFiltro = False
        For i = 1 To LstResul.ListItems.Count
            If pbIntitucion = False Then
                If LstResul.ListItems(i).Text = psProdCod And _
                     LstResul.ListItems(i).SubItems(2) = psMoneda And _
                     LstResul.ListItems(i).SubItems(4) = psAgeCod Then
                     ExisteFiltro = True
                    Exit For
                End If
            Else
                If LstResul.ListItems(i).Text = psProdCod And _
                     LstResul.ListItems(i).SubItems(2) = psMoneda And _
                     LstResul.ListItems(i).SubItems(4) = psAgeCod And _
                     LstResul.ListItems(i).SubItems(6) = psInstitucion Then
                     ExisteFiltro = True
                    Exit For
                End If
            End If
        Next i
    
End Function


Private Sub ChkAgeTotal_Click()
Dim i As Integer
    If ChkAgeTotal.value = 1 Then
        For i = 0 To LstAgencias.ListCount - 1
            LstAgencias.Selected(i) = True
        Next i
    Else
        For i = 0 To LstAgencias.ListCount - 1
            LstAgencias.Selected(i) = False
        Next i
    End If
End Sub

Private Sub ChkIntitucion_Click()
    If ChkIntitucion.value = 1 Then
        'verificar que sea solo descuento por planilla
        If LstProd.ListIndex <> -1 Then
            'If Trim(Right(LstProd.List(LstProd.ListIndex), 10)) = "301" Then
            'MAVM 20100703 BAS II
            If Trim(Right(LstProd.List(LstProd.ListIndex), 10)) = "751" Then
                CmdInstitucion.Enabled = True
            Else
               MsgBox "Esta Opcion solo es para descuento por planilla", vbInformation, "AVISO"
            End If
        Else
            MsgBox "Debe seleccionar un producto", vbInformation, "AVISO"
        End If
    Else
        CmdInstitucion.Enabled = False
    End If
End Sub

Private Sub Chkprod_Click()
Dim i As Integer
    If Chkprod.value = 1 Then
        For i = 0 To LstProd.ListCount - 1
            LstProd.Selected(i) = True
        Next i
    Else
        For i = 0 To LstProd.ListCount - 1
            LstProd.Selected(i) = False
        Next i
    End If
End Sub

Private Sub cmdAplicar_Click()
Dim L As ListItem
Dim i As Integer
Dim J As Integer
Dim K As Integer

Dim bInstitucion As Boolean
Dim sInstitucion As String
Dim sNomInstitucion As String
Dim objDGastos As COMDCredito.DCOMGasto
    
    If ChkIntitucion.value = 1 Then
       bInstitucion = True
       'sInstitucion = MatInstitucion(0)
    Else
       bInstitucion = False
       sInstitucion = ""
    End If
    
'ARCV 31-01-2007
    LstResul.ListItems.Clear
    If bInstitucion Then
            For K = 0 To UBound(MatInstitucion) - 1
                For i = 0 To LstProd.ListCount - 1
                    If LstProd.Selected(i) Then
                        For J = 0 To LstAgencias.ListCount - 1
                            If LstAgencias.Selected(J) Then
                        
                                Set L = LstResul.ListItems.Add(, , Right(LstProd.List(i), 3))
                                L.SubItems(1) = Trim(Left(LstProd.List(i), Len(LstProd.List(i)) - 4))
                                L.SubItems(2) = Trim(str(vnMoneda))
                                L.SubItems(3) = Trim(IIf(vnMoneda = gMonedaNacional, "SOLES", "DOLARES"))
                                L.SubItems(4) = Right(LstAgencias.List(J), 2)
                                L.SubItems(5) = Trim(Left(LstAgencias.List(J), Len(LstAgencias.List(J)) - 3))
                                L.SubItems(6) = MatInstitucion(K + 1, 1)
                                L.SubItems(7) = MatInstitucion(K + 1, 2)
                            End If
                        Next J
                    End If
                Next i
            Next K
Else
'-----------
    
    Set objDGastos = New COMDCredito.DCOMGasto
    sNomInstitucion = objDGastos.RecupNomInstitucion(sInstitucion)
    Set objDGastos = Nothing

    LstResul.ListItems.Clear
    For i = 0 To LstProd.ListCount - 1
        If LstProd.Selected(i) Then
            For J = 0 To LstAgencias.ListCount - 1
                If LstAgencias.Selected(J) Then
                        If Not ExisteFiltro(Right(LstProd.List(i), 3), Trim(str(vnMoneda)), Right(LstAgencias.List(J), 2), bInstitucion, sInstitucion) Then
                            Set L = LstResul.ListItems.Add(, , Right(LstProd.List(i), 3))
                            L.SubItems(1) = Trim(Left(LstProd.List(i), Len(LstProd.List(i)) - 4))
                            L.SubItems(2) = Trim(str(vnMoneda))
                            L.SubItems(3) = Trim(IIf(vnMoneda = gMonedaNacional, "SOLES", "DOLARES"))
                            L.SubItems(4) = Right(LstAgencias.List(J), 2)
                            L.SubItems(5) = Trim(Left(LstAgencias.List(J), Len(LstAgencias.List(J)) - 3))
                            L.SubItems(6) = sInstitucion
'                            Set objDGastos = New COMDCredito.DCOMGasto
'                            sNomInstitucion = objDGastos.RecupNomInstitucion(sInstitucion)
'                            Set objDGastos = Nothing
                            L.SubItems(7) = sNomInstitucion
                        End If
                End If
            Next J
        End If
    Next i
    
End If  'ARCV 31-01-2007

    MsgBox "Los Filtros se ha Generado con Exito", vbInformation, "Mensaje"
End Sub

Private Sub cmdeliminar_Click()
Dim i As Integer
Dim bNoElim As Boolean

    If MsgBox("Se va Eliminar el Registro, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If

    bNoElim = False
    Do While Not bNoElim
        bNoElim = True
        For i = 1 To LstResul.ListItems.Count
            If LstResul.ListItems(i).Selected Then
                Call LstResul.ListItems.Remove(i)
                bNoElim = False
                Exit For
            End If
        Next i
    Loop
    
    
End Sub

Private Sub cmdGrabar_Click()
Dim oGasto As COMDCredito.DCOMGasto
Dim MatDatos() As String
Dim i As Integer
Dim oDCred As COMDCredito.DCOMCredActBD
    
    
    
    
    'If LstResul.ListItems.Count = 0 Then
    '    MsgBox "No existen datos para Grabar", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    
    If MsgBox("Se va a Grabar los Datos, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    Set oDCred = New COMDCredito.DCOMCredActBD
    Call oDCred.dDeleteProductoConceptoFitroTotal(nConcepto, IIf(vbCargaGar, "G", "P"))
    Set oDCred = Nothing
        
    ReDim MatDatos(LstResul.ListItems.Count, 7)
    For i = 1 To LstResul.ListItems.Count
        MatDatos(i - 1, 0) = Trim(str(nConcepto))
        MatDatos(i - 1, 1) = Trim(LstResul.ListItems(i).Text)
        MatDatos(i - 1, 2) = Trim(LstResul.ListItems(i).SubItems(1))
        MatDatos(i - 1, 3) = Trim(LstResul.ListItems(i).SubItems(2))
        MatDatos(i - 1, 4) = Trim(LstResul.ListItems(i).SubItems(3))
        MatDatos(i - 1, 5) = Trim(LstResul.ListItems(i).SubItems(4))
        MatDatos(i - 1, 6) = Trim(LstResul.ListItems(i).SubItems(5))
        MatDatos(i - 1, 7) = Trim(LstResul.ListItems(i).SubItems(6))
    Next i
    
    Set oGasto = New COMDCredito.DCOMGasto
    Call oGasto.ActualizaGastosProdFiltro(MatDatos, IIf(vbCargaGar, "G", "P"))
    Set oGasto = Nothing
        
    MsgBox "Datos Grabados", vbInformation, "Aviso"
End Sub

Private Sub CmdInstitucion_Click()
'ARCV 31-01-2007
'    Dim nContAge As Integer
'    Dim i As Integer
'    'If CmdInstitucion.Visible Then
'     frmSelectAnalistas.SeleccionaInstituciones
'        ReDim MatInstitucion(0)
'        nContAge = 0
'        If frmSelectAnalistas.LstAnalista.ListCount > 0 Then
'            For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
'                If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
'                    nContAge = nContAge + 1
'                    ReDim Preserve MatInstitucion(nContAge)
'                    MatInstitucion(nContAge - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
'                End If
'            Next i
'        End If
'        If UBound(MatInstitucion) = 0 Then
'            MsgBox "Seleccione una Institucion"
'            Exit Sub
'        End If
'    'End If
'------------
    Dim nContAge As Integer
    Dim i As Integer
    Dim nCont As Integer

        Call frmSelectAnalistas.SeleccionaInstituciones
        
        nCont = 0
        For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
            If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                nCont = nCont + 1
            End If
        Next i
        ReDim MatInstitucion(nCont, 2)
        nCont = 1
        For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
            If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                'MatInstitucion(nCont, 1) = Trim(frmSelectAnalistasECSA.lvAnalistas.ListItems(I).Tag)
                'MatInstitucion(nCont, 2) = Trim(frmSelectAnalistasECSA.lvAnalistas.ListItems(I).Text)
                'rs!cPersNombre & Space(100) & rs!cPersCod
                MatInstitucion(nCont, 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
                MatInstitucion(nCont, 2) = Trim(Mid(frmSelectAnalistas.LstAnalista.List(i), 1, Len(frmSelectAnalistas.LstAnalista.List(i)) - Len(Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))) - 100))
                nCont = nCont + 1
            End If
        Next i
        
        
        If UBound(MatInstitucion) = 0 Then
            MsgBox "Para continuar con el proceso, Tiene que seleccionar una Institucion ?", vbInformation, "Aviso"
            Exit Sub
        End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaControles
    Call CargaResultados
End Sub


Private Sub LstProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LstAgencias.SetFocus
    End If
End Sub


