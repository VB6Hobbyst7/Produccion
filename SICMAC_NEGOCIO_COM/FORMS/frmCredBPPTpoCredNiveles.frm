VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPTpoCredNiveles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Tipo de Créditos X Niveles"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "frmCredBPPTpoCredNiveles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTCategorias 
      CausesValidation=   0   'False
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9049
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parametros Tipo Créditos"
      TabPicture(0)   =   "frmCredBPPTpoCredNiveles.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAplicable"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGuardar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbNiveles"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.ComboBox cmbNiveles 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   11
         Top             =   2880
         Width           =   1170
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   2880
         Width           =   1170
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rangos de Cartera"
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
         Height          =   1575
         Left            =   5400
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
         Begin SICMACT.EditMoney txtMinCartera 
            Height          =   300
            Left            =   960
            TabIndex        =   6
            Top             =   480
            Width           =   1575
            _extentx        =   2778
            _extenty        =   529
            font            =   "frmCredBPPTpoCredNiveles.frx":0326
            text            =   "0"
            enabled         =   -1
         End
         Begin SICMACT.EditMoney txtMaxCartera 
            Height          =   300
            Left            =   960
            TabIndex        =   7
            Top             =   840
            Width           =   1575
            _extentx        =   2778
            _extenty        =   529
            font            =   "frmCredBPPTpoCredNiveles.frx":034E
            text            =   "0"
            enabled         =   -1
         End
         Begin VB.Label Label2 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.Frame fraAplicable 
         Caption         =   " Aplicable a Tipo Créditos"
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
         Height          =   3735
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
         Begin VB.CheckBox chkTodosTpoCred 
            Caption         =   "Todos"
            Height          =   255
            Left            =   285
            TabIndex        =   3
            Top             =   360
            Width           =   1215
         End
         Begin VB.ListBox lstTpoCred 
            Height          =   2760
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   2
            Top             =   675
            Width           =   4575
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   765
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmCredBPPTpoCredNiveles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fnTipo As Integer
'Public Sub Inicio(ByVal pnTipo As Integer)
'fnTipo = 0
'Select Case pnTipo
'    Case 1:
'            Me.Caption = "BPP - Sub Producto X Niveles"
'            SSTCategorias.Caption = "Parametros Sub Productos"
'            fraAplicable.Caption = "Aplicable a Sub Productos"
'    Case 2:
'            Me.Caption = "BPP - Tipo de Créditos X Niveles"
'            SSTCategorias.Caption = "Parametros Tipo Créditos"
'            fraAplicable.Caption = "Aplicable a Tipo Créditos"
'End Select
'
'fnTipo = pnTipo
'
'CargoControles
'Me.Show 1
'End Sub
'Private Sub CargoControles()
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'
''CARGA NIVELES
'Set oConst = New COMDConstantes.DCOMConstantes
'Set rsConst = oConst.RecuperaConstantes(7064)
'
''CARGA COMBO DE NIVELES
'Call Llenar_Combo_con_Recordset(rsConst, cmbNiveles)
'
'LlenaListas lstTpoCred
'End Sub
'Private Sub LlenaListas(ByRef lista As ListBox, Optional pnNivel As Integer = 0)
'    Dim oCredito As COMDCredito.DCOMCredito
'    Set oCredito = New COMDCredito.DCOMCredito
'    Dim oLista As COMDCredito.DCOMBPPR
'
'    Dim rs As ADODB.Recordset
'    Dim rsLista As ADODB.Recordset
'    Dim i As Integer, J As Integer
'
'    Set oLista = New COMDCredito.DCOMBPPR
'
'    If pnNivel = 0 Then
'        Set rs = Nothing
'    Else
'        Set rs = oLista.ObtenerSubProTpoCredXNivel(pnNivel, fnTipo)
'    End If
'
'    Call CheckLista(False, lstTpoCred)
'
'    Set oLista = Nothing
'
'    If fnTipo = 1 Then
'        Set rsLista = oCredito.RecuperaSubProductosCrediticios
'    ElseIf fnTipo = 2 Then
'        Set rsLista = oCredito.RecuperaSubTipoCrediticios
'    End If
'
'    For i = 0 To rsLista.RecordCount - 1
'        lista.AddItem rsLista!nConsValor & " " & Trim(rsLista!cConsDescripcion)
'        If pnNivel <> 0 Then
'            If Not (rs.EOF And rs.BOF) Then
'                rs.MoveFirst
'                For J = 0 To rs.RecordCount - 1
'                    If CStr(Trim(rsLista!nConsValor)) = Trim(rs!cTpoSubProd) Then
'                        lista.Selected(i) = True
'                    End If
'                    rs.MoveNext
'                Next J
'            End If
'        End If
'        rsLista.MoveNext
'    Next i
'End Sub
'
'Private Sub chkTodosTpoCred_Click()
' Call CheckLista(IIf(chkTodosTpoCred.value = 1, True, False), lstTpoCred)
'End Sub
'
'
'
'Private Sub cmbNiveles_Click()
'Dim oBPP As COMDCredito.DCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMDCredito.DCOMBPPR
'
'
'If Trim(cmbNiveles.Text) <> "" Then
'    lstTpoCred.Clear
'    Call LlenaListas(lstTpoCred, Int(Trim(Right(cmbNiveles.Text, 5))))
'
'    Set rsBPP = oBPP.ObtenerCarteraTpoCred(CInt(Trim(Right(cmbNiveles.Text, 5))), fnTipo)
'
'    If Not (rsBPP.EOF And rsBPP.BOF) Then
'        txtMinCartera.Text = Format(CDbl(rsBPP!nMin), "###," & String(15, "#") & "#0." & String(2, "0"))
'        txtMaxCartera.Text = Format(CDbl(rsBPP!nMax), "###," & String(15, "#") & "#0." & String(2, "0"))
'    Else
'        txtMinCartera.Text = 0
'        txtMaxCartera.Text = 0
'    End If
'End If
'End Sub
'
'Private Sub cmdCancelar_Click()
'LimpiaDatos
'End Sub
'
'Private Sub cmdGuardar_Click()
'If ValidaDatos Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim i As Integer
'        Dim oBPP As COMDCredito.DCOMBPPR
'
'        Set oBPP = New COMDCredito.DCOMBPPR
'
'        Call oBPP.OpeCarteraTpoCred(2, Trim(Right(cmbNiveles.Text, 5)), , , fnTipo)
'        Call oBPP.OpeSubProductosCredNiveles(3, , Trim(Right(cmbNiveles.Text, 5)), fnTipo)
'
'        Call oBPP.OpeCarteraTpoCred(1, Trim(Right(cmbNiveles.Text, 5)), CDbl(txtMinCartera.Text), CDbl(txtMaxCartera.Text), fnTipo)
'        For i = 0 To lstTpoCred.ListCount - 1
'            If lstTpoCred.Selected(i) = True Then
'                Call oBPP.OpeSubProductosCredNiveles(1, Trim(Left(lstTpoCred.List(i), 3)), Trim(Right(cmbNiveles.Text, 5)), fnTipo)
'            End If
'        Next i
'
'        MsgBox "Se grabaron correctamente los Datos", vbInformation, "Aviso"
'
'        LimpiaDatos
'    End If
'End If
'End Sub
'Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
'Dim i As Integer
'    For i = 0 To lstLista.ListCount - 1
'        lstLista.Selected(i) = bCheck
'    Next i
'End Sub
'
'Private Function ValidaDatos() As Boolean
'Dim i As Integer
'ValidaDatos = False
'
'    If Trim(cmbNiveles.Text) = "" Then
'        MsgBox "Seleccione un nivel", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    For i = 0 To lstTpoCred.ListCount - 1
'        If lstTpoCred.Selected(i) = True Then
'            ValidaDatos = True
'        End If
'    Next i
'
'    If Not ValidaDatos Then
'        MsgBox "Seleccione Por lo menos 1 Tipo de Crédito", vbInformation, "Aviso"
'        Exit Function
'    End If
'
'
'ValidaDatos = True
'End Function
'Private Sub LimpiaDatos()
'    cmbNiveles.ListIndex = -1
'    chkTodosTpoCred.value = 0
'    txtMinCartera.Text = 0
'    txtMaxCartera.Text = 0
'    Call CheckLista(False, lstTpoCred)
'End Sub
