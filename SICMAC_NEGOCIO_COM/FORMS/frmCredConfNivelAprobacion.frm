VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCredConfNivelAprobacion 
   Caption         =   "Configuracion de Niveles de Aprobacion"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   Icon            =   "frmCredConfNivelAprobacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Nivel de Aprobación"
      TabPicture(0)   =   "frmCredConfNivelAprobacion.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tipo de Crédito"
      TabPicture(1)   =   "frmCredConfNivelAprobacion.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Chkprod"
      Tab(1).Control(2)=   "LstProd"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Agencias"
      TabPicture(2)   =   "frmCredConfNivelAprobacion.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "LstAgencias"
      Tab(2).Control(2)=   "ChkAgeTotal"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   22
         Top             =   5040
         Width           =   8295
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&ACEPTAR"
            Height          =   350
            Left            =   3000
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&CANCELAR"
            Height          =   350
            Left            =   4560
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox ChkAgeTotal 
         Caption         =   "Todas"
         Height          =   240
         Left            =   -69600
         TabIndex        =   20
         Top             =   600
         Width           =   795
      End
      Begin VB.ListBox LstAgencias 
         Height          =   3210
         Left            =   -72240
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   960
         Width           =   3435
      End
      Begin VB.ListBox LstProd 
         Height          =   3210
         Left            =   -72240
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   960
         Width           =   3360
      End
      Begin VB.CheckBox Chkprod 
         Caption         =   "Todos"
         Height          =   240
         Left            =   -69600
         TabIndex        =   16
         Top             =   600
         Width           =   795
      End
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   8295
         Begin VB.TextBox txtMontoMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5160
            TabIndex        =   25
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox txtMontoMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Top             =   3120
            Width           =   1935
         End
         Begin VB.ComboBox cmbNivelAprob 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdEliminarArea 
            Caption         =   "Eliminar"
            Height          =   350
            Left            =   7200
            TabIndex        =   6
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregarArea 
            Caption         =   "Agregar"
            Height          =   350
            Left            =   7200
            TabIndex        =   5
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cmbRiesgoAprob 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3600
            Width           =   615
         End
         Begin SICMACT.FlexEdit FlexCargo 
            Height          =   2400
            Left            =   1440
            TabIndex        =   4
            Top             =   600
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4233
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Cod Cargo-Cargo"
            EncabezadosAnchos=   "200-900-4500"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-1-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L"
            FormatosEdit    =   "0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            Appearance      =   0
            ColWidth0       =   195
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Hasta S/."
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
            Height          =   195
            Left            =   4200
            TabIndex        =   26
            Top             =   3120
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nivel: "
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
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cargo:"
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
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Riesgo:"
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
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Mayor de S/."
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
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   3120
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8295
         Begin VB.OptionButton optNivelesPor 
            Caption         =   "CLIENTES PREFERENCIALES"
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   2
            Top             =   240
            Width           =   2655
         End
         Begin VB.OptionButton optNivelesPor 
            Caption         =   "TIPO DE CREDITO"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nivel Por:"
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
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Agencias:"
         Height          =   225
         Left            =   -72240
         TabIndex        =   21
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Tipos de Créditos :"
         Height          =   225
         Left            =   -72240
         TabIndex        =   18
         Top             =   600
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmCredConfNivelAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objConst As COMDConstantes.DCOMConstantes

Private Sub Form_Load()
    CentraForm Me
    CargarNivelJerarq
    CargarNivelAprob
    CargarRiesgoAprob
    CargaControles
End Sub

Private Sub CargarNivelJerarq()
    Dim rsNivelJerarq As ADODB.Recordset
    
    Set objConst = New COMDConstantes.DCOMConstantes
    Set rsNivelJerarq = objConst.ObtenerCargosArea
    
    FlexCargo.rsTextBuscar = rsNivelJerarq
    
    Set objConst = Nothing
    Set rsNivelJerarq = Nothing
End Sub

Private Sub CargarNivelAprob()
    Dim rsNivelAprob As ADODB.Recordset
    
    Set objConst = New COMDConstantes.DCOMConstantes
    Set rsNivelAprob = objConst.ObtenerNivelAprob
    
    With rsNivelAprob
    Do Until .EOF
        cmbNivelAprob.AddItem .Fields(1) & Space(50) & .Fields(0)
        .MoveNext
    Loop
    End With
    
    cmbNivelAprob.ListIndex = 0
    Set objConst = Nothing
    Set rsNivelAprob = Nothing
End Sub

Private Sub CargarRiesgoAprob()
    Dim rsRiesgoAprob As ADODB.Recordset
    
    Set objConst = New COMDConstantes.DCOMConstantes
    Set rsRiesgoAprob = objConst.ObtenerRiesgoAprob
    
    With rsRiesgoAprob
    Do Until .EOF
        cmbRiesgoAprob.AddItem .Fields(1) & Space(50) & .Fields(0)
        .MoveNext
    Loop
    End With
    
    cmbRiesgoAprob.ListIndex = 0
    Set objConst = Nothing
    Set rsRiesgoAprob = Nothing
End Sub

Private Sub Limpiar()
    Dim i, J As Integer
    optNivelesPor(0).value = True
    cmbNivelAprob.ListIndex = 0
    TxtMontoMin.Text = ""
    TxtMontoMax.Text = ""
    cmbRiesgoAprob.ListIndex = 0
    FlexCargo.Clear
    FlexCargo.Rows = 2
    FlexCargo.FormaCabecera
    For i = 0 To LstProd.ListCount - 1
        LstProd.Selected(i) = False
    Next i
    Chkprod.value = 0
    For J = 0 To LstAgencias.ListCount - 1
        LstAgencias.Selected(J) = False
    Next J
    ChkAgeTotal.value = 0
    SSTab.Tab = 0
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    Dim sCorrelativo As String
    Dim objNivApr As COMNCredito.NCOMNivelAprobacion
    Set objNivApr = New COMNCredito.NCOMNivelAprobacion
    
    If Validar = False Then
        MsgBox ("Verifique que los Datos esten Completos"), vbCritical, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se Va A Grabar los Datos, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        sCorrelativo = Format(ObtenerCorrelativo + 1, "000000")
        objNivApr.InsertarNivApr sCorrelativo, IIf(optNivelesPor(0).value = True, 1, 0), Trim(Right(cmbNivelAprob.Text, 3)), TxtMontoMin.Text, TxtMontoMax.Text, Trim(Right(cmbRiesgoAprob.Text, 3)), FlexCargo.GetRsNew
        
        RegistrarTipoCredito_Agencias (sCorrelativo)
        MsgBox ("Los Datos se registraron con exito"), vbInformation, "Aviso"
        sCorrelativo = ""
        Limpiar
        Set objNivApr = Nothing
    End If
End Sub

Private Sub CmdCancelar_Click()
    Limpiar
End Sub

Private Function Validar() As Boolean
    Dim K, L, M, nCuentaCargo, nCuentaTipoProducto, nCuentaAgencias As Integer
    For K = 1 To FlexCargo.Rows - 1
        If Len(Me.FlexCargo.TextMatrix(K, 0)) > 0 Then
            nCuentaCargo = nCuentaCargo + 1
        End If
    Next K
    
    If nCuentaCargo = 0 Then
        Validar = False
        Exit Function
    End If
    
    If TxtMontoMin.Text = "" Then
        Validar = False
        Exit Function
    End If
    
    If TxtMontoMax.Text = "" Then
        Validar = False
        Exit Function
    End If
    
    For L = 0 To LstProd.ListCount - 1
        If LstProd.Selected(L) = True Then
            nCuentaTipoProducto = nCuentaTipoProducto + 1
        End If
    Next L
    
    If nCuentaTipoProducto = 0 Then
        Validar = False
        Exit Function
    End If
    
    For M = 0 To LstAgencias.ListCount - 1
        If LstAgencias.Selected(M) = True Then
            nCuentaAgencias = nCuentaAgencias + 1
        End If
    Next M
    
    If nCuentaAgencias = 0 Then
        Validar = False
        Exit Function
    End If
    
    Validar = True
End Function

Private Function ObtenerCorrelativo() As String
    Dim objNivApr As COMDCredito.DCOMNivelAprobacion
    Dim rsNivApr As ADODB.Recordset
    
    Set objNivApr = New COMDCredito.DCOMNivelAprobacion
    Set rsNivApr = objNivApr.ObtenerCorrelativoNivAprCod
    
    If rsNivApr.RecordCount <> "0" Then
        ObtenerCorrelativo = Trim(rsNivApr.Fields(0))
    End If
    
    Set objNivApr = Nothing
    Set rsNivApr = Nothing
End Function

Private Sub RegistrarTipoCredito_Agencias(ByVal sCorrelativo As String)
    Dim i As Integer
    Dim J As Integer
    Dim objNivAprob As COMDCredito.DCOMNivelAprobacion
    Set objNivAprob = New COMDCredito.DCOMNivelAprobacion
    
    For i = 0 To LstProd.ListCount - 1
        If LstProd.Selected(i) Then
            For J = 0 To LstAgencias.ListCount - 1
                If LstAgencias.Selected(J) Then
                    objNivAprob.InsertarNivAprTipoCredito sCorrelativo, Right(LstProd.List(i), 3), Right(LstAgencias.List(J), 2)
                End If
            Next J
        End If
    Next i
    Set objNivAprob = Nothing
End Sub

Private Sub cmdAgregarArea_Click()
    Me.FlexCargo.AdicionaFila
    FlexCargo.SetFocus
End Sub

Private Sub cmdEliminarArea_Click()
    If MsgBox("Desea Eliminar esta area " & FlexCargo.TextMatrix(FlexCargo.Row, 2) & " ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    FlexCargo.EliminaFila FlexCargo.Row
End Sub

Private Sub CargaControles()
    Dim oCred As COMDCredito.DCOMCredito
    Dim oGasto As COMDCredito.DCOMGasto
    Dim R As ADODB.Recordset
    
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaProductosDeCredito
    Set oCred = Nothing
    LstProd.Clear
    Do While Not R.EOF
        LstProd.AddItem Trim(R!cConsDescripcion) & Space(100) & Trim(R!nConsValor)
        R.MoveNext
    Loop
    R.Close
    
    LstAgencias.Clear
    Set oGasto = New COMDCredito.DCOMGasto
    Set R = oGasto.RecuperaAgencias
    Do While Not R.EOF
        LstAgencias.AddItem Trim(R!cAgeDescripcion) & Space(100) & Trim(R!cAgeCod)
        R.MoveNext
    Loop
    Set oGasto = Nothing
End Sub

Private Sub optNivelesPor_Click(Index As Integer)
    If optNivelesPor(0).value = True Then
        SSTab.TabVisible(1) = True
    Else
        SSTab.TabVisible(1) = False
    End If
End Sub

Private Sub TxtMontoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(TxtMontoMin.Text) > 0 Then
            TxtMontoMin.Text = Format(TxtMontoMin, "#,##0.00")
        End If
        TxtMontoMax.SetFocus
    End If
End Sub

Private Sub TxtMontoMin_LostFocus()
    If Len(TxtMontoMin.Text) > 0 Then
        TxtMontoMin.Text = Format(TxtMontoMin, "#,##0.00")
    End If
    TxtMontoMax.SetFocus
End Sub

Private Sub TxtMontoMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(TxtMontoMax.Text) > 0 Then
            TxtMontoMax.Text = Format(TxtMontoMax, "#,##0.00")
        End If
        cmbRiesgoAprob.SetFocus
    End If
End Sub

Private Sub TxtMontoMax_LostFocus()
    If Len(TxtMontoMax.Text) > 0 Then
            TxtMontoMax.Text = Format(TxtMontoMax, "#,##0.00")
    End If
    cmbRiesgoAprob.SetFocus
End Sub

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
