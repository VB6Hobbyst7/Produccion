VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvTransferenciaBS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TRANSFERENCIA DE BIENES"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvTransferenciaBS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "TRANSFERENCIA DE BIENES"
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   6360
         Width           =   8895
         Begin VB.CommandButton Command4 
            Caption         =   "SALIR"
            Height          =   375
            Left            =   4680
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            Caption         =   "ACEPTAR"
            Height          =   375
            Left            =   3480
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "TIPO DE TRANSFERENCIA"
         Height          =   1095
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   8895
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6480
            TabIndex        =   23
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16908289
            CurrentDate     =   39881
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   2640
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "FECHA:"
            Height          =   255
            Left            =   5400
            TabIndex        =   22
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "TIPO DE TRANSFERENCIA:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   9240
         TabIndex        =   4
         Top             =   1560
         Width           =   975
         Begin VB.CommandButton Command2 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   1440
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "DESTINO:"
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   5400
         Width           =   8895
         Begin VB.TextBox txtLugarEntrega1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.TextBox txtAgeDesc1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            TabIndex        =   11
            Top             =   360
            Width           =   5640
         End
         Begin Sicmact.TxtBuscar txtArea1 
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
         Begin Sicmact.TxtBuscar txtCodLugarEntrega1 
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "ORIGEN:"
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   4440
         Width           =   8895
         Begin VB.TextBox txtLugarEntrega 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.TextBox txtAgeDesc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   360
            Width           =   5640
         End
         Begin Sicmact.TxtBuscar txtArea 
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
         End
         Begin Sicmact.TxtBuscar txtCodLugarEntrega 
            Height          =   315
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   556
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "BIENES A TRANSFERIR"
         Height          =   2775
         Left            =   240
         TabIndex        =   1
         Top             =   1560
         Width           =   8895
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
            Height          =   2325
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   4101
            _Version        =   393216
            Rows            =   51
            Cols            =   6
            ForeColorSel    =   -2147483643
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483637
            AllowBigSelection=   0   'False
            TextStyleFixed  =   3
            FocusRect       =   0
            HighLight       =   2
            GridLinesFixed  =   1
            Appearance      =   0
            RowSizingMode   =   1
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
      End
   End
End
Attribute VB_Name = "frmInvTransferenciaBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Index As Integer

Private Sub cmbTipo_Click()
'    If cmbTipo.ListIndex = 1 Then
'        CargarTransferenciaInterna
'
'        txtArea.Visible = True
'        txtAgeDesc.Visible = True
'        txtArea1.Visible = True
'        txtAgeDesc1.Visible = True
'
'        txtCodLugarEntrega.Visible = False
'        txtLugarEntrega.Visible = False
'        txtCodLugarEntrega1.Visible = False
'        txtLugarEntrega1.Visible = False
'
'    ElseIf cmbTipo.ListIndex = 2 Then
'        CargarTransferenciaEntreAgencias
'
'        txtArea.Visible = False
'        txtAgeDesc.Visible = False
'        txtArea1.Visible = False
'        txtAgeDesc1.Visible = False
'
'        txtCodLugarEntrega.Visible = True
'        txtLugarEntrega.Visible = True
'        txtCodLugarEntrega1.Visible = True
'        txtLugarEntrega1.Visible = True
'    End If
End Sub

Private Sub Command1_Click()
    frmInvBuscarBS.Show
End Sub

Public Function ValidarDatos(ByVal sMovNro As String, ByVal sCodBS As String) As Boolean
    Dim k  As Integer
    k = 1
    Do While k < fgDetalle.Rows
       If fgDetalle.TextMatrix(k, 1) = sCodBS And fgDetalle.TextMatrix(k, 3) = sMovNro Then
          ValidarDatos = True
          Exit Function
        Else
          ValidarDatos = False
        End If
        k = k + 1
    Loop
End Function

Public Sub ActualizaFG(nItem As Integer, sCodInv As String, sDesc As String, ByVal sMovNro As String, ByVal cValor As Currency, ByVal cValorDepre As Currency)
    fgDetalle.TextMatrix(nItem, 1) = sCodInv
    fgDetalle.TextMatrix(nItem, 2) = sDesc
    fgDetalle.TextMatrix(nItem, 3) = sMovNro
    fgDetalle.TextMatrix(nItem, 4) = Format(cValor, "#,##0.00")
    fgDetalle.TextMatrix(nItem, 5) = Format(cValorDepre, "#,##0.00")
    Index = Index + 1
End Sub

Private Sub CargarCabecera()
    fgDetalle.Cols = 6
    fgDetalle.TextMatrix(0, 0) = " "
    fgDetalle.TextMatrix(1, 0) = " "
    fgDetalle.TextMatrix(0, 1) = "COD. INVENTARIO"
    fgDetalle.TextMatrix(0, 2) = "DESCRIPCION"
    fgDetalle.TextMatrix(0, 3) = "sMovNro"
    fgDetalle.TextMatrix(0, 4) = "VALOR"
    
    fgDetalle.TextMatrix(0, 5) = "Depre"
    
    fgDetalle.ColWidth(0) = 400
    fgDetalle.ColWidth(1) = 1500
    fgDetalle.ColWidth(2) = 6100
    fgDetalle.ColWidth(3) = 0
    fgDetalle.ColWidth(4) = 1500
    fgDetalle.ColWidth(5) = 1500
    
    fgDetalle.RowHeight(0) = 300
    fgDetalle.RowHeight(1) = 300
End Sub

Private Sub Command2_Click()
    If fgDetalle.TextMatrix(fgDetalle.RowSel, 1) <> "" Then
        EliminaFila fgDetalle.Row
        Index = Index - 1
    Else
        MsgBox "Debe Escoger una Fila!"
    End If
End Sub

Private Sub EliminaFila(nItem As Integer)
    Dim k  As Integer, m As Integer
    k = 1
    Do While k < fgDetalle.Rows
       If Len(fgDetalle.TextMatrix(k, 1)) > 0 Then
          If Val(fgDetalle.TextMatrix(k, 0)) = nItem Then
             EliminaRow fgDetalle, k
             k = k - 1
          Else
             k = k + 1
          End If
       Else
          k = k + 1
       End If
    Loop
End Sub

Private Sub Command3_Click()
    If cmbTipo.Text <> "SELECCIONE" And (Not Validar = False) And txtArea.Text <> "" And txtArea1.Text <> "" Then
        Dim liTransferenciaId, k As Integer
        Dim oInventario As NInvTransferencia
        Set oInventario = New NInvTransferencia
        liTransferenciaId = oInventario.InsertarTransferencia(cmbTipo.ListIndex, DTPicker1.value, txtAgeDesc.Text, txtAgeDesc1.Text, IIf(Len(txtArea.Text) > 3, txtArea.Text, txtArea.Text & "01"), IIf(Len(txtArea1.Text) > 3, txtArea1.Text, txtArea1.Text & "01"))
        k = 1
        Do While k < fgDetalle.Rows
            If fgDetalle.TextMatrix(k, 1) <> "" Then
                oInventario.InsertarInventarioTransferencia liTransferenciaId, fgDetalle.TextMatrix(k, 1)
                'oInventario.ActualizarBS fgDetalle.TextMatrix(k, 3), fgDetalle.TextMatrix(k, 1), Left(txtArea1.Text, 3), IIf(Len(txtArea1.Text) > 3, Right(txtArea1.Text, 2), "01"), GenerarNvoCod(fgDetalle.TextMatrix(k, 1), IIf(Len(txtArea1.Text) > 3, Right(txtArea1.Text, 2), "01"))
                Dim i As Integer
                For i = 0 To 1
                    oInventario.InsertarAsientoTransferencia IIf(i = 0, fgDetalle.TextMatrix(k, 1), GenerarNvoCod(fgDetalle.TextMatrix(k, 1), IIf(Len(txtArea1.Text) > 3, Right(txtArea1.Text, 2), "01"))), IIf(i = 0, Mid(fgDetalle.TextMatrix(k, 1), 10, 5), Mid(GenerarNvoCod(fgDetalle.TextMatrix(k, 1), IIf(Len(txtArea1.Text) > 3, Right(txtArea1.Text, 2), "01")), 10, 5)), fgDetalle.TextMatrix(k, 4), DTPicker1.value, IIf(i = 0, "O", "D"), fgDetalle.TextMatrix(k, 5)
                Next i
                
            End If
            k = k + 1
        Loop
        MsgBox "Los Datos se Registraron Correctamente", vbInformation
        Limpiar
    Else
        MsgBox "Debe Seleccionar Los Datos Correctos!", vbCritical
    End If
End Sub

Private Function GenerarNvoCod(ByVal sCodInventario As String, ByVal sCodAge As String) As String
    Dim sCodBSNvo As String
    sCodBSNvo = Replace(sCodInventario, Mid(sCodInventario, 7, 3), Format(sCodAge, "000"))
    GenerarNvoCod = sCodBSNvo
End Function

Private Function Validar() As Boolean
    Dim i As Integer
    i = 1
    Do While i < fgDetalle.Rows
        If fgDetalle.TextMatrix(i, 2) <> "" Then
            Validar = True
            Exit Function
        Else
            Validar = False
        End If
        i = i + 1
    Loop
End Function

Private Sub Limpiar()
    Dim n, nItem As Integer
    txtArea.Text = ""
    txtCodLugarEntrega.Text = ""
    txtArea1.Text = ""
    txtCodLugarEntrega1.Text = ""
    txtLugarEntrega.Text = ""
    txtAgeDesc.Text = ""
    txtLugarEntrega1.Text = ""
    txtAgeDesc1.Text = ""
    cmbTipo.Text = "SELECCIONE"
    DTPicker1.value = Date
    Index = 1
    For n = 1 To fgDetalle.Rows - 1
        For nItem = 1 To fgDetalle.Cols - 1
            fgDetalle.TextMatrix(n, nItem) = ""
        Next
    Next
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub TxtArea_EmiteDatos()
    Me.txtAgeDesc.Text = txtArea.psDescripcion
End Sub

Private Sub TxtCodLugarEntrega_EmiteDatos()
    Me.txtLugarEntrega.Text = txtCodLugarEntrega.psDescripcion
End Sub

Private Sub txtArea1_EmiteDatos()
    Me.txtAgeDesc1.Text = txtArea1.psDescripcion
End Sub

Private Sub txtCodLugarEntrega1_EmiteDatos()
    Me.txtLugarEntrega1.Text = txtCodLugarEntrega1.psDescripcion
End Sub

Private Sub Form_Load()
    Index = 1
    CargarCabecera
    EnumeraItems fgDetalle
    Call CentraForm(Me)
    CargarTipoTransferencia
    DTPicker1.value = Date
    CargarTransferenciaInterna
End Sub

Private Sub CargarTransferenciaInterna()
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim oConst As DConstantes
    Set oConst = New DConstantes
    
    Me.txtArea.rs = oArea.GetAgenciasAreas
    Me.txtArea1.rs = oArea.GetAgenciasAreas
End Sub

Private Sub CargarTipoTransferencia()
    cmbTipo.AddItem "SELECCIONE", 0
    cmbTipo.AddItem "INTERNO", 1
    cmbTipo.AddItem "ENTRE AGENCIAS", 2
    cmbTipo.Text = "SELECCIONE"
End Sub
