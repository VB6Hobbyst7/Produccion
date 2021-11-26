VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPRequisitosPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Requisitos Pago Bono"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "frmCredBPPRequisitosPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstDetalleBono 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle Condición Bono:"
      TabPicture(0)   =   "frmCredBPPRequisitosPago.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBonoPlus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraBonoPlus 
         Caption         =   "Datos del Nivel"
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
         Height          =   4455
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   9375
         Begin VB.CommandButton cmdCerrar 
            Caption         =   "Cerrar"
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
            Left            =   4080
            TabIndex        =   6
            Top             =   3960
            Width           =   1290
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nivel:"
            Height          =   195
            Left            =   360
            TabIndex        =   5
            Top             =   405
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Categoría:"
            Height          =   195
            Left            =   1920
            TabIndex        =   4
            Top             =   405
            Width           =   750
         End
         Begin VB.Label lblNivel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblCategoria 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2760
            TabIndex        =   2
            Top             =   360
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmCredBPPRequisitosPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fAnalista As AnalistaBPP
'Private i As Integer
'Public Function Inicio(ByVal pnTipo As Integer, ByRef pAnalista As AnalistaBPP)
'Select Case pnTipo
'    Case 1: sstDetalleBono.TabCaption(0) = sstDetalleBono.TabCaption(0) & " Meta"
'    Case 2: sstDetalleBono.TabCaption(0) = sstDetalleBono.TabCaption(0) & " Rendimiento"
'End Select
'fAnalista = pAnalista
'lblCategoria.Caption = fAnalista.Categoria
'lblNivel.Caption = fAnalista.Nivel
'CargaConstantes
'Me.Show 1
'End Function
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub CargaConstantes()
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'
'Set oConst = New COMDConstantes.DCOMConstantes
'
'LimpiaFlex feRequisitoPago
'Set rsConst = oConst.RecuperaConstantes(7071)
'If Not (rsConst.EOF And rsConst.BOF) Then
'     For i = 0 To rsConst.RecordCount - 1
'        feRequisitoPago.AdicionaFila
'        feRequisitoPago.TextMatrix(i + 1, 1) = Trim(rsConst!cConsDescripcion)
'
'        Select Case i
'            Case 0, 1, 2: feRequisitoPago.TextMatrix(i + 1, 2) = "Meta"
'            Case 3: feRequisitoPago.TextMatrix(i + 1, 2) = "Plus"
'            Case 4, 5: feRequisitoPago.TextMatrix(i + 1, 2) = "Rendimiento"
'            Case 6: feRequisitoPago.TextMatrix(i + 1, 2) = "Global"
'        End Select
'        rsConst.MoveNext
'    Next i
'End If
'Set rsConst = Nothing
'Set oConst = Nothing
'End Sub
'
'Private Sub feRequisitoPago_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
'Cancel = ValidaFlex(feRequisitoPago, pnCol)
'End Sub
