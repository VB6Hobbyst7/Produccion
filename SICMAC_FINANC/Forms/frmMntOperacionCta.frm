VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntOperacionCta 
   Caption         =   "Operaciones:Asignación de Cuentas Contables"
   ClientHeight    =   5535
   ClientLeft      =   2040
   ClientTop       =   2025
   ClientWidth     =   8400
   Icon            =   "frmMntOperacionCta.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOpeCod 
      BackColor       =   &H00F0FFFF&
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
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtOpeDesc 
      BackColor       =   &H00F0FFFF&
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
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Top             =   240
      Width           =   5355
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   5340
      TabIndex        =   7
      Top             =   5070
      Width           =   1290
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   6690
      TabIndex        =   8
      Top             =   5070
      Width           =   1230
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cuenta Contable "
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
      Height          =   2160
      Left            =   120
      TabIndex        =   9
      Top             =   2790
      Width           =   8160
      Begin Sicmact.TxtBuscar txtCta 
         Height          =   330
         Left            =   330
         TabIndex        =   0
         Top             =   270
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cta para Exportar"
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
         Height          =   1080
         Left            =   5760
         TabIndex        =   20
         Top             =   810
         Width           =   2130
         Begin VB.TextBox txtCtaContN 
            Height          =   375
            Left            =   225
            MaxLength       =   24
            TabIndex        =   6
            Top             =   405
            Width           =   1680
         End
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2610
         TabIndex        =   18
         Top             =   270
         Width           =   5295
      End
      Begin VB.Frame fraClase 
         Caption         =   "Clase"
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
         Height          =   1080
         Left            =   4260
         TabIndex        =   17
         Top             =   840
         Width           =   1365
         Begin VB.OptionButton OpDebe 
            Caption         =   "&Debe"
            Height          =   240
            Left            =   180
            TabIndex        =   4
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton OpHaber 
            Caption         =   "&Haber"
            Height          =   315
            Left            =   180
            TabIndex        =   5
            Top             =   600
            Width           =   795
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Caráter"
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
         Height          =   1080
         Left            =   2280
         TabIndex        =   16
         Top             =   840
         Width           =   1605
         Begin VB.OptionButton Op1 
            Caption         =   "O&bligatoria"
            Height          =   240
            Left            =   180
            TabIndex        =   2
            Top             =   360
            Width           =   1155
         End
         Begin VB.OptionButton Op2 
            Caption         =   "O&pcional"
            Height          =   240
            Left            =   180
            TabIndex        =   3
            Top             =   660
            Width           =   1140
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ubicación "
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
         Height          =   1080
         Left            =   330
         TabIndex        =   13
         Top             =   840
         Width           =   1575
         Begin VB.TextBox txtOpeCtaOrden 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   480
            MaxLength       =   1
            TabIndex        =   19
            Top             =   420
            Width           =   525
         End
      End
   End
   Begin VB.Frame fraOpe 
      Caption         =   "Operación "
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
      Height          =   675
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8175
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         Height          =   195
         Left            =   1680
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid grdCtas 
      CausesValidation=   0   'False
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   3334
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cCtaContDesc"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   2280.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5309.858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMntOperacionCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim varCta As String
Dim varOpDH As String, varOpOp As String
Dim varOpeOrden As String

Dim EnGrid As Boolean

Dim clsCta As DCtaCont
Dim rsCta  As ADODB.Recordset
Dim sOpeCod As String, sOpeDesc As String
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(OpeCod As String, OpeDesc As String)
sOpeDesc = OpeDesc
sOpeCod = OpeCod
Me.Show 1
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(Trim(txtDesc.Text)) = 0 Then
   N = MsgBox("La Cuenta Contable no está registrada en el Plan de Cuentas...", vbCritical, "Error")
   txtCta.Text = ""
   txtCta.SetFocus
   Exit Function
End If

If txtOpeCtaOrden = "" Then
   txtOpeCtaOrden = "0"
End If

varCta = txtCta.Text
If OpDebe.value Then
   varOpDH = "D"
Else
   varOpDH = "H"
End If

If Op1.value Then
   varOpOp = "1"
Else
   varOpOp = "2"
End If
varOpeOrden = txtOpeCtaOrden
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
Dim N As Integer, strGRABA As String, sQRY As String
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro de Asignar Cuenta Contable a Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
Dim clsOpe As New DOperacion
gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
clsOpe.InsertaOpeCta sOpeCod, varOpeOrden, Trim(varCta), varOpDH, varOpOp, txtCtaContN, gsMovNro
Set clsOpe = Nothing
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Se Asigno la |Cta Cod : " & txtCta.Text & "|" & txtDesc & " a la Operacion : " & txtOpeDesc.Text
            Set objPista = Nothing
            '*******
glAceptar = True
Unload Me
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCerrar_Click()
glAceptar = False
Unload Me
End Sub

Private Sub Form_Load()
EnGrid = False
Set clsCta = New DCtaCont
Set rsCta = clsCta.CargaCtaCont(" cCtaContCod LIKE '__[^0]%' or len(cCtaContCod) = 2 ", "CtaCont")
Set grdCtas.DataSource = rsCta

CentraForm Me
txtCta.rs = rsCta
txtCta.lbUltimaInstancia = False
txtCta.EditFlex = False
txtCta.TipoBusqueda = BuscaDatoEnGrid

OpDebe.value = True
Op1.value = True
txtOpeCtaOrden = "0"

Me.Caption = "Operaciones: Mantenimiento: Asignación de Cuentas"

txtOpeCod.Text = Trim(sOpeCod)
txtOpeDesc.Text = UCase(sOpeDesc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsCta
Set clsCta = Nothing
End Sub

Private Sub grdCtas_GotFocus()
EnGrid = True
End Sub

Private Sub grdCtas_LostFocus()
EnGrid = False
End Sub

Private Sub grdCtas_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If EnGrid Then
   If Not rsCta.EOF Then
      txtCta.Text = rsCta!cCtaContCod
      txtDesc.Text = rsCta!cCtaContDesc
   End If
End If
End Sub

Private Sub Op1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If OpDebe.value Then
      OpDebe.SetFocus
   Else
      OpHaber.SetFocus
   End If
End If
End Sub

Private Sub Op2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If OpDebe.value Then
      OpDebe.SetFocus
   Else
      OpHaber.SetFocus
   End If
End If
End Sub

Private Sub OpDebe_KeyPress(KeyAscii As Integer)
varOpDH = "D"
If txtCtaContN.Visible And txtCtaContN.Enabled Then
   txtCtaContN.SetFocus
Else
   cmdAceptar.SetFocus
End If
End Sub

Private Sub OpHaber_KeyPress(KeyAscii As Integer)
varOpDH = "H"
If txtCtaContN.Visible And txtCtaContN.Enabled Then
   txtCtaContN.SetFocus
Else
   cmdAceptar.SetFocus
End If
End Sub

Private Sub OpUbi2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Op1.value Then
      Op1.SetFocus
   Else
      Op2.SetFocus
   End If
End If
End Sub

Private Sub txtCta_Change()
If Len(Trim(txtCta.Text)) <> 0 Then
   BuscaDato "cCtaContCod like '" & txtCta.Text & "'", rsCta, 1, False
End If
End Sub

Private Sub txtCta_EmiteDatos()
txtDesc = txtCta.psDescripcion
If txtDesc <> "" And txtOpeCtaOrden.Visible Then
   txtOpeCtaOrden.SetFocus
End If
End Sub

Private Sub txtCta_GotFocus()
txtCta.BackColor = "&H00F0FFFF"
End Sub

Private Sub txtCta_LostFocus()
txtCta.BackColor = "&H80000005"
End Sub

Private Sub txtCtaContN_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtOpeCtaOrden_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Op1.value Then
      Op1.SetFocus
   Else
      Op2.SetFocus
   End If
End If
End Sub
