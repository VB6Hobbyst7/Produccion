VERSION 5.00
Begin VB.Form frmLogSubastaCuadreCaja 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmLogSubastaCuadreCaja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Sicmact.Usuario usu 
      Left            =   675
      Top             =   3540
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5340
      TabIndex        =   4
      Top             =   3555
      Width           =   1095
   End
   Begin VB.CommandButton cmdSobFal 
      Caption         =   "Declarar Sobante / Faltante  >>>"
      Height          =   360
      Left            =   2460
      TabIndex        =   3
      Top             =   3555
      Width           =   2775
   End
   Begin Sicmact.FlexEdit flex 
      Height          =   2190
      Left            =   30
      TabIndex        =   0
      Top             =   435
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   3863
      Cols0           =   3
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Operacion-Monto"
      EncabezadosAnchos=   "300-4000-1500"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X"
      ListaControles  =   "0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-R"
      FormatosEdit    =   "0-0-2"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblSobFal 
      Caption         =   "Sobrante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   3135
      Width           =   3090
   End
   Begin VB.Label lblDiffG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   4440
      TabIndex        =   10
      Top             =   3120
      Width           =   1965
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "DIFF   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   3180
      Width           =   660
   End
   Begin VB.Label lblBilletajeG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   1185
      TabIndex        =   8
      Top             =   2670
      Width           =   1965
   End
   Begin VB.Label lblBilletaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BILLE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   225
      TabIndex        =   7
      Top             =   2730
      Width           =   645
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1035
      TabIndex        =   2
      Top             =   105
      Width           =   5385
   End
   Begin VB.Label lblUsus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   945
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Top             =   2670
      Width           =   1965
   End
   Begin VB.Label lblTotala 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   2730
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   3300
      Top             =   2655
      Width           =   3120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   45
      Top             =   2655
      Width           =   3120
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   345
      Left            =   3300
      Top             =   3105
      Width           =   3120
   End
End
Attribute VB_Name = "frmLogSubastaCuadreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsOpeCod As String
Dim lsUser As String
Dim lsCaption As String
Dim lnMonto As Currency
Dim lnMontoBill As Currency
Dim lsPersCod As String

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub Ini(psUser As String, psOpeCod As String, psCaption As String)
    lsOpeCod = psOpeCod
    lsCaption = psCaption
    lsUser = psUser
    Me.Show 1
End Sub

Private Sub cmdSobFal_Click()
    Dim vCodConta1 As String
    Dim vCodConta2 As String
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim oMov As DMov
    Set oMov = New DMov
    
    If MsgBox("Desea Registrar el Sobrante/falnate de su Caja ?", vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oMov.BeginTrans
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, lsOpeCod, Caption, gMovEstContabMovContable, gMovFlagVigente
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
        oMov.InsertaMovGasto lnMovNro, lsPersCod, ""
        
        If lnMonto > lnMontoBill Then
            'Faltante
            vCodConta1 = "111102" & Right(gsCodAge, 2)
            vCodConta2 = IIf(lnMonto - lnMontoBill > 1, "191802020702", "63110101")
            'Cambios en Moneda y Agencia
            
            oMov.InsertaMovCta lnMovNro, 1, vCodConta1, Abs(Val(Me.lblDiffG.Caption))
            oMov.InsertaMovCta lnMovNro, 2, vCodConta2, Abs(Val(Me.lblDiffG.Caption)) * -1
        ElseIf lnMonto < lnMontoBill Then
            'Sobrante
            vCodConta1 = "111102" & Right(gsCodAge, 2)
            vCodConta2 = IIf(lnMonto - lnMontoBill > 5, "291201020702", "521229029907")
            
            oMov.InsertaMovCta lnMovNro, 1, vCodConta1, Abs(Val(Me.lblDiffG.Caption)) * -1
            oMov.InsertaMovCta lnMovNro, 2, vCodConta2, Abs(Val(Me.lblDiffG.Caption))
        End If
    oMov.CommitTrans
    
    Form_Load
End Sub

Private Sub Form_Load()
    Dim oSubasta As DSubasta
    Dim lnI As Integer
    Set oSubasta = New DSubasta
    
    Me.lblUsus.Caption = lsUser
    usu.Inicio lsUser
    Me.lblNombre.Caption = usu.UserNom
    
    flex.rsFlex = oSubasta.GetMovSubasta(lsUser, lsOpeCod, "111102")
    lnMontoBill = GetTotUserBilletaje(gnSubRegBilletaje, lsUser, gdFecSis)
    Me.lblBilletajeG.Caption = Format(lnMontoBill, "#,##0.00")
    
    If oSubasta.VerfDevSubasta(gsCodUser, lsOpeCod, gdFecSis) Then
        Me.cmdSobFal.Enabled = False
    Else
        Me.cmdSobFal.Enabled = True
    End If
    
    lnMonto = 0
    For lnI = 1 To Me.flex.Rows - 1
        If Me.cmdSobFal.Enabled Then
            lnMonto = lnMonto + CCur(IIf(IsNumeric(Me.flex.TextMatrix(lnI, 2)), Me.flex.TextMatrix(lnI, 2), "0"))
        Else
            If IsNumeric(Me.flex.TextMatrix(lnI, 2)) Then
                If Val(Me.flex.TextMatrix(lnI, 2)) > 0 And (Me.flex.Rows - 1) = lnI Then
                    lnMonto = lnMonto - Abs(CCur(IIf(IsNumeric(Me.flex.TextMatrix(lnI, 2)), Me.flex.TextMatrix(lnI, 2), "0")))
                Else
                    lnMonto = lnMonto + Abs(CCur(IIf(IsNumeric(Me.flex.TextMatrix(lnI, 2)), Me.flex.TextMatrix(lnI, 2), "0")))
                End If
            End If
        End If
    Next lnI
        
    Me.lblTotal.Caption = Format(lnMonto, "#,##0.00")
        
    If lnMonto > lnMontoBill Then
        Me.lblSobFal.Caption = "Faltante"
    ElseIf lnMonto < lnMontoBill Then
        Me.lblSobFal.Caption = "Sobrante"
    Else
        Me.lblSobFal.Caption = ""
    End If
    
    Me.lblDiffG.Caption = Format(lnMontoBill - lnMonto, "#,##0.00")
    
End Sub
