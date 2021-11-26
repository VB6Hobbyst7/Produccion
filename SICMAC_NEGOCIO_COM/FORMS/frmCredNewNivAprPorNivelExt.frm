VERSION 5.00
Begin VB.Form frmCredNewNivAprPorNivelExt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de Niveles de Aprobación"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmCredNewNivAprPorNivelExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Extornar Niveles de Aprobación"
      Top             =   2780
      Width           =   1000
   End
   Begin VB.CommandButton cmdHistorial 
      Caption         =   "&Historial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Historial de los Niveles de Aprobación"
      Top             =   2780
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   80
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   2780
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4995
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   2780
      Width           =   1000
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Datos de la cuenta"
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
      Height          =   2715
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtTasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   1920
         Width           =   1050
      End
      Begin VB.TextBox txtAnalista 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   2280
         Width           =   4290
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   1920
         Width           =   1410
      End
      Begin VB.TextBox txtMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   1920
         Width           =   450
      End
      Begin VB.TextBox txtTipoCredito 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   1560
         Width           =   4290
      End
      Begin VB.TextBox txtTipoProducto 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   1200
         Width           =   4290
      End
      Begin VB.TextBox txtTitular 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "txtPrincipal"
         Top             =   840
         Width           =   4290
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   6
         ToolTipText     =   "Buscar"
         Top             =   320
         Width           =   1000
      End
      Begin SICMACT.ActXCodCta ActXCodCta1 
         Height          =   375
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tasa :"
         Height          =   195
         Left            =   3960
         TabIndex        =   18
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Analista :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Crédito :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Producto :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titular :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmCredNewNivAprPorNivelExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
'** Nombre : frmCredNewNivAprPorNivelExt
'** Descripción : Para extornar los resultados de los niveles de Aprobación creado segun TI-ERS028-2016
'** Creación : EJVG, 20160607 09:00:00 AM
'******************************************************************************************************
Option Explicit

Private Sub ActXCodCta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(ActXCodCta1.NroCuenta) = 18 Then
        cmdBuscar_Click
    End If
End Sub
Private Sub cmdBuscar_Click()
    If Len(ActXCodCta1.NroCuenta) < 18 Then
        MsgBox "Ud. debe especificar el Nro. de Cuenta", vbInformation, "Aviso"
        EnfocaControl ActXCodCta1
        Exit Sub
    End If
    If Not EsAgenciaConNivApr(ActXCodCta1.Age) Then
        MsgBox "Los Niveles de Aprobación no aplican para la Agencia del crédito.", vbInformation, "Aviso"
        EnfocaControl ActXCodCta1
        Exit Sub
    End If
    Call CargarDatos
End Sub
Private Sub cmdCancelar_Click()
    ActXCodCta1.NroCuenta = ""
    ActXCodCta1.CMAC = gsCodCMAC
    ActXCodCta1.Age = Right(gsCodAge, 2)
    LlenarCampos
    fraBusqueda.Enabled = True
    cmdHistorial.Enabled = False
    cmdExtornar.Enabled = False
End Sub
Private Sub cmdExtornar_Click()
    Dim oExterno As COMDCredito.DCOMCredExtorno
    If MsgBox("¿Está seguro de realizar el Extorno de los Niveles de Aprobación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oExterno = New COMDCredito.DCOMCredExtorno
    oExterno.ExtornarNivelAprobacion (ActXCodCta1.NroCuenta)
    
    MsgBox "Se realizó el extorno satisfactoriamente", vbInformation, "Aviso"
    cmdCancelar_Click
End Sub
Private Sub cmdHistorial_Click()
    Call frmCredNewNivAprHist.InicioCredito(ActXCodCta1.NroCuenta)
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    cmdCancelar_Click
End Sub
Private Sub CargarDatos()
    Dim oExtorno As New COMDCredito.DCOMCredExtorno
    Dim rs As ADODB.Recordset
    Dim nResolver As Integer
    
    On Error GoTo ErrHandler
    
    Screen.MousePointer = 11
    Set rs = oExtorno.RecuperaDatosXExtornoNivApr(ActXCodCta1.NroCuenta)
    Set oExtorno = Nothing
    Screen.MousePointer = 0
    
    If Not rs.EOF Then
        LlenarCampos rs!cTitular, rs!cTpoProdDesc, rs!cTpoCredDesc, rs!cMoneda, rs!nMonto, rs!nTasaInteres, rs!cAnalista
        fraBusqueda.Enabled = False
        cmdHistorial.Enabled = True
        cmdExtornar.Enabled = rs!bExtorno
        
        MsgBox "Se ha cargado los datos del crédito", vbInformation, "Aviso"
    Else
        MsgBox "No se ha podido cargar datos del crédito especificado." & Chr(13) & Chr(13) & "Verifique que el crédito esté con estado SUGERIDO.", vbInformation, "Aviso"
    End If
    RSClose rs
    Exit Sub
ErrHandler:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub LlenarCampos(Optional ByVal psTitular As String = "", Optional ByVal psTipoProd As String = "", Optional ByVal psTipoCred As String = "", Optional ByVal psMoneda As String = "", Optional ByVal pnMonto As Double = 0#, Optional ByVal pnTasa As Double = 0#, Optional ByVal psAnalista As String = "")
    txtTitular.Text = psTitular
    txtTipoProducto.Text = psTipoProd
    txtTipoCredito.Text = psTipoCred
    txtMoneda.Text = psMoneda
    txtMonto.Text = Format(pnMonto, "#,##0.00")
    txtTasa.Text = Format(pnTasa, "#,##0.0000")
    txtAnalista.Text = psAnalista
End Sub
