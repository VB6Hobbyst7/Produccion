VERSION 5.00
Begin VB.Form frmCapTasaEspSeg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   Icon            =   "frmCapTasaEspSeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   90
      TabIndex        =   22
      Top             =   4430
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   6660
      TabIndex        =   21
      Top             =   4430
      Width           =   1275
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   5355
      TabIndex        =   20
      Top             =   4430
      Width           =   1275
   End
   Begin VB.Frame fraSolicitud 
      Caption         =   "Datos Solicitud Tasa Especial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2950
      Left            =   90
      TabIndex        =   24
      Top             =   1395
      Width           =   7845
      Begin VB.TextBox txtPlazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5460
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "0"
         Top             =   735
         Width           =   885
      End
      Begin SICMACT.EditMoney txtTasa 
         Height          =   315
         Left            =   5460
         TabIndex        =   17
         Top             =   1155
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboSubProducto 
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
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   735
         Width           =   2265
      End
      Begin VB.ComboBox cboMoneda 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1155
         Width           =   2265
      End
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   315
         Width           =   2265
      End
      Begin VB.TextBox txtComentario 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   225
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "(Escriba aquí su comentario)"
         Top             =   1890
         Width           =   7440
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   315
         Left            =   5460
         TabIndex        =   15
         Top             =   315
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label9 
         Caption         =   "SubProducto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   7
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblDias 
         Caption         =   "Días"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6405
         TabIndex        =   27
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblPlazo 
         AutoSize        =   -1  'True
         Caption         =   "Plazo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4795
         TabIndex        =   13
         Top             =   800
         Width           =   540
      End
      Begin VB.Label lblMon 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7050
         TabIndex        =   26
         Top             =   350
         Width           =   300
      End
      Begin VB.Label Label8 
         Caption         =   "Monto Apertura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3980
         TabIndex        =   12
         Top             =   340
         Width           =   1410
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "% TEA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6970
         TabIndex        =   25
         Top             =   1230
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   1575
         Width           =   1020
      End
      Begin VB.Label Label5 
         Caption         =   "Tasa Solicitada :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3885
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1185
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Datos Persona"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1185
      Left            =   90
      TabIndex        =   23
      Top             =   105
      Width           =   7845
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   330
         Left            =   1260
         TabIndex        =   1
         Top             =   293
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblDocID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4590
         TabIndex        =   3
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1260
         TabIndex        =   5
         Top             =   720
         Width           =   5730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Doc ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3825
         TabIndex        =   2
         Top             =   390
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   0
         Top             =   390
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmCapTasaEspSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

Dim nTasaTarifada As Double 'Agregado por RIRO EL 20130327
Dim nPersoneria As PersPersoneria 'Agregado por RIRO EL 201304161643

' -- Modificado por RIRO el 20130327,
' Se agrego un parametro opcional 'sFiltro3', su valor por defecto permite la carga de constantes activas
Private Sub IniciaCombo(ByRef cboConst As ComboBox, ByVal nCapConst As ConstanteCabecera, _
    Optional sFiltro1 As String = "", Optional sFiltro2 As String = "", Optional sFiltro3 As String = " ")
    
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(nCapConst, sFiltro1, sFiltro2, sFiltro3) ' Modificado por RIRO el 20130327
    Set clsGen = Nothing
    cboConst.Clear ' Agregado por RIRO el 20130327, se considero vaciar el combo antes de iniciarlo
    Do While Not rsConst.EOF
        cboConst.AddItem rsConst("cDescripcion") & Space(100) & rsConst("nConsValor")
        rsConst.MoveNext
    Loop
    cboConst.ListIndex = 0

End Sub
' -- Fin RIRO

Private Sub ClearScreen()
fraPersona.Enabled = True
txtCodigo.Text = ""
lblNombre.Caption = ""
lblDocID.Caption = ""
fraSolicitud.Enabled = False
txtTasa.Text = "0.0000"
txtMonto.Text = "0.00"
txtComentario.Text = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
'Me.txtPlazo.Text = "___0" ' Comentado por RIRO el 20130422
'Por RIRO 20130411
txtPlazo.Text = "30"
cboProducto.ListIndex = 0
cboMoneda.ListIndex = 0
' Fin RIRO
End Sub

Private Sub cboMoneda_Click()
Dim nMon As COMDConstantes.Moneda
nMon = CInt(Trim(Right(cboMoneda.Text, 2)))
If nMon = gMonedaNacional Then
    lblMon.Caption = "S/."
    txtTasa.BackColor = &HFFFFFF
    txtMonto.BackColor = &HFFFFFF
Else
    lblMon.Caption = "US$"
    txtTasa.BackColor = &HC0FFC0
    txtMonto.BackColor = &HC0FFC0
End If

' Agregado por RIRO el 201303271810
cboSubProducto_Click
' Fin RIRO

End Sub

Private Sub cboProducto_Click()

    Dim nProd As COMDConstantes.Producto
    nProd = CInt(Trim(Right(cboProducto.Text, 4)))
    
    If nProd = gCapPlazoFijo Then
        lblPlazo.Visible = True
        lblDias.Visible = True
        txtPlazo.Visible = True
    Else
        lblPlazo.Visible = False
        lblDias.Visible = False
        txtPlazo.Visible = False
        txtPlazo.Text = "30"
    End If

   'Agregado por RIRO el 201303271811
    cboSubProducto.Clear
    If nProd = gCapPlazoFijo Then
        IniciaCombo cboSubProducto, gCaptacSubProdPlazoFijo
    ElseIf nProd = gCapAhorros Then
        IniciaCombo cboSubProducto, gCaptacSubProdAhorros
    ElseIf nProd = gCapCTS Then
        IniciaCombo cboSubProducto, gCaptacSubProdCTS
    End If
    cboSubProducto_Click
    'Fin RIRO

End Sub

' -- Modificado por RIRO el 20130327
Private Sub cboProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboSubProducto.SetFocus
        'cboMoneda.SetFocus
    End If
End Sub
' -- Fin RIRO


' Agregado por RIRO 20130327
' Funcion consiste en actualizar TextBox que muestra las tasas de acuerdo a los parametros
Private Sub cboSubProducto_Click()

    Dim nProducto As COMDConstantes.Producto
    Dim nSubProducto As Integer
    Dim nTipoTasa As COMDConstantes.CaptacTipoTasa
    Dim nmoneda As COMDConstantes.Moneda
    Dim nValor As Double, nPlazo As Long
    Dim rsTasas As ADODB.Recordset

    Dim oTasas As COMNCaptaGenerales.NCOMCaptaDefinicion

    If cboSubProducto.ListCount > 0 Then
    
        Set oTasas = New COMNCaptaGenerales.NCOMCaptaDefinicion
        nTasaTarifada = 0
        nProducto = Trim(Right(cboProducto.Text, 3))
        nSubProducto = Trim(Right(cboSubProducto.Text, 3))
        nTipoTasa = gCapTasaNormal
        nmoneda = Trim(Right(cboMoneda.Text, 3))
        nValor = txtMonto.Text
                
        Dim a As String
        
        If nProducto = gCapPlazoFijo Then
            
            If Not IsNumeric(txtPlazo.Text) Then
            
                MsgBox "Debe ingresar valores numéricos", vbInformation, "Aviso"
                If txtPlazo.Visible Then
                    txtPlazo.SetFocus
                End If
                
                Exit Sub
                
            End If
            
            If nSubProducto = 1 Then ' Plazo Primium
            
                Set rsTasas = oTasas.GetTarifario(gCapPlazoFijo, gMonedaNacional, gCapTasaNormal, gsCodAge, 1)
                
                If Not (rsTasas.EOF And rsTasas.BOF) Then
                
                    txtPlazo.Text = rsTasas!nPlazoFin
                    txtPlazo.Locked = True
                    txtTasa.Enabled = False
                    
                End If
               
           
            Else
                txtPlazo.Locked = False
                txtTasa.Enabled = True
            
            End If
        
            nPlazo = txtPlazo.Text
            
            If nPersoneria <> gPersonaNat Then
                nTasaTarifada = ConvierteTNAaTEA(oTasas.GetCapTasaInteresPF(gCapPlazoFijo, nmoneda, nTipoTasa, nPlazo, nValor, gsCodAge, , nSubProducto))
            
            Else
                nTasaTarifada = ConvierteTNAaTEA(oTasas.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, nPlazo, nValor, gsCodAge, , nSubProducto))
                               
            End If
            
        Else
            nTasaTarifada = ConvierteTNAaTEA(oTasas.GetCapTasaInteres(nProducto, nmoneda, nTipoTasa, , nValor, gsCodAge, , nSubProducto))
                    
        End If
        
        txtTasa.Text = Format$(nTasaTarifada, "#,##0.0000")
          
    End If
                
End Sub

Private Sub cboSubProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMonto.SetFocus
End If
End Sub

' Fin RIRO

Private Sub cmdCancelar_Click()
ClearScreen
txtCodigo.SetFocus
End Sub

' *** Modificado por RIRO el 20130327 *******************************************

Private Sub CmdGrabar_Click()
Dim nTasa, nMonto, nPlazo As Double
Dim nProd As COMDConstantes.Producto, nMon As COMDConstantes.Moneda
Dim sSubProducto As String ' Agregado por RIRO
Dim sComent As String, sPersona As String

nTasa = CDbl(txtTasa.Text)

If nTasa = 0 Or nTasa >= 100 Then
    MsgBox "Tasa no Válida", vbInformation, "Error"
    If txtTasa.Enabled Then
    txtTasa.SetFocus
    End If
    
    Exit Sub
End If
nMonto = txtMonto.value
If nMonto = 0 Then
    MsgBox "Monto de Apertura no Válida", vbInformation, "Error"
    txtMonto.SetFocus
    Exit Sub
End If
sComent = Trim(txtComentario.Text)
If sComent = "" Then
    MsgBox "Comentario no Válido", vbInformation, "Error"
    txtComentario.SetFocus
    Exit Sub
End If
If Not IsNumeric(txtPlazo.Text) = True Then
    MsgBox "Ingrese solo Numeros para los dia", vbInformation, "Aviso"
    Exit Sub
End If
    nProd = Trim(Right(cboProducto.Text, 5))
    If nProd = gCapPlazoFijo Then
        nPlazo = CDbl(txtPlazo.Text)
        If nPlazo = 0 Then
            MsgBox "Plazo no Válido", vbInformation, "Error"
            txtPlazo.SetFocus
            Exit Sub
        End If
    Else
        nPlazo = 0
    End If

'Por RIRO el 20130422 **
If Trim(Right(cboProducto.Text, 5)) = "233" Then

    If Trim(Right(cboMoneda.Text, 5)) = "2" Then
    
        Select Case Trim(Right(cboSubProducto.Text, 5))

            Case 1, 2, 3
               MsgBox "Sub Producto no corresponde a tipo de moneda", vbInformation, "Aviso"
               Exit Sub
           
        End Select
    
    End If

    If Trim(Right(cboSubProducto.Text, 5)) = 3 Then
    
        If txtPlazo.Text < 180 Then
            MsgBox "Plazo Fijo debe ser mayor o igual a 180 dias", vbInformation, "Aviso"
            Exit Sub
        End If
        
    ElseIf Trim(Right(cboSubProducto.Text, 5)) = 2 Then
    
        If txtPlazo.Text < 30 Then
            MsgBox "Plazo Fijo debe ser mayor o igual a 30 dias", vbInformation, "Aviso"
            Exit Sub
        End If
    
    End If

End If
' End RIRO **

If MsgBox("¿Desea grabar la solicitud de aprobación de Tasa Especial", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
    Dim nNumSolicitud As Long
    
    Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCont = Nothing
    sPersona = Trim(txtCodigo.Text)
    nProd = Trim(Right(cboProducto.Text, 5))
    sSubProducto = generarCodigoSubProducto()  ' Modificado por RIRO el 20130327
    nMon = Trim(Right(cboMoneda.Text, 2))
    nTasaTarifada = ConvierteTEAaTNA(nTasaTarifada) ' Modificado por RIRO el 20130327
    nTasa = Format$(ConvierteTEAaTNA(nTasa), "#0.0000")
    Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
        nNumSolicitud = oserv.GetNumSolcitudCapTasaEspecial()
        oserv.AgregaCapTasaEspecial nNumSolicitud, sPersona, nProd, nMon, 0, sMovNro, nTasa, sComent, nMonto, , nPlazo, , , sSubProducto, nTasaTarifada, nTasa 'Modificado por RIRO el 20130327
        MsgBox "NRO DE SOLICITUD GENERADO: " & UCase(CStr(nNumSolicitud)), vbOKOnly + vbInformation, "AVISO"
        ClearScreen
    Set oserv = Nothing
    'By Capi 21012009
     objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Solicitud", str(nNumSolicitud), gNumeroSolicitud
    'End by
    
    cmdCancelar_Click
End If
End Sub

'Agregado por RIRO EL 20130327
Private Function generarCodigoSubProducto() As String

    Dim sCodigo As String
    Dim sProducto As Integer
    
    sProducto = Trim(Right(cboProducto.Text, 3))
    If sProducto = gCapAhorros Then
        sCodigo = "2030"
    ElseIf sProducto = gCapPlazoFijo Then
        sCodigo = "2032"
    ElseIf sProducto = gCapCTS Then
        sCodigo = "2033"
    End If
    
    sCodigo = sCodigo & Trim(Right(cboSubProducto.Text, 3))
    
    generarCodigoSubProducto = sCodigo
    
End Function
'Fin RIRO *****************************************************

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Tasa Especial - Solicitud"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
IniciaCombo cboMoneda, gMoneda
IniciaCombo cboProducto, gProducto, "'230'", "'23%'"
ClearScreen
'By Capi 20012009
Set objPista = New COMManejador.Pista
gsOpeCod = gCapSolicitudTasasPreferen
'End By

' Agregado por RIRO
cboProducto_Click
' Fin RIRO

End Sub

Private Sub txtCodigo_EmiteDatos()
If txtCodigo.Text <> "" Then
    lblNombre.Caption = txtCodigo.psDescripcion
    lblDocID.Caption = txtCodigo.sPersNroDoc
    fraPersona.Enabled = False
    fraSolicitud.Enabled = True
    cmdCancelar.Enabled = True
    cmdGrabar.Enabled = True
    cboProducto.SetFocus
    nPersoneria = txtCodigo.PersPersoneria 'Agregado por RIRO el 201304161643
End If
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
If Chr(KeyAscii) = "'" Then
     KeyAscii = 0
End If
End Sub

Private Sub txtPlazo_GotFocus()
txtPlazo.SelStart = 0
txtPlazo.SelLength = Len(txtPlazo.Text)
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtTasa.Enabled Then
        txtTasa.SetFocus
    End If
    
Else
    KeyAscii = NumerosEnteros(KeyAscii)
    
End If

End Sub

Private Sub txtTasa_Change()

    If Not IsNumeric(txtTasa.Text) Then
    
        MsgBox "Debe ingresar valores numericos", vbInformation, "Aviso"
        txtTasa.Text = "0.00"
        
    Else
        If txtTasa.Text > 100 Then
        txtTasa.Text = "100.00"
        End If
        
    End If

End Sub

Private Sub txtTasa_GotFocus()
'txtTasa.MarcaTexto
With txtTasa
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

'Modificado por RIRO el 20130422
Private Sub txtTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtComentario.SetFocus
    
   
End If
'ElseIf KeyAscii <> 13 And Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46) Then
'    KeyAscii = 0
'End If

End Sub

' Agregado por RIRO el 20130327
Private Sub txtMonto_Change()

    If txtMonto.Text = "." Then
         txtMonto.Text = Format$(0, "#,##0.00")
    End If

    If cboSubProducto.ListCount > 0 Then
         cboSubProducto_Click
    End If
    
End Sub

Private Sub txtPlazo_Change()

    If txtPlazo.Text = "" Then
        txtPlazo.Text = "30"
        txtPlazo.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtPlazo.Text) Then
        MsgBox "Debe ingresar valores numericos", vbInformation, "Aviso"
        txtPlazo.Text = "30"
        txtPlazo.SetFocus
        Exit Sub
    ElseIf val(txtPlazo.Text) = 0 Then
        MsgBox "Debe ingresar valores mayores a cero", vbInformation, "Aviso"
        txtPlazo.Text = "30"
        txtPlazo.SetFocus
        Exit Sub
    End If
    
    If cboSubProducto.ListCount > 0 Then
        cboSubProducto_Click
    End If
    
End Sub
' Fin RIRO

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtPlazo.Visible Then
        txtPlazo.SetFocus
    Else
        txtTasa.SetFocus
    End If
End If
End Sub

Private Sub txtComentario_GotFocus()
txtComentario.SelStart = 0
txtComentario.SelLength = Len(txtComentario.Text)
End Sub

Private Sub txtTasa_LostFocus()
If txtTasa.Text = "" Then txtTasa.Text = "0.000"

End Sub
