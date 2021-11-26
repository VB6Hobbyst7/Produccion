VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredCalendarioDesemb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolsos Parciales"
   ClientHeight    =   7470
   ClientLeft      =   2820
   ClientTop       =   2775
   ClientWidth     =   7440
   Icon            =   "frmCredCalendarioDesemb.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Height          =   660
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   7200
      Begin VB.CommandButton cmdgrabar 
         Caption         =   "&Grabar"
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
         Height          =   345
         Left            =   105
         TabIndex        =   14
         ToolTipText     =   "Grabar Datos de Sugerencia"
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   5850
         TabIndex        =   13
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "Ca&ncelar"
         Height          =   345
         Left            =   4755
         TabIndex        =   12
         ToolTipText     =   "Limpiar la Pantalla"
         Top             =   225
         Width           =   1080
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Credito"
      ForeColor       =   &H80000006&
      Height          =   750
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "E&xaminar"
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
         Left            =   5040
         TabIndex        =   9
         ToolTipText     =   "Buscar Credito"
         Top             =   255
         Width           =   1695
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   165
         TabIndex        =   10
         Top             =   255
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   741
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Desembolsos"
      TabPicture(0)   =   "frmCredCalendarioDesemb.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   5145
         Left            =   160
         TabIndex        =   1
         Top             =   360
         Width           =   6860
         Begin VB.Frame Frame1 
            Height          =   2460
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   6615
            Begin VB.CommandButton CmdEliminarDesemb 
               Caption         =   "&Eliminar"
               Height          =   420
               Left            =   4680
               TabIndex        =   32
               Top             =   765
               Width           =   1710
            End
            Begin VB.CommandButton CmdNuevoDesemb 
               Caption         =   "&Nuevo"
               Height          =   420
               Left            =   4680
               TabIndex        =   31
               Top             =   240
               Width           =   1710
            End
            Begin SICMACT.FlexEdit FEDesPar 
               Height          =   2100
               Left            =   150
               TabIndex        =   33
               Top             =   240
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   3704
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "-Fecha-Monto-Estado-Estado"
               EncabezadosAnchos=   "350-1200-1200-0-1500"
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-1-2-X-X"
               ListaControles  =   "0-2-0"
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-R-L-L"
               FormatosEdit    =   "0-0-2"
               lbEditarFlex    =   -1  'True
               lbUltimaInstancia=   -1  'True
               lbBuscaDuplicadoText=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483630
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2415
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   6615
            Begin VB.Frame Frame2 
               Height          =   1455
               Left            =   120
               TabIndex        =   20
               Top             =   840
               Width           =   6375
               Begin VB.Frame fratipodes 
                  Caption         =   "Desembolso"
                  Enabled         =   0   'False
                  Height          =   585
                  Left            =   3000
                  TabIndex        =   21
                  Top             =   720
                  Width           =   3210
                  Begin VB.OptionButton OptDesemb 
                     Caption         =   "&Parcial"
                     Height          =   285
                     Index           =   1
                     Left            =   1575
                     TabIndex        =   23
                     Top             =   225
                     Width           =   780
                  End
                  Begin VB.OptionButton OptDesemb 
                     Caption         =   "&Total"
                     Height          =   195
                     Index           =   0
                     Left            =   315
                     TabIndex        =   22
                     Top             =   255
                     Value           =   -1  'True
                     Width           =   660
                  End
               End
               Begin MSMask.MaskEdBox TxtFecDesemb 
                  Height          =   315
                  Left            =   1570
                  TabIndex        =   24
                  Top             =   720
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label LblMontoApr 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   1580
                  TabIndex        =   34
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Desembolso :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   29
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Label LblPlazoApr 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   285
                  Left            =   3600
                  TabIndex        =   28
                  Top             =   240
                  Width           =   675
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  Caption         =   "Plazo :"
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   27
                  Top             =   240
                  Width           =   480
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "Monto :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  Caption         =   "dias"
                  Height          =   195
                  Left            =   4320
                  TabIndex        =   25
                  Top             =   240
                  Width           =   285
               End
            End
            Begin VB.Label Label1 
               Caption         =   "Codigo :"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   255
               Width           =   645
            End
            Begin VB.Label LblCodCli 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1710
               TabIndex        =   18
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Apellidos y Nombres :"
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   585
               Width           =   1515
            End
            Begin VB.Label LblNomCli 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   1710
               TabIndex        =   16
               Top             =   555
               Width           =   4710
            End
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Solicitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -72660
         TabIndex        =   7
         Top             =   75
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Solicitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   -72645
         TabIndex        =   6
         Top             =   90
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Su&gerencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -71010
         TabIndex        =   5
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Su&gerencia"
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
         Left            =   -70995
         TabIndex        =   4
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Aprobacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   -69210
         TabIndex        =   3
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Aprobacion"
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
         Left            =   -69195
         TabIndex        =   2
         Top             =   90
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCredCalendarioDesemb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredCalendarioDesemb
'***     Descripcion:
'***     Creado por:
'***     Maquina:
'***     Fecha-Tiempo:
'***     Ultima Modificacion:
'*****************************************************************************************
Option Explicit

Private bDesembParcialGenerado As Boolean
Private MatDesemb As Variant
Private bRefinanc As Boolean
Private nUltNroCalend As Integer

'Variables Desembolsos Parciales
Dim MatCalend() As String
Dim dFecDesMin As Date


Private Sub Form_Load()

    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraForm Me
    bDesembParcialGenerado = False
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    ReDim MatDesemb(0, 0)

End Sub


Private Sub cmdBuscar_Click()
    ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstAprob, gColocEstVigNorm, gColocEstVigVenc, gColocEstVigMor, gColocEstRefNorm, gColocEstRefVenc, gColocEstRefMor), "Creditos Aprobados")
    'ActxCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstAprob), "Creditos Aprobados")
    If ActxCta.NroCuenta <> "" Then
        Call ActxCta_KeyPress(13)
    Else
        ActxCta.CMAC = gsCodCMAC
        ActxCta.Age = gsCodAge
        ActxCta.SetFocusProd
        ActxCta.Enabled = True
    End If
End Sub


Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CargaDatos(ActxCta.NroCuenta) Then
            ActxCta.Enabled = False
            cmdgrabar.Enabled = True
            cmdbuscar.Enabled = False
        Else
            ActxCta.Enabled = True
            cmdgrabar.Enabled = False
            cmdbuscar.Enabled = True
            'MsgBox "No se Encontro el Credito", vbExclamation, "Aviso"
        End If
    End If
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean

Dim R As ADODB.Recordset
'Dim oCredito As COMDCredito.DCOMCredito
Dim oNCredito As COMNCredito.NCOMCredito
'Dim oCalend As COMDCredito.DCOMCalendario
Dim RDes As ADODB.Recordset

    On Error GoTo ErrorCargaDatos
    
 '   Set oCredito = New COMDCredito.DCOMCredito
 '   Set R = oCredito.RecuperaDatosAprobados(psCtaCod)
    Set oNCredito = New COMNCredito.NCOMCredito
    Call oNCredito.CargarDatosCalendarioDesembolso(psCtaCod, R, RDes, bRefinanc)
    Set oNCredito = Nothing
    
    If Not R.BOF And Not R.EOF Then
        CargaDatos = True
        
        nUltNroCalend = R!nMaxCalend
        'Set oNCredito = New COMNCredito.NCOMCredito
        'bRefinanc = oNCredito.EsRefinanciado(ActxCta.NroCuenta)
        'Set oNCredito = Nothing
        
        LblCodCli.Caption = R!cPersCod
        LblNomCli.Caption = R!cPersNombre
        
        LblPlazoApr.Caption = R!nPlazo
        'LblCuotasApr.Caption = R!nCuotas
        LblMontoApr.Caption = Format(IIf(IsNull(R!nMonto), 0, R!nMonto), "#0.00")
        
        If Not IsNull(R!dVigencia) Then
           TxtFecDesemb.Text = Format(R!dVigencia, "dd/mm/yyyy")
        Else
           TxtFecDesemb.Text = Format(gdFecSis, "dd/mm/yyyy")
        End If
      
        'Tipo de Desembolso
        If R!nTipoDesembolso = gColocTiposDesembolsoTotal Then
           OptDesemb(0).value = True
        End If
        If R!nTipoDesembolso = gColocTiposDesembolsoParcial Then
           OptDesemb(1).value = True
        End If
        
        'Tipo de Desembolso
        If R!nTipoDesembolso = gColocTiposDesembolsoTotal Then
           MatDesemb = ""
           CargaDatos = False
           
           FEDesPar.lbEditarFlex = False
           CmdEliminarDesemb.Enabled = False
           CmdNuevoDesemb.Enabled = False
        
        Else
        '    Set oCalend = New COMDCredito.DCOMCalendario
        '    Set RDes = oCalend.RecuperaCalendarioDesemb(ActxCta.NroCuenta)
        '    Set oCalend = Nothing
            
            If RDes.RecordCount > 0 Then
               ReDim MatDesemb(RDes.RecordCount, 4)
            
               Do While Not RDes.EOF
                  MatDesemb(RDes.Bookmark - 1, 0) = Format(RDes!dvenc, "dd/mm/yyyy")
                  MatDesemb(RDes.Bookmark - 1, 1) = Format(RDes!nCapital, "#0.00")
                  MatDesemb(RDes.Bookmark - 1, 2) = RDes!nColocCalendEstado
                  MatDesemb(RDes.Bookmark - 1, 3) = RDes!cColocCalendEstado
                  RDes.MoveNext
               Loop
                
               RDes.Close
                
               'AQUI verificar que pasa cuando no trae nada desembolsos
               Call CargaDesembolsos
               
               CmdEliminarDesemb.Enabled = True
               CmdNuevoDesemb.Enabled = True
               
            Else
               RDes.Close
               
               MatDesemb = ""
               CargaDatos = False
                
               FEDesPar.lbEditarFlex = False
            
               CmdEliminarDesemb.Enabled = False
               CmdNuevoDesemb.Enabled = False
            End If
            
        End If
        
        R.Close
        
    Else
        CargaDatos = False
    End If
    Set R = Nothing
    
    'Set oCredito = Nothing
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"
End Function


Private Function ValidaDatosDesembolsos() As Boolean
Dim i As Integer
Dim dFecTemp As Date
    
    dFecTemp = dFecDesMin
    ValidaDatosDesembolsos = True
    
    If Trim(FEDesPar.TextMatrix(1, 0)) = "" Then
       ReDim MatCalend(0, 0)
       ValidaDatosDesembolsos = False
       MsgBox "Debe Ingresar el Calendario de Desembolsos", vbInformation, "Aviso"
       Exit Function
    End If
        
    'VERIFICAR SI ESTA DEMAS
    ReDim MatCalend(FEDesPar.Rows - 1, 4)
    For i = 1 To FEDesPar.Rows - 1
        MatCalend(i - 1, 0) = FEDesPar.TextMatrix(i, 1)
        MatCalend(i - 1, 1) = FEDesPar.TextMatrix(i, 2)
        MatCalend(i - 1, 2) = FEDesPar.TextMatrix(i, 3)
        MatCalend(i - 1, 3) = FEDesPar.TextMatrix(i, 4)
    Next i
    
    MatDesemb = MatCalend
    '''''''''''''''''''''''''''''''''''''''''
    
    Dim nSumaDesPar As Double
    Dim MonDesAnt As Double
    
    MonDesAnt = Format(LblMontoApr, "#0.00")
    If UBound(MatDesemb) > 0 Then
        nSumaDesPar = 0
        
        For i = 0 To UBound(MatDesemb) - 1
            nSumaDesPar = nSumaDesPar + CDbl(MatDesemb(i, 1))
        Next i
        
        If MonDesAnt <> nSumaDesPar Then
           ValidaDatosDesembolsos = False
           MsgBox "La Suma de los Desembolsos no es igual al Monto del Crédito", vbInformation, "Aviso"
           FEDesPar.SetFocus
           Exit Function
        End If
        bDesembParcialGenerado = True
    Else
        bDesembParcialGenerado = False
    End If

    
    If Trim(FEDesPar.TextMatrix(1, 0)) = "" Then
        Exit Function
    End If
    For i = 1 To FEDesPar.Rows - 1
        If ValidaFecha(FEDesPar.TextMatrix(i, 1)) <> "" Then
            ValidaDatosDesembolsos = False
            MsgBox ValidaFecha(FEDesPar.TextMatrix(i, 1)), vbInformation, "Aviso"
            FEDesPar.Row = i
            FEDesPar.Col = 1
            FEDesPar.SetFocus
            Exit Function
        End If
        If Trim(FEDesPar.TextMatrix(i, 2)) = "" Then
            FEDesPar.TextMatrix(i, 2) = "0.00"
        End If
        If CDbl(FEDesPar.TextMatrix(i, 2)) <= 0 Then
            ValidaDatosDesembolsos = False
            MsgBox "Monto de Desembolso debe ser mayor que Cero", vbInformation, "Aviso"
            FEDesPar.Row = i
            FEDesPar.Col = 2
            FEDesPar.SetFocus
            Exit Function
        End If
        'If CDate(FEDesPar.TextMatrix(i, 1)) <= dFecTemp Or CDate(FEDesPar.TextMatrix(i, 1)) < dFecDesMin Then
        If CDate(FEDesPar.TextMatrix(i, 1)) <= dFecTemp And CDate(FEDesPar.TextMatrix(i, 1)) < dFecDesMin Then
            ValidaDatosDesembolsos = False
            MsgBox "Fecha de Desembolso No puede ser Menor o Igual que la Fecha de Desembolso Anterior", vbInformation, "Aviso"
            FEDesPar.Row = i
            FEDesPar.Col = 1
            FEDesPar.SetFocus
            Exit Function
        End If
        dFecTemp = CDate(FEDesPar.TextMatrix(i, 1))
    Next i
End Function


Private Sub CargaDesembolsos()
Dim nSumaDesPar As Double
Dim i As Integer
Dim MonDesAnt As Double
    
    MonDesAnt = CDbl(LblMontoApr.Caption)
    MatDesemb = Inicio(CDate(TxtFecDesemb.Text), MatDesemb)
    If UBound(MatDesemb) > 0 Then
        nSumaDesPar = 0
        For i = 0 To UBound(MatDesemb) - 1
            nSumaDesPar = nSumaDesPar + CDbl(MatDesemb(i, 1))
        Next i
        If MonDesAnt <> nSumaDesPar Then
        End If
        bDesembParcialGenerado = True
    Else
        bDesembParcialGenerado = False
    End If
End Sub


Private Function Inicio(ByVal pdFecDesMin As Date, ByVal pMatDesPar As Variant) As Variant
Dim i As Integer
    'On Error Resume Next
    
    'LimpiaFlex FEDesPar
    If OptDesemb(0).value = True Then
       Set pMatDesPar = Nothing
    End If

    'If Not pMatDesPar Is Nothing Then
        If UBound(pMatDesPar) > 0 Then
            For i = 0 To UBound(pMatDesPar) - 1
                FEDesPar.AdicionaFila
                FEDesPar.TextMatrix(i + 1, 1) = pMatDesPar(i, 0)
                FEDesPar.TextMatrix(i + 1, 2) = pMatDesPar(i, 1)
                FEDesPar.TextMatrix(i + 1, 3) = pMatDesPar(i, 2)
                FEDesPar.TextMatrix(i + 1, 4) = pMatDesPar(i, 3)
                
                If pMatDesPar(i, 2) = 1 Then
                   Call FEDesPar.BackColorRow(&HFFC0C0)
                Else
                   Call FEDesPar.BackColorRow(vbWhite)
                End If
                
            Next i
        End If
   ' End If
    dFecDesMin = pdFecDesMin
    'FEDesPar.lbEditarFlex = True
    ReDim MatCalend(0, 0)
    
    
    i = 0
        If Trim(FEDesPar.TextMatrix(1, 0)) = "" Then
            ReDim MatCalend(0, 0)
            FEDesPar.lbEditarFlex = False
        Else
        FEDesPar.lbEditarFlex = True
        ReDim MatCalend(FEDesPar.Rows - 1, 4)
        For i = 1 To FEDesPar.Rows - 1
            MatCalend(i - 1, 0) = FEDesPar.TextMatrix(i, 1)
            MatCalend(i - 1, 1) = FEDesPar.TextMatrix(i, 2)
            MatCalend(i - 1, 2) = FEDesPar.TextMatrix(i, 3)
            MatCalend(i - 1, 3) = FEDesPar.TextMatrix(i, 4)
        Next i
       End If
    Inicio = MatCalend
    
End Function


Private Sub CmdEliminarDesemb_Click()
    If CInt(FEDesPar.TextMatrix(FEDesPar.Row, 3)) = gColocCalendEstadoPagado Then
       MsgBox "El Desembolso ha sido realizado", vbInformation, "Aviso"
       Exit Sub
    End If
    
    If MsgBox("Se va a Eliminar el Desembolso de la Fecha : " & FEDesPar.TextMatrix(FEDesPar.Row, 1) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call FEDesPar.EliminaFila(FEDesPar.Row)
    End If
End Sub

Private Sub CmdNuevoDesemb_Click()
    FEDesPar.lbEditarFlex = True
    FEDesPar.AdicionaFila
    FEDesPar.TextMatrix(FEDesPar.Row, 3) = 0
    FEDesPar.TextMatrix(FEDesPar.Row, 4) = "PENDIENTE"
    FEDesPar.SetFocus
    
End Sub


Private Sub Optdesemb_Click(Index As Integer)
    If Index = 1 Then
        ReDim MatDesemb(0, 0)
        bDesembParcialGenerado = False
    Else
        bDesembParcialGenerado = False
    End If
End Sub

Private Sub LimpiarPantalla()
    bRefinanc = False
    bDesembParcialGenerado = False
    ReDim MatDesemb(0, 0)
    LimpiaControles Me
    LimpiaFlex FEDesPar
    'FEDesPar.lbEditarFlex = False
    ActxCta.NroCuenta = ""
    ActxCta.Enabled = True
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    ActxCta.SetFocusProd
    TxtFecDesemb.Text = "__/__/____"
    OptDesemb(0).value = True
    cmdgrabar.Enabled = False
    cmdbuscar.Enabled = True
End Sub

Private Function ValidaDatos() As Boolean
Dim sCad As String
Dim oCred As COMNCredito.NCOMCredito

    ValidaDatos = True
    
    'Verificacion de Niveles de Aprobacion
    Set oCred = New COMNCredito.NCOMCredito
    If Not oCred.ValidaNivelAprUsuario(gsCodUser, CDbl(LblMontoApr.Caption), Mid(ActxCta.NroCuenta, 9, 1), Mid(ActxCta.NroCuenta, 6, 3), bRefinanc) Then
        MsgBox "No Tiene el Nivel de Aprobacion Adecuado para este Proceso", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
  
    'Valida Fecha de Desembolso
    sCad = ValidaFecha(TxtFecDesemb.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
End Function

Private Sub cmdGrabar_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Dim pnTipoCuota As Integer
Dim sError As String
Dim sMontosARef As Variant

    If Not ValidaDatos Then
        Exit Sub
    End If
            
    If Not ValidaDatosDesembolsos Then
        Exit Sub
    End If
    
    If MsgBox("Se Va A Grabar los Datos, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then

       Set oNCredito = New COMNCredito.NCOMCredito
        
       sError = oNCredito.CalendarioDesembolsos(ActxCta.NroCuenta, CDbl(LblMontoApr.Caption), _
       IIf(OptDesemb(0).value, gColocTiposDesembolsoTotal, gColocTiposDesembolsoParcial), nUltNroCalend + 1, MatDesemb)
       Set oNCredito = Nothing
       If sError <> "" Then
          MsgBox sError, vbInformation, "Aviso"
       Else
          Call ImprimirCalendarioDesemb
          Call LimpiarPantalla
       End If
    End If
End Sub


Private Sub ImprimirCalendarioDesemb()

Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim sCadImp As String
Dim sCadImp_2 As String
Dim Prev As Previo.clsPrevio

    On Error GoTo ErrorImprimirCalendarioDesemb

            Set oCredDoc = New COMNCredito.NCOMCredDoc
            Set Prev = New clsPrevio
            sCadImp = oCredDoc.ImprimeCalendarioDesemb(ActxCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, CDbl(LblMontoApr.Caption), False, , gsNomCmac, False, 0)
            sCadImp_2 = ""

            Prev.Show sCadImp & sCadImp_2, "", False
            Set Prev = Nothing
            Set oCredDoc = Nothing
            
    Exit Sub

ErrorImprimirCalendarioDesemb:
        MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub


Private Sub cmdCancelar_Click()
    Call LimpiarPantalla
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

