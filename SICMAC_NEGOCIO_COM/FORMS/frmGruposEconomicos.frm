VERSION 5.00
Begin VB.Form frmGruposEconomicos 
   Caption         =   "Grupos Economicos"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmGruposEconomicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReporte20 
      Caption         =   "Reporte 20"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdReporte19 
      Caption         =   "Reporte 19"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   7200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupos Economicos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      Begin VB.CommandButton cmdNuevoGrupo 
         Caption         =   "Grupos"
         Height          =   375
         Left            =   6360
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Empresas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   7695
         Begin VB.TextBox txtBuscaEmpresa 
            Height          =   285
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   6015
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Height          =   375
            Left            =   6480
            TabIndex        =   13
            Top             =   4080
            Width           =   855
         End
         Begin VB.CommandButton cmdNuevaEmpresa 
            Caption         =   "Nuevo"
            Height          =   375
            Left            =   6480
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin VB.Frame Frame3 
            Caption         =   "Vinculados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   0
            TabIndex        =   5
            Top             =   2760
            Width           =   7695
            Begin VB.CommandButton cmdDetalleVinculados 
               Caption         =   "Relación"
               Height          =   375
               Left            =   6480
               TabIndex        =   10
               Top             =   840
               Width           =   855
            End
            Begin VB.CommandButton cmdCargoVinculado 
               Caption         =   "Cargos"
               Height          =   375
               Left            =   6480
               TabIndex        =   8
               Top             =   360
               Width           =   855
            End
            Begin SICMACT.FlexEdit FEVinculados 
               Height          =   2655
               Left            =   240
               TabIndex        =   6
               Top             =   360
               Width           =   6135
               _extentx        =   10821
               _extenty        =   4683
               cols0           =   10
               highlight       =   2
               allowuserresizing=   3
               rowsizingmode   =   1
               encabezadosnombres=   "-RL-CodVinculado-Vinculado-Acciones-CodCargo-Cargo-CodOtroCargo-OtroCargo-Gestion"
               encabezadosanchos=   "0-400-1500-3000-1000-0-1200-0-1200-1200"
               font            =   "frmGruposEconomicos.frx":030A
               font            =   "frmGruposEconomicos.frx":0336
               font            =   "frmGruposEconomicos.frx":0362
               font            =   "frmGruposEconomicos.frx":038E
               font            =   "frmGruposEconomicos.frx":03BA
               fontfixed       =   "frmGruposEconomicos.frx":03E6
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               backcolorcontrol=   -2147483643
               columnasaeditar =   "X-X-X-X-X-X-X-X-X-X"
               listacontroles  =   "0-0-0-0-0-0-0-0-0-0"
               backcolor       =   -2147483644
               encabezadosalineacion=   "C-C-C-L-C-R-C-R-L-C"
               formatosedit    =   "0-0-0-0-0-3-0-3-0-0"
               selectionmode   =   1
               rowheight0      =   300
               forecolorfixed  =   -2147483630
               cellbackcolor   =   -2147483644
            End
         End
         Begin SICMACT.FlexEdit FEEmpresas 
            Height          =   2055
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   6135
            _extentx        =   10821
            _extenty        =   3625
            cols0           =   3
            encabezadosnombres=   "-CodEmpresa-Nombre Empresa"
            encabezadosanchos=   "0-1500-4000"
            font            =   "frmGruposEconomicos.frx":0414
            font            =   "frmGruposEconomicos.frx":0440
            font            =   "frmGruposEconomicos.frx":046C
            font            =   "frmGruposEconomicos.frx":0498
            font            =   "frmGruposEconomicos.frx":04C4
            fontfixed       =   "frmGruposEconomicos.frx":04F0
            columnasaeditar =   "X-X-X"
            listacontroles  =   "0-0-0"
            backcolor       =   -2147483644
            encabezadosalineacion=   "C-L-L"
            formatosedit    =   "0-0-0"
            rowheight0      =   300
            cellbackcolor   =   -2147483644
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboGruposEconomicos 
         Height          =   315
         ItemData        =   "frmGruposEconomicos.frx":051E
         Left            =   240
         List            =   "frmGruposEconomicos.frx":0520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmGruposEconomicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoOperacion As Integer '0 Nuevo...1 Modificar
Dim loDatos As ADODB.Recordset
Dim loGestionVinculados As ADODB.Recordset
Dim loGestion As ADODB.Recordset
Dim lnGrupoEco As Integer
Dim i, J As Integer
Dim nPost As Integer
Dim sMatPersonas() As String
Dim sMatGestion() As String
Dim sMatVinculados() As String
Dim K As Integer
Dim kg As Integer
Dim nCantVin As Integer
Dim lsPersCodEmpre As String
Dim lsPersCodVincu As String
Dim lbBuscaTexto As Boolean 'JAME20140303
Dim nFilaGreen As Long 'PTI1 17/08/2018
Dim loDatosVinculados As ADODB.Recordset 'PTI1 17/08/2018
Private Sub cboGruposEconomicos_Click()

'*** PEAC 20130619 - se traslado esta validacion del boton Mostrar
Dim lnPersona As COMNPersona.NCOMPersona
Set lnPersona = New COMNPersona.NCOMPersona

'EJVG20120419 Para restringir adición de personas a grupos


Dim lbPermiso As Boolean

If Right(cboGruposEconomicos.Text, 4) <> lnGrupoEco Then
LimpiaFlex FEEmpresas 'ADD PTI1 INFORME n°002-2017-AC-TI/CMACM
LimpiaFlex FEVinculados 'ADD PTI1 INFORME n°002-2017-AC-TI/CMACM
End If

lnGrupoEco = Right(cboGruposEconomicos.Text, 4)
lbPermiso = lnPersona.TienePermisoAdicionarPersonaGrupoEco(lnGrupoEco, gsGruposUser)
cmdNuevaEmpresa.Enabled = lbPermiso
cmdCargoVinculado.Enabled = lbPermiso
cmdDetalleVinculados.Enabled = lbPermiso
cmdEditar.Enabled = lbPermiso



End Sub

Private Sub cmdDetalleVinculados_Click()
    If lsPersCodEmpre <> "" And lsPersCodVincu <> "" Then
        Call frmGruposEconomicosGestion.iniciar(lnGrupoEco, lsPersCodEmpre, lsPersCodVincu)
        Call FEEmpresas_Click 'add pti1 INFORME n°002-2017-AC-TI/CMACM
        'Call cmdMostrar_Click 'comentado por pti1 20/08/2018
    Else
         MsgBox "Seleccione nuevamente datos vacíos", vbInformation, "Aviso" 'ADD PTI1 17/08/2018 INFORME n°002-2017-AC-TI/CMACM
    End If
End Sub

Private Sub cmdEditar_Click()
   If lsPersCodEmpre <> "" And lsPersCodVincu <> "" Then
    Call frmGrupoEcoEmpresa.Modificar(Right(cboGruposEconomicos.Text, 4), Left(cboGruposEconomicos.Text, 40), lsPersCodEmpre, lsPersCodVincu)
    'Call cmdMostrar_Click 'COMENTADO POR PTI1 20/08/2018 INFORME n°002-2017-AC-TI/CMACM
    Call FEEmpresas_Click 'add pti1 20/08/2018
    'lsPersCodEmpre = "" 'comentado por pti1 20/08/2018
    'lsPersCodVincu = "" 'comentado por pti1 20/08/2018
    Else
     MsgBox "Seleccione nuevamente datos vacíos", vbInformation, "Aviso" 'ADD PTI1 17/08/2018 NFORME n°002-2017-AC-TI/CMACM
    End If
End Sub


Private Sub cmdMostrar_Click()
nFilaGreen = -1 'ADD PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM

Dim lnPersona As COMNPersona.NCOMPersona
Set lnPersona = New COMNPersona.NCOMPersona
If Trim(cboGruposEconomicos.Text) = "" Then
    MsgBox "Seleccionar un grupo economico ", vbCritical
    Exit Sub
End If

Screen.MousePointer = 11

lnGrupoEco = Right(cboGruposEconomicos.Text, 4)
'JAME20140303
    If lnGrupoEco = 2 Then
        Me.txtBuscaEmpresa.Visible = True
    Else
      Me.txtBuscaEmpresa.Visible = False
    End If
    If lbBuscaTexto = True Then
        Call lnPersona.ObtenerDatosSoloGrupoEconomicos(loDatos, loGestionVinculados, loGestion, Right(cboGruposEconomicos.Text, 4), True, Me.txtBuscaEmpresa.Text) 'ADD PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
        'Call lnPersona.ObtenerDatosGrupoEconomicos(loDatos, loGestionVinculados, loGestion, Right(cboGruposEconomicos.Text, 4), True, Me.txtBuscaEmpresa.Text) 'COMENTADO POR PTI1 17/08/2018
    Else
        Call lnPersona.ObtenerDatosSoloGrupoEconomicos(loDatos, loGestionVinculados, loGestion, Right(cboGruposEconomicos.Text, 4)) 'ADD PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
        'Call lnPersona.ObtenerDatosGrupoEconomicos(loDatos, loGestionVinculados, loGestion, Right(cboGruposEconomicos.Text, 4)) 'COMENTADO POR PTI1 17/08/2018
    End If
'JAME FIN
'Call lnPersona.ObtenerDatosGrupoEconomicos(loDatos, loGestionVinculados, loGestion, Right(cboGruposEconomicos.Text, 4))'JAME20140303 - COMENTADO
Call MostrarEmpresar
'Call PrimerFEEmpresas  'COMENTADO POR PTI1 17/08/2018

'*** PEAC 20130619 - se traslado esta validacion al combo de grupos
''EJVG20120419 Para restringir adición de personas a grupos
'Dim lbPermiso As Boolean
'lbPermiso = lnPersona.TienePermisoAdicionarPersonaGrupoEco(lnGrupoEco, gsGruposUser)
'cmdNuevaEmpresa.Enabled = lbPermiso
'cmdCargoVinculado.Enabled = lbPermiso
'cmdDetalleVinculados.Enabled = lbPermiso
'cmdEditar.Enabled = lbPermiso

Screen.MousePointer = 0

End Sub
Private Sub MostrarEmpresar()
'   K = 0 'COMENTADO PTI1 17/08/2018 NFORME n°002-2017-AC-TI/CMACM
'    Dim nEncontrado As Integer
'**** INICIO COMENTADO POR PTI1 17/08/2018
'        If nPost > 0 Then
'            For i = 1 To nPost
'                FEEmpresas.EliminaFila (1)
'            Next i
'        End If
'        If nCantVin > 0 Then
'            For i = 0 To nCantVin
'               FEVinculados.EliminaFila (1)
'            Next i
'        End If
     
'        nCantVin = 0
'       nPost = 0
'        j = 0
'        K = 0
'        kg = 0
'**** FIN comentado
 
        If loDatos.EOF Or loDatos.BOF Then
        MsgBox "No existen registros en este grupo", vbInformation, "Aviso" 'AGREGADO POR PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
        LimpiaFlex FEEmpresas 'AGREGADO POR PTI1 (17/08/2018)
        LimpiaFlex FEVinculados 'AGREGADO POR PTI1 (17/08/2018)
            Exit Sub
        End If
'        nEncontrado = 0 'COMENTADO PTI1 17/08/2018
        '*'ADD PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
        Me.FEEmpresas.Clear
        Me.FEEmpresas.FormaCabecera
        Me.FEEmpresas.rsFlex = loDatos
        Me.FEEmpresas.BackColor = (vbWhite)
        
        
        
    
         
    
        LimpiaFlex FEVinculados 'AGREGADO POR PTI1 (17/08/2018)
        lsPersCodVincu = ""
        lsPersCodEmpre = ""
        '*'FIN AGREGADO
        
        '**** INICIO COMENTADO POR PTI1 17/08/2018
'        Do Until loDatos.EOF
'            J = J + 1
'            For i = 1 To J - 1
'                If sMatPersonas(1, i) = loDatos!cPersCod Then
'                    nEncontrado = 1
'                End If
'            Next i
'            If nEncontrado = 0 Then
'                k = k + 1
'                FEEmpresas.AdicionaFila
'                FEEmpresas.TextMatrix(k, 0) = ""
'                FEEmpresas.TextMatrix(k, 1) = loDatos!cPersCod
'                FEEmpresas.TextMatrix(k, 2) = loDatos!cPersEmpresa
'            End If
'            nPost = J
'            ReDim Preserve sMatPersonas(1 To 13, 1 To J)
'
'            sMatPersonas(1, J) = loDatos!cPersCod
'            sMatPersonas(2, J) = loDatos!cPersEmpresa
'            sMatPersonas(3, J) = loDatos!nRepresentanteLegal
'            sMatPersonas(4, J) = loDatos!cPersCodOtro
'            sMatPersonas(5, J) = loDatos!cPersVinculado
'            sMatPersonas(6, J) = loDatos!nPorcenOtro
'
'            sMatPersonas(7, J) = loDatos!nCargo
'            sMatPersonas(8, J) = loDatos!cDesCargo
'            sMatPersonas(9, J) = loDatos!nCargoOtro
'            sMatPersonas(10, J) = loDatos!cDesOtroCargo
'            nEncontrado = 0
'       loDatos.MoveNext
'    Loop
 '****
 
'    If loGestionVinculados.EOF Or loGestionVinculados.BOF Then
'            Exit Sub
'    End If
'
'    Do Until loGestionVinculados.EOF
'        kg = kg + 1
'        ReDim Preserve sMatGestion(1 To 2, 1 To kg)
'
'        sMatGestion(1, kg) = loGestionVinculados!cPersCod
'        sMatGestion(2, kg) = loGestionVinculados!cRelacionGestion
'        loGestionVinculados.MoveNext
'    Loop  FIN COMENTADO POR PTI1

End Sub
Private Sub ObtenerVinculados(lsCodPersona As String)
    
    Dim nEncontrado As Integer
    Dim sGestion As String
    Dim ig As Integer
    
    On Error GoTo ERRORForm 'AGREGADO PTI1 17/08/2018

    
    'INICIO COMENTAOD POR PTI1 (17/08/2018)
'        If nCantVin > 0 Then
'            For i = 0 To nCantVin
'               FEVinculados.EliminaFila (1)
'            Next i
'        End If
    'FIN COMENTADO PTI1 (17/08/2018)
    
        nEncontrado = 0
        nCantVin = 0
        
        'AGREGADO POR PTI1 (17/08/2018)
        LimpiaFlex FEVinculados 'AGREGADO POR PTI1 (17/08/2018)
        Dim nInicioFin As Integer
        nInicioFin = 0
        
          For nInicioFin = 1 To loDatosVinculados.RecordCount
            sGestion = ""
           If loDatosVinculados!cPersCod = lsCodPersona Then
                lsPersCodEmpre = lsCodPersona
                nEncontrado = 1
            End If
            
            If nEncontrado = 1 Then
             nCantVin = nCantVin + 1
             For ig = 1 To kg
                    
                    If Trim(loDatosVinculados!cPersCodOtro) = sMatGestion(1, ig) And lsPersCodEmpre = sMatGestion(3, ig) Then
                    sGestion = sGestion & sMatGestion(2, ig) & ","
                    End If
             Next ig
             If Len(sGestion) > 0 Then
                sGestion = Mid(sGestion, 1, Len(sGestion) - 1)
             End If
             
             
            End If
                FEVinculados.AdicionaFila
                FEVinculados.TextMatrix(nCantVin, 0) = nCantVin
                FEVinculados.TextMatrix(nCantVin, 1) = IIf(Trim(loDatosVinculados!nRepresentanteLegal) = "0", "N", "S")
                FEVinculados.TextMatrix(nCantVin, 2) = loDatosVinculados!cPersCodOtro
                FEVinculados.TextMatrix(nCantVin, 3) = loDatosVinculados!cPersVinculado
                FEVinculados.TextMatrix(nCantVin, 4) = loDatosVinculados!nPorcenOtro
                FEVinculados.TextMatrix(nCantVin, 5) = loDatosVinculados!nCargo
                FEVinculados.TextMatrix(nCantVin, 6) = loDatosVinculados!cDesCargo
                FEVinculados.TextMatrix(nCantVin, 7) = loDatosVinculados!nCargoOtro
                FEVinculados.TextMatrix(nCantVin, 8) = loDatosVinculados!cDesOtroCargo
                FEVinculados.TextMatrix(nCantVin, 9) = sGestion
                
           
                loDatosVinculados.MoveNext
            nEncontrado = 0
         Next nInicioFin
         'FIN AGREGADO PTI1
        
        
'        For i = 1 To J
'            sGestion = ""
'            If sMatPersonas(1, i) = lsCodPersona Then
'                lsPersCodEmpre = lsCodPersona
'                nEncontrado = 1
'            End If
'            If nEncontrado = 1 Then
'                nCantVin = nCantVin + 1
'                For ig = 1 To kg
'                    If Trim(sMatPersonas(4, i)) = sMatGestion(1, ig) Then
'                    sGestion = sGestion & sMatGestion(2, ig) & ","
'                    End If
'                Next ig
'                If Len(sGestion) > 0 Then
'                sGestion = Mid(sGestion, 1, Len(sGestion) - 1)
'                End If
'                FEVinculados.AdicionaFila
'                ReDim Preserve sMatVinculados(1 To 9, 1 To nCantVin)
'
'                sMatVinculados(1, nCantVin) = IIf(Trim(sMatPersonas(3, i)) = "0", "N", "S")
'                sMatVinculados(2, nCantVin) = sMatPersonas(4, i)
'                sMatVinculados(3, nCantVin) = sMatPersonas(5, i)
'                sMatVinculados(4, nCantVin) = sMatPersonas(6, i)
'                sMatVinculados(5, nCantVin) = sMatPersonas(7, i)
'                sMatVinculados(6, nCantVin) = sMatPersonas(8, i)
'                sMatVinculados(7, nCantVin) = sMatPersonas(9, i)
'                sMatVinculados(8, nCantVin) = sMatPersonas(10, i)
'                sMatVinculados(9, nCantVin) = sGestion
'
'
'                FEVinculados.TextMatrix(nCantVin, 0) = ""
'                FEVinculados.TextMatrix(nCantVin, 1) = IIf(Trim(sMatPersonas(3, i)) = "0", "N", "S")
'                FEVinculados.TextMatrix(nCantVin, 2) = sMatPersonas(4, i)
'                FEVinculados.TextMatrix(nCantVin, 3) = sMatPersonas(5, i)
'                FEVinculados.TextMatrix(nCantVin, 4) = sMatPersonas(6, i)
'                FEVinculados.TextMatrix(nCantVin, 5) = sMatPersonas(7, i)
'                FEVinculados.TextMatrix(nCantVin, 6) = sMatPersonas(8, i)
'                FEVinculados.TextMatrix(nCantVin, 7) = sMatPersonas(9, i)
'                FEVinculados.TextMatrix(nCantVin, 8) = sMatPersonas(10, i)
'                FEVinculados.TextMatrix(nCantVin, 9) = sGestion
'            End If
'            nEncontrado = 0
'       Next i
        'FIN COMENTADO PTI1
         Exit Sub
    
ERRORForm: 'AGREGADO PTI1 17/08/2018
    MsgBox "Intente nuevamente, si el error persiste comuniquese con TI", vbInformation, "Aviso"
End Sub

Private Sub cmdNuevaEmpresa_Click()
    If Trim(cboGruposEconomicos.Text) = "" Then
        MsgBox "Seleccionar un grupo economico ", vbCritical
        Exit Sub
    End If
    Call frmGrupoEcoEmpresa.Nuevo(Right(cboGruposEconomicos.Text, 4), Left(cboGruposEconomicos.Text, 40))
    Call cmdMostrar_Click
    lsPersCodEmpre = ""
    lsPersCodVincu = ""
End Sub

Private Sub cmdCargoVinculado_Click()
   
    If lsPersCodEmpre <> "" And lsPersCodVincu <> "" Then
   
        Call frmGruposEconomicosCargos.Nuevo(lnGrupoEco, lsPersCodEmpre, lsPersCodVincu)
         Call FEEmpresas_Click 'add pti1 INFORME n°002-2017-AC-TI/CMACM
        'Call cmdMostrar_Click 'comentado por pti1 17/08/2018
    Else
        MsgBox "Seleccione nuevamente datos vacíos", vbInformation, "Aviso" 'ADD PTI1 17/08/2018
    End If
End Sub

Private Sub cmdNuevoGrupo_Click()
        Call frmGruposEnocomicosNuevo.iniciar
        Call CargarGrupos
        LimpiaFlex FEEmpresas 'AGREGADO PTI1 17/082018 INFORME n°002-2017-AC-TI/CMACM
        LimpiaFlex FEVinculados 'AGREGADO PTI1 17/082018
        lsPersCodEmpre = "" 'AGREGADO PTI1 17/082018
        lsPersCodVincu = "" 'AGREGADO PTI1 17/082018
End Sub
Private Sub CargarGrupos()
    cboGruposEconomicos.Clear
    Dim oCons As COMDPersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    Set oCons = New COMDPersona.DCOMGrupoE
    Set rs = oCons.ListarGrupoEconomico(1)
    Call Llenar_Combo_con_Recordset(rs, cboGruposEconomicos)
End Sub

Private Sub cmdReporte19_Click()
    Call Reporte19GruposEconomicos(gdFecSis)
End Sub

Private Sub cmdReporte20_Click()
    Call Reporte20RiesgoUnico(gdFecSis)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
'ALPA 20100106************************************
Private Sub PrimerFEEmpresas()
'Dim nCol As Integer
'Dim nPosFila As Integer
'    nCol = FEEmpresas.Col
'    nPosFila = FEEmpresas.Row
'    For i = 1 To K
'    FEEmpresas.Row = i
'    FEEmpresas.BackColorRow (vbWhite)
'    Next i
'    FEEmpresas.Row = nPosFila
''    If Trim(FEEmpresas.TextMatrix(FEEmpresas.Row, 1)) <> "" Then
'        FEEmpresas.Col = 1
'        If FEEmpresas.CellBackColor = vbGreen Then
'            FEEmpresas.BackColorRow (vbWhite)
'            FEEmpresas.TextMatrix(FEEmpresas.Row, 0) = "."
'            lsPersCodEmpre = ""
'            lsPersCodVincu = ""
'        Else
'
'            FEEmpresas.BackColorRow (vbGreen)
'            FEEmpresas.TextMatrix(FEEmpresas.Row, 0) = "R" '
'            lsPersCodVincu = ""
'        End If
'If FEEmpresas.TextMatrix(FEEmpresas.Row, 1) <> "" Then
'    Call ObtenerVinculados(FEEmpresas.TextMatrix(FEEmpresas.Row, 1))
'End If
FEEmpresas.row = 1
Call FEEmpresas_Click
End Sub
'*************************************************
Private Sub FEEmpresas_Click()
Dim nCol As Integer
Dim nPosFila As Integer

    nCol = FEEmpresas.col
    nPosFila = FEEmpresas.row
    'PTI1 (17/08/2018) COMENTADO INFORME n°002-2017-AC-TI/CMACM
'    For i = 1 To k
'    FEEmpresas.row = i
'    FEEmpresas.BackColorRow (vbWhite)
'    Next i
    'FIN COMENTADO PTI1
    
    'AGREGADO PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
     If nFilaGreen <> -1 Then
        If FEEmpresas.rows > nFilaGreen Then
        FEEmpresas.row = nFilaGreen
        FEEmpresas.BackColorRow (vbWhite)
        End If
     End If
     'FIN AGREGADO
    
    FEEmpresas.row = nPosFila
'    If Trim(FEEmpresas.TextMatrix(FEEmpresas.Row, 1)) <> "" Then
        FEEmpresas.col = 1
        If FEEmpresas.CellBackColor = vbGreen Then
            FEEmpresas.BackColorRow (vbWhite)
            FEEmpresas.TextMatrix(FEEmpresas.row, 0) = "."
            lsPersCodEmpre = ""
            lsPersCodVincu = ""
            nFilaGreen = -1 'ADD PTI1 17/08/2018
        Else
            'FEEmpresas.CellBackColor = vbGreen
            FEEmpresas.BackColorRow (vbGreen)
            FEEmpresas.TextMatrix(FEEmpresas.row, 0) = "R"
            nFilaGreen = FEEmpresas.row 'AGREGADO POR PTI1 17/08/2018
'            lsPersCodEmpre = FEEmpresas.TextMatrix(FEEmpresas.Row, 1)
            lsPersCodVincu = ""
        End If
        
            'ADD PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
    kg = 0
    Dim lnPersona As COMNPersona.NCOMPersona
    Set lnPersona = New COMNPersona.NCOMPersona
    
    If FEEmpresas.TextMatrix(FEEmpresas.row, 1) <> "" And Right(cboGruposEconomicos.Text, 4) <> "" Then
        Dim cper As String
        cper = Trim(FEEmpresas.TextMatrix(FEEmpresas.row, 1))
      
        Call lnPersona.ObtenerDatosSoloVinculadoxEmpresa(loDatosVinculados, loGestionVinculados, loGestion, cper, Right(cboGruposEconomicos.Text, 4))
      
        If loGestionVinculados.EOF Or loGestionVinculados.BOF Then
         Call ObtenerVinculados(FEEmpresas.TextMatrix(FEEmpresas.row, 1)) 'ADD PTI1 26092018
            Exit Sub
        End If

        Do Until loGestionVinculados.EOF
            kg = kg + 1
            ReDim Preserve sMatGestion(1 To 3, 1 To kg)

            sMatGestion(1, kg) = loGestionVinculados!cPersCod
            sMatGestion(2, kg) = loGestionVinculados!cRelacionGestion
            sMatGestion(3, kg) = loGestionVinculados!cPersCodIndependiente
            
            
            loGestionVinculados.MoveNext
        Loop
    
    Call ObtenerVinculados(FEEmpresas.TextMatrix(FEEmpresas.row, 1))
    Else
    LimpiaFlex FEEmpresas
    LimpiaFlex FEVinculados
    MsgBox "Por favor seleccione un grupo económico y presione mostrar nuevamente", vbInformation, "Detalle"
    End If
    'FIN AGREGADO PTI1 (17/08/2018) INFORME n°002-2017-AC-TI/CMACM
'    End If
'If FEEmpresas.TextMatrix(FEEmpresas.row, 1) <> "" Then 'comentado por pti1 17/08/2018
   ' Call ObtenerVinculados(FEEmpresas.TextMatrix(FEEmpresas.row, 1)) 'comentado por pti1 17/08/2018
'End If  'comentado por pti1 17/08/2018
End Sub

Private Sub FEVinculados_Click()
    If nCantVin > 0 Then
        lsPersCodVincu = Trim(FEVinculados.TextMatrix(FEVinculados.row, 2)) 'AGREGADO POR PTI1 17/08/2018
        'lsPersCodVincu = sMatVinculados(2, FEVinculados.row) 'COMENTADO POR PTI1 17/08/2018
    End If
End Sub

Private Sub Form_Load()
    Call CargarGrupos
End Sub

Private Sub Reporte19GruposEconomicos(ByVal pdFecha As Date)
 
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsNomHoja3  As String
    Dim lsNomHoja4  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsTipoGarantia As Integer
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lnContadorMatrix As Integer
    Dim lnPosY1 As Integer
    Dim lnPosY2 As Integer
    Dim lnPosYInicial19 As Integer
    Dim lnPosYVincula19 As Integer
    Dim lsArchivo As String
    Dim lnTipoSalto As Integer
    Dim oGrupEc As COMDPersona.DCOMGrupoE
    Set oGrupEc = New COMDPersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    Dim lnNumColumns As Integer
    Dim sMatrixVinculados() As String
    Dim ctipper As String
    Dim i As Integer
    Dim J As Integer
    Dim lsPersCodEmpresa As String
    Set rs = New ADODB.Recordset
    Set rs = oGrupEc.ListarDatosGrupoEconomicoxGrupoRepo19y20(1)
    Set oGrupEc = Nothing
    Dim nX As Integer
    Dim nY As Integer
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Reporte19"
    lsNomHoja1 = "Reporte19"
    lsArchivo1 = "\spooler\Reporte19" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja1 Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
    lnPosYInicial19 = 10
    i = 0
    Dim nEncontrado As Integer
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                If i > 0 Then
                    For J = 1 To i
                        If sMatrixVinculados(1, J) = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                    Next J
                End If
                If nEncontrado = 0 Then
                    i = i + 1
                    ReDim Preserve sMatrixVinculados(1 To 1, 1 To i)
                    sMatrixVinculados(1, i) = rs!cPersCod
                    xlHoja1.Cells(lnPosYInicial19, 1) = i
                    xlHoja1.Cells(lnPosYInicial19, 2) = rs!cPersEmpresa
                    xlHoja1.Cells(lnPosYInicial19, 3) = IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")
                    xlHoja1.Cells(lnPosYInicial19, 4) = rs!cPersCIIU
                    xlHoja1.Cells(lnPosYInicial19, 5) = rs!cDirecciEmpresa
                    xlHoja1.Cells(lnPosYInicial19, 6) = IIf(Trim(rs!P1DNI) = "-", "DNI", rs!P1DNI)
                    xlHoja1.Cells(lnPosYInicial19, 7) = IIf(Trim(rs!P1DNI) = "", "-", rs!P1DNI)
                    xlHoja1.Cells(lnPosYInicial19, 8) = IIf(Trim(rs!P1RUC) = "", "-", rs!P1RUC)
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 1), xlHoja1.Cells(lnPosYInicial19, 9)).Borders.LineStyle = 1
                    lnPosYInicial19 = lnPosYInicial19 + 1
                End If
                rs.MoveNext
            Wend
    End If
   
    
   
    lsNomHoja1 = "Reporte19A"
    For Each xlHoja1 In xlsLibro.Worksheets
    If xlHoja1.Name = lsNomHoja1 Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
    Next
   
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    
    lnPosYInicial19 = 8
    lnPosYVincula19 = 15
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
    If i >= 1 Then
    rs.MoveFirst
    End If
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                If i > 0 Then
                    For J = 1 To i
                        If lsPersCodEmpresa = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                    Next J
                End If
                If nEncontrado = 0 Then
                    If lnPosYVincula19 = 15 Then
                        lnPosYInicial19 = lnPosYInicial19 + 1
                    Else
                        lnPosYInicial19 = lnPosYVincula19 + 15
                        nY = lnPosYInicial19 + 4
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 1), xlHoja1.Cells(nY, 7)).Borders.LineStyle = 1
                        nY = lnPosYInicial19 + 6
                        xlHoja1.Cells(nY, 1) = xlHoja1.Cells(15, 1)
                        xlHoja1.Cells(nY, 2) = xlHoja1.Cells(15, 2)
                        xlHoja1.Cells(nY, 3) = xlHoja1.Cells(15, 3)
                        xlHoja1.Cells(nY, 4) = xlHoja1.Cells(15, 4)
                        xlHoja1.Cells(nY, 5) = xlHoja1.Cells(15, 5)
                        xlHoja1.Cells(nY, 6) = xlHoja1.Cells(15, 6)
                        xlHoja1.Cells(nY, 7) = xlHoja1.Cells(15, 7)
                        xlHoja1.Cells(nY, 8) = xlHoja1.Cells(15, 8)
                        xlHoja1.Cells(nY, 9) = xlHoja1.Cells(15, 9)
                        xlHoja1.Cells(nY, 10) = xlHoja1.Cells(15, 10)
                        xlHoja1.Cells(nY, 11) = xlHoja1.Cells(15, 11)
                        xlHoja1.Range(xlHoja1.Cells(nY, 1), xlHoja1.Cells(nY, 11)).Borders.LineStyle = 1
                    End If
             
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19 + 6, 3)).HorizontalAlignment = xlLeft
                   
                    lsPersCodEmpresa = rs!cPersCod
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(9, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(9, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = rs!cPersEmpresa
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(10, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(10, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = rs!cPersCodSbs1
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(11, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(11, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = IIf(rs!P1RUC = "", "-", rs!P1RUC)
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(12, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(12, 2)
                    xlHoja1.Cells(lnPosYInicial19, 3) = rs!cDirecciEmpresa
                    lnPosYInicial19 = lnPosYInicial19 + 1
                    xlHoja1.Range(xlHoja1.Cells(lnPosYInicial19, 3), xlHoja1.Cells(lnPosYInicial19, 7)).Merge True
                    xlHoja1.Cells(lnPosYInicial19, 1) = xlHoja1.Cells(13, 1)
                    xlHoja1.Cells(lnPosYInicial19, 2) = xlHoja1.Cells(13, 2)
                    If rs!nRepresentanteLegal = 1 Then
                        xlHoja1.Cells(lnPosYInicial19, 3) = rs!cPersVinculado
                    End If
                End If
                    If nEncontrado = 0 And lnPosYVincula19 > 15 Then
                        lnPosYVincula19 = lnPosYVincula19 + 22
                    Else
                        lnPosYVincula19 = lnPosYVincula19 + 1
                    End If
                    xlHoja1.Cells(lnPosYVincula19, 1) = i
                    xlHoja1.Cells(lnPosYVincula19, 2) = rs!cPersVinculado
                    xlHoja1.Cells(lnPosYVincula19, 3) = rs!cPersCodSbs2
                    xlHoja1.Cells(lnPosYVincula19, 4) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                    xlHoja1.Cells(lnPosYVincula19, 5) = IIf(Trim(rs!P2DNI) = "-", "DNI", rs!P2DNI)
                    xlHoja1.Cells(lnPosYVincula19, 6) = IIf(Trim(rs!P2DNI) = "", "-", rs!P2DNI)
                    xlHoja1.Cells(lnPosYVincula19, 7) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                    xlHoja1.Cells(lnPosYVincula19, 8) = rs!cDirecciVinculado
                    xlHoja1.Cells(lnPosYVincula19, 10) = IIf(rs!nPorcenOtro > 0, CStr(rs!nPorcenOtro) & "%", "")
                    xlHoja1.Cells(lnPosYVincula19, 10) = rs!nCargo
                    xlHoja1.Cells(lnPosYVincula19, 11) = rs!nCargoOtro
                    xlHoja1.Range(xlHoja1.Cells(lnPosYVincula19, 1), xlHoja1.Cells(lnPosYVincula19, 11)).Borders.LineStyle = 1
                    
                rs.MoveNext
            Wend
    End If
    
    rs.Close
    Set rs = Nothing
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub

Private Sub Reporte20RiesgoUnico(ByVal pdFecha As Date)
 
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet

    Dim fs As Scripting.FileSystemObject
    Dim lbExisteHoja As Boolean
    Dim lsArchivo1 As String
    Dim lsNomHoja1  As String
    Dim lsNomHoja2  As String
    Dim lsNomHoja3  As String
    Dim lsNomHoja4  As String
    Dim lsNombreAgencia As String
    Dim lsCodAgencia As String
    Dim lsTipoGarantia As Integer
    Dim lsMes As String
    Dim lnContador As Integer
    Dim lnContadorMatrix As Integer
    Dim lnPosY1 As Integer
    Dim lnPosY2 As Integer
    Dim lnPosYInicial20 As Integer
    Dim lnPosYVincula19 As Integer
    Dim lsArchivo As String
    Dim lnTipoSalto As Integer
    Dim oGrupEc As COMDPersona.DCOMGrupoE
    Set oGrupEc = New COMDPersona.DCOMGrupoE
    Dim rs As ADODB.Recordset
    Dim lnNumColumns As Integer
    Dim sMatrixVinculados() As String
    Dim sMatrixPropiedadDirecIn() As String
    Dim sMatrixEmpresas() As String
    Dim ctipper As String
    Dim i, ix, ipdi As Integer
    Dim J, K, h, L As Integer
    Dim m, n As Integer
    Dim lsPersCodEmpresa As String
    Dim nPosEmpresa As Integer
    Dim lsPropiedadDirecta, lsPropiedadIndirecta As String
    Set rs = New ADODB.Recordset
    Set rs = oGrupEc.ListarDatosGrupoEconomicoxRURepo20(1, pdFecha)
    Set oGrupEc = Nothing
    Dim nX As Integer
    Dim nY As Integer
    Dim nTJ As Integer
    Dim oPerPDI As COMDPersona.DCOMGrupoE
    Dim RsProDI As ADODB.Recordset
    Set RsProDI = New ADODB.Recordset
    Set oPerPDI = New COMDPersona.DCOMGrupoE
    Set RsProDI = oPerPDI.ObtenerDatosPropiedadDirectaEIndirecta
    If Not (RsProDI.EOF And RsProDI.BOF) Then
            While Not RsProDI.EOF
                ipdi = ipdi + 1
                ReDim Preserve sMatrixPropiedadDirecIn(1 To 4, 1 To ipdi)
                sMatrixPropiedadDirecIn(1, ipdi) = RsProDI!nPorcenOtro
                sMatrixPropiedadDirecIn(2, ipdi) = RsProDI!cPersCod
                sMatrixPropiedadDirecIn(3, ipdi) = RsProDI!cPersCodOtro
                sMatrixPropiedadDirecIn(4, ipdi) = RsProDI!nTipo
                RsProDI.MoveNext
            Wend
    End If
    
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "Reporte20"
    lsNomHoja1 = "Reporte20"
    lsArchivo1 = "\spooler\Reporte20" & gsCodUser & "_" & Format(pdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja1 Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
    lnPosYInicial20 = 8
    i = 0
    Dim nEncontrado As Integer
    Dim nEncontradoVin As Integer
    nTJ = 0
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
            If rs!nSaldo > 0 Then
                nEncontrado = 0
                nEncontradoVin = 0
                If i > 0 Then
                    For J = 1 To i
                        If sMatrixVinculados(1, J) = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                        If sMatrixVinculados(1, J) = rs!cPersCodOtro Then
                            nEncontradoVin = 1
                        End If
                    Next J
                End If
              
                If nEncontrado = 0 Or nEncontradoVin = 0 Then
                    If nEncontrado = 0 Then
                        i = i + 1
                        ix = ix + 1
                        If nTJ = 1 Then
                            lnPosYInicial20 = lnPosYInicial20 + 2
                        Else
                            lnPosYInicial20 = lnPosYInicial20 + 1
                        End If
                        nTJ = 0
                        ReDim Preserve sMatrixVinculados(1 To 2, 1 To i)
                        
                        sMatrixVinculados(1, i) = rs!cPersCod
                        sMatrixVinculados(2, i) = ix
                        xlHoja1.Cells(lnPosYInicial20, 1) = ix & ".-"
                        xlHoja1.Cells(lnPosYInicial20, 2) = "INFORMACION DEL CLIENTE"
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Nombre, razon o denominación Social"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersEmpresa
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Codigo SBS"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs1
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Tipo de Persona"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Documento de Identidad y Número"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P1DNI) = "", "", "DNI - ") & IIf(Trim(rs!P1DNI) = "", "", rs!P1DNI)
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "RUC"
                        xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P1RUC) = "", "-", rs!P1RUC)
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Direccción"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cDirecciEmpresa
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        xlHoja1.Cells(lnPosYInicial20, 2) = "Representante Legal"
                        xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersNombreReprLegal
                        xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                        lnPosYInicial20 = lnPosYInicial20 + 1
                        End If
                         
                        If (IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")) = "NAT" And nEncontradoVin = 0 And rs!cPersCodOtro <> rs!cPersCod Then
                            i = i + 1
                            ix = ix + 1
                            ReDim Preserve sMatrixVinculados(1 To 2, 1 To i)
                            sMatrixVinculados(1, i) = rs!cPersCodOtro
                            sMatrixVinculados(2, i) = ix
                            xlHoja1.Cells(lnPosYInicial20, 1) = ix & ".-"
                            xlHoja1.Cells(lnPosYInicial20, 2) = "INFORMACION DEL CLIENTE"
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "Nombre, razon o denominación Social"
                            xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersVinculado
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "Codigo SBS"
                            xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs2
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "Tipo de Persona"
                            xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "Documento de Identidad y Número"
                            xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P2DNI) = "", "", "DNI - ") & IIf(Trim(rs!P2DNI) = "", "", rs!P2DNI)
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "RUC"
                            xlHoja1.Cells(lnPosYInicial20, 3) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "Direccción"
                            xlHoja1.Cells(lnPosYInicial20, 3) = rs!cDirecciVinculado
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            xlHoja1.Cells(lnPosYInicial20, 2) = "Representante Legal"
                            xlHoja1.Cells(lnPosYInicial20, 3) = ""
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 3)).Borders.LineStyle = 1
                            lnPosYInicial20 = lnPosYInicial20 + 1
                        ElseIf (IIf(Mid(rs!P1RUC, 1, 1) = "2", "JUR", "NAT")) = "JUR" And nEncontradoVin = 0 Then
                            i = i + 1
                            If lsPersCodEmpresa <> rs!cPersCod Then
                                xlHoja1.Cells(lnPosYInicial20, 2) = "2. ACCIONISTAS, DIRECTORES, GERENTES, PRINCIPALES FUNCIONARIOS Y ASESORES"
                                lnPosYInicial20 = lnPosYInicial20 + 2
                                xlHoja1.Cells(lnPosYInicial20, 2) = "Nombre"
                                xlHoja1.Cells(lnPosYInicial20, 3) = "Cod Sbs"
                                xlHoja1.Cells(lnPosYInicial20, 4) = "TIP PER"
                                xlHoja1.Cells(lnPosYInicial20, 5) = "TIP DOC"
                                xlHoja1.Cells(lnPosYInicial20, 6) = "NRO DOC"
                                xlHoja1.Cells(lnPosYInicial20, 7) = "RUC"
                                xlHoja1.Cells(lnPosYInicial20, 8) = "Residencia"
                                xlHoja1.Cells(lnPosYInicial20, 9) = "ACC"
                                xlHoja1.Cells(lnPosYInicial20, 10) = "CAR"
                                xlHoja1.Cells(lnPosYInicial20, 11) = "OTR CAR"
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 11)).Borders.LineStyle = 1
                            End If
                                lsPersCodEmpresa = rs!cPersCod
                                i = i + 1
                                lnPosYInicial20 = lnPosYInicial20 + 1
                                ReDim Preserve sMatrixVinculados(1 To 2, 1 To i)
                                xlHoja1.Cells(lnPosYInicial20, 2) = rs!cPersVinculado
                                xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs2
                                xlHoja1.Cells(lnPosYInicial20, 4) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                                xlHoja1.Cells(lnPosYInicial20, 5) = IIf(Trim(rs!P2DNI) = "", "", "DNI")
                                xlHoja1.Cells(lnPosYInicial20, 6) = IIf(Trim(rs!P2DNI) = "", "", rs!P2DNI)
                                xlHoja1.Cells(lnPosYInicial20, 7) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                                xlHoja1.Cells(lnPosYInicial20, 8) = rs!cDirecciVinculado
                                xlHoja1.Cells(lnPosYInicial20, 9) = IIf(rs!nPorcenOtro > 0, CStr(rs!nPorcenOtro) & "%", "")
                                xlHoja1.Cells(lnPosYInicial20, 10) = rs!nCargo
                                xlHoja1.Cells(lnPosYInicial20, 11) = rs!nCargoOtro
                                xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 2), xlHoja1.Cells(lnPosYInicial20, 11)).Borders.LineStyle = 1
                                nTJ = 1
                       End If
                    End If
                    End If
                rs.MoveNext
            Wend
    End If
    
    lsNomHoja1 = "Reporte20A"
    For Each xlHoja1 In xlsLibro.Worksheets
    If xlHoja1.Name = lsNomHoja1 Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
    Next
 
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja1
    End If
    If i >= 1 Then
        rs.MoveFirst
    End If
    m = i
    m = 0
    h = 0
    xlHoja1.Cells(7, 3) = Format(pdFecha, "DD.MM.YYYY")
     If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
                nEncontrado = 0
                nEncontradoVin = 0
                If m > 0 Then
                    For J = 1 To m
                        If sMatrixVinculados(1, J) = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                        If sMatrixVinculados(1, J) = rs!cPersCodOtro Then
                            nEncontradoVin = 1
                        End If
                    Next J
                End If
              
                If nEncontrado = 0 Or nEncontradoVin = 0 Then
   
                            If lsPersCodEmpresa <> rs!cPersCod Then
                                 h = h + 1
                                 ReDim Preserve sMatrixEmpresas(1 To 2, 1 To h)
                                  sMatrixEmpresas(1, h) = rs!cPersCod
                                  sMatrixEmpresas(2, h) = h
                            End If
                            lsPersCodEmpresa = rs!cPersCod
                End If
                rs.MoveNext
            Wend
    End If
 lsPersCodEmpresa = ""

    If i >= 1 Then
        rs.MoveFirst
    End If
    lnPosYInicial20 = 10
    i = 0
    ix = 0
    If Not (rs.EOF And rs.BOF) Then
            lnContadorMatrix = 1
            While Not rs.EOF
'            If rs!nSaldo > 0 Then
                nEncontrado = 0
                nEncontradoVin = 0
                If i > 0 Then
                    For J = 1 To i
                        If sMatrixVinculados(1, J) = rs!cPersCod Then
                            nEncontrado = 1
                        End If
                        If sMatrixVinculados(1, J) = rs!cPersCodOtro Then
                            nEncontradoVin = 1
                        End If
                    Next J
                End If
              
                If nEncontrado = 0 Or nEncontradoVin = 0 Then
                        If rs!cPersCodOtro <> rs!cPersCod Then
                            lnPosYInicial20 = lnPosYInicial20 + 1
                            i = i + 1
                            If lsPersCodEmpresa <> rs!cPersCod Then
                                ix = ix + 1
                            End If
                            If lsPersCodEmpresa <> rs!cPersCod Then
                                xlHoja1.Cells(lnPosYInicial20, 1) = ix
                                xlHoja1.Cells(lnPosYInicial20, 2) = rs!cPersCodSbs1
                            End If
                            xlHoja1.Cells(lnPosYInicial20, 3) = rs!cPersCodSbs2
                            xlHoja1.Cells(lnPosYInicial20, 4) = rs!cPersVinculado
                            xlHoja1.Cells(lnPosYInicial20, 5) = "CIIU"
                            xlHoja1.Cells(lnPosYInicial20, 6) = rs!cDirecciVinculado
                            xlHoja1.Cells(lnPosYInicial20, 7) = IIf(Mid(rs!P2RUC, 1, 1) = "2", "JUR", "NAT")
                            xlHoja1.Cells(lnPosYInicial20, 8) = IIf(Trim(rs!P2DNI) = "", "", "DNI")
                            xlHoja1.Cells(lnPosYInicial20, 9) = IIf(Trim(rs!P2DNI) = "", "", rs!P2DNI)
                            xlHoja1.Cells(lnPosYInicial20, 10) = IIf(Trim(rs!P2RUC) = "", "-", rs!P2RUC)
                            lsPropiedadDirecta = ""
                            lsPropiedadIndirecta = ""
                            For K = 1 To ipdi
                                If sMatrixPropiedadDirecIn(3, K) = rs!cPersCodOtro Then
                                    For L = 1 To h
                                        If sMatrixEmpresas(1, L) = sMatrixPropiedadDirecIn(2, K) Then
                                            nPosEmpresa = sMatrixEmpresas(2, L)
                                        End If
                                    Next L
                                    If sMatrixPropiedadDirecIn(4, K) = 1 Then
                                    lsPropiedadDirecta = lsPropiedadDirecta & " " & sMatrixPropiedadDirecIn(1, K) & "% (" & nPosEmpresa & ")"
                                    Else
                                    lsPropiedadIndirecta = lsPropiedadIndirecta & " " & sMatrixPropiedadDirecIn(1, K) & "% (" & nPosEmpresa & ")"
                                    End If
                                End If
                            Next K
                            
                            xlHoja1.Cells(lnPosYInicial20, 11) = lsPropiedadDirecta
                            xlHoja1.Cells(lnPosYInicial20, 12) = lsPropiedadIndirecta
                            
                            xlHoja1.Cells(lnPosYInicial20, 13) = rs!cGestion
                            xlHoja1.Range(xlHoja1.Cells(lnPosYInicial20, 1), xlHoja1.Cells(lnPosYInicial20, 13)).Borders.LineStyle = 1
                            lsPersCodEmpresa = rs!cPersCod
                       End If
'                    End If
                rs.MoveNext
                End If
            Wend
    End If
    
    rs.Close
    Set rs = Nothing
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing

End Sub

'JAME20140303 ************************
'Private Sub txtBuscaEmpresa_Change()
'  Call cmdMostrar_Click
'End Sub

Private Sub txtBuscaEmpresa_GotFocus()
    lbBuscaTexto = True
End Sub

Private Sub txtBuscaEmpresa_LostFocus()
    lbBuscaTexto = False
End Sub
'JAME FIN ***************************S
'EJVG20160611 ***
Private Sub txtBuscaEmpresa_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        If Len(Trim(txtBuscaEmpresa.Text)) = 0 Then
            MsgBox "Falta ingresar el nombre del Cliente", vbInformation, "Aviso"
            Exit Sub
        End If
        cmdMostrar_Click
    End If
End Sub
'END EJVG *******
