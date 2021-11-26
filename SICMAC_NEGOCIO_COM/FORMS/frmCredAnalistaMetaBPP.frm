VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCredAnalistaMetaBPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Meta Analista"
   ClientHeight    =   7710
   ClientLeft      =   1065
   ClientTop       =   2115
   ClientWidth     =   13170
   Icon            =   "frmCredAnalistaMetaBPP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13170
   Begin VB.Frame Frame4 
      Caption         =   "Metas de Analista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5625
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   13095
      Begin VB.CheckBox chkAnalista 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcCordinador 
         Height          =   315
         Left            =   8880
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAnalista 
         Height          =   4815
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   8493
         _Version        =   393216
         Cols            =   14
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   14
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11400
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Analista"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton cmdConsolidar 
         Caption         =   "Consolidar Bono Mes Anterior"
         Height          =   615
         Left            =   9120
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdAnterior 
         Caption         =   "Utilizar Mes Anterior"
         Height          =   375
         Left            =   9120
         TabIndex        =   16
         Top             =   1170
         Width           =   1575
      End
      Begin VB.ComboBox cmbAnio 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCredAnalistaMetaBPP.frx":030A
         Left            =   7200
         List            =   "frmCredAnalistaMetaBPP.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cmbMes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCredAnalistaMetaBPP.frx":030E
         Left            =   5640
         List            =   "frmCredAnalistaMetaBPP.frx":033C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dcTipoCartera 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   315
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTipoClasificacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAgencia 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
      End
      Begin VB.Label Label7 
         Caption         =   "Mes Meta:"
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cartera:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Clasificación:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCredAnalistaMetaBPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------REALIZADO X JACA 20110309--------------
Option Explicit
Dim objCOMDCredito As COMDCredito.DCOMBPPR
Dim objCOMNCredito As COMNCredito.NCOMBPPR
Dim lnNroClieMin As Integer
Dim lnNroOpeMin As Integer
Dim lnSaldoMin As Long
Dim lsFechaMeta As String
Dim bExisteMetas As Boolean
'Dim matris() As String

Private Sub cmbAnio_Click()
    If (cmbMes.ListIndex <> -1 And cmbAnio.ListIndex <> -1) And (cmbMes.ListIndex <> 0 And cmbAnio.ListIndex <> 0) And Me.dcAgencia.BoundText <> 0 Then
        'If lsFechaMeta <> "" Then
            'If val(cmbMes.ItemData(cmbMes.ListIndex)) < val(Mid(lsFechaMeta, 5, 2)) And val(cmbAnio.ItemData(cmbAnio.ListIndex)) <= val(Mid(lsFechaMeta, 1, 4)) Then
            If val(cmbMes.ItemData(cmbMes.ListIndex)) < val(Mid(gdFecSis, 4, 2)) And val(cmbAnio.ItemData(cmbAnio.ListIndex)) <= val(Mid(gdFecSis, 7, 4)) Then
'
                'desHabilitarControles , , , , 1, 1, 1, 1
                desHabilitarControles , , , 1, 1, 1, 1, 1, 1, 1
                'CargarListaAnalistaMetasAnteriores Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2)
                CargarListaAnalistaMetas Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2), 1
            ElseIf cmbMes.ItemData(cmbMes.ListIndex) = val(Mid(gdFecSis, 4, 2)) And cmbAnio.ItemData(cmbAnio.ListIndex) = val(Mid(gdFecSis, 7, 4)) Then
                'CargaListaAnalista
                CargarListaAnalistaMetas Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2), 1
                
'                HabilitarControles , , , , , 1
'                desHabilitarControles , , , , 1, , 1, 1
                HabilitarControles , , , 1, , 1
                desHabilitarControles , , , , 1, , 1, 1
            Else
                MsgBox "Aun Hay Informacion Consolidada de " + cmbMes.Text + " del " + cmbAnio.Text, vbInformation
                 desHabilitarControles , , , 1, 1, 1, 1, 1, 1, 1
'                CargaListaAnalista
'                cmdEditar.Enabled = False
'                cmdRegistrar.Enabled = True
'                cmdCancelar.Enabled = False
'                grdAnalista.Enabled = True
            End If
        
        'End If
    End If
End Sub



Private Sub cmbMes_Click()
    If (cmbMes.ListIndex <> -1 And cmbAnio.ListIndex <> -1) And (cmbMes.ListIndex <> 0 And cmbAnio.ListIndex <> 0) And Me.dcAgencia.BoundText <> 0 Then
        
            'Meses Anteriores al mes Actual
            If val(cmbMes.ItemData(cmbMes.ListIndex)) < val(Mid(gdFecSis, 4, 2)) And val(cmbAnio.ItemData(cmbAnio.ListIndex)) <= val(Mid(gdFecSis, 7, 4)) Then

                desHabilitarControles , , , 1, 1, 1, 1, 1, 1, 1
                Call LimpiaFlex(grdAnalista)
                 CargarListaAnalistaMetas Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2)
            
            'Mes Actual
            ElseIf cmbMes.ItemData(cmbMes.ListIndex) = val(Mid(gdFecSis, 4, 2)) And cmbAnio.ItemData(cmbAnio.ListIndex) = val(Mid(gdFecSis, 7, 4)) Then
                CargaListaAnalista
                CargarListaAnalistaMetas Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2), 1
              
                HabilitarControles , , , 1, , 1
                desHabilitarControles , , , , 1, , 1, 1
            
            'Meses siguientes al mes Actual
            Else
                MsgBox "Aun No Existe Informacion Consolidada de " + Right(cmbMes.Text, Len(cmbMes.Text) - 3) + " del " + cmbAnio.Text, vbInformation
                desHabilitarControles , , , 1, 1, 1, 1, 1, 1, 1

            End If
        
    End If
End Sub

Private Sub cmdAnterior_Click()
    If Me.dcAgencia.BoundText <> "0" Then
        CargarListaAnalistaMetas "0", 1
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    
    HabilitarControles , , , , , 1
    desHabilitarControles , , , , 1, , 1, 1, 1
    'Unload Me
End Sub

Private Sub cmdConsolidar_Click()
        Dim rsBono As New Recordset
        Dim clsTC As COMDConstSistema.NCOMTipoCambio
        Dim sMovNro As String
        Dim clsMov As COMNContabilidad.NCOMContFunciones
        
        Dim nTC As Double
        Dim dFecha As Date
        Dim dFechaMesAnt As Date
        Dim dFechaTC As Date
        Dim i As Integer
      
       
       If Me.dcAgencia.BoundText = "0" Then
        Exit Sub
       End If
       Set clsTC = New COMDConstSistema.NCOMTipoCambio
       
      
       Me.MousePointer = vbHourglass
      
       Set objCOMNCredito = New COMNCredito.NCOMBPPR
       
       
       If MsgBox("Esto Puede Tomar Varios Minutos", vbOKCancel, "Aviso") = vbOk Then
            dFecha = DateAdd("d", -Day(gdFecSis), gdFecSis)
            dFechaTC = DateAdd("d", 1, dFecha)
            nTC = clsTC.EmiteTipoCambio(dFechaTC, TCFijoDia)
            
            Set clsMov = New COMNContabilidad.NCOMContFunciones
            sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                 
            
            Set rsBono.DataSource = objCOMNCredito.obtenerBonoAnalistas(Me.dcAgencia.BoundText, nTC, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, dFecha)
            
            If Not rsBono.EOF Then
                For i = 0 To rsBono.RecordCount - 1
                    objCOMNCredito.guardarBonoAnalistas rsBono!cAgeCodAct, rsBono!cCodAna, rsBono!iTipoCarteraId, rsBono!iTipoClasificacionId, _
                                                       rsBono!NroCred, rsBono!NroClie, rsBono!Saldo, _
                                                        rsBono!CredTrans, rsBono!ClieTrans, rsBono!SaldoTrans, _
                                                        rsBono!nMetaClie, rsBono!nMetaOpe, rsBono!nMetaSaldo, rsBono!Cord, _
                                                        rsBono!TOTALCRED, rsBono!TOTALCLIE, rsBono!TOTALSALDO, _
                                                        rsBono!TOTALCLIENUE, rsBono!TOTALCLIERECU, _
                                                        rsBono!TOTALOPE, rsBono!TOTALMORA, rsBono!TOTALTASA, _
                                                        rsBono!BPPClieNue, rsBono!BPPClieRecu, rsBono!BPPSaldo, _
                                                        rsBono!BPPOpe, rsBono!BPPMora, rsBono!BPPTasa, rsBono!BONIFICACION, _
                                                        dFecha, sMovNro
                    rsBono.MoveNext
                Next i
                cmdConsolidar.Enabled = False
                MsgBox "Se Ha Guardado la Bonificacion del Mes Anterior", vbInformation, "Aviso"
            Else
                MsgBox "No Existe Bonificacion del Mes Anterior", vbInformation, "Aviso"
            End If
        End If
    
    Me.MousePointer = vbDefault
            
End Sub

Private Sub CmdEditar_Click()
        
    HabilitarControles , , , , 1, , 1, 1, 1
    desHabilitarControles , , , , , 1
    'chkAnalista.value = 1
End Sub

Private Sub cmdRegistrar_Click()
        
        
        If Me.dcAgencia.BoundText = 0 Then
            MsgBox ("Debe Seleccionar una Agencia")
            Exit Sub
        End If
        
        If Me.dcTipoCartera.BoundText = 0 Then
            MsgBox ("Debe Seleccionar un Tipo Cartera")
            Exit Sub
        End If
        
        If Me.dcTipoClasificacion.BoundText = 0 Then
            MsgBox ("Debe Seleccionar un Tipo de Clasificacion de Cartera")
            Exit Sub
        End If
        
        
        Dim i As Integer
        Dim nNroAna As Integer
        nNroAna = 0
        
        'verificar si selecciono un analista y si todos los datos de los analista esta completo
        For i = 1 To grdAnalista.Rows - 1
            If grdAnalista.TextMatrix(i, 1) = Chr(254) Then
                                  
                  If grdAnalista.TextMatrix(i, 2) = "" Or grdAnalista.TextMatrix(i, 3) = "" Or grdAnalista.TextMatrix(i, 10) = "" Or grdAnalista.TextMatrix(i, 11) = "" Or grdAnalista.TextMatrix(i, 12) = "" Or grdAnalista.TextMatrix(i, 13) = "" Then
                      MsgBox ("Debe Ingresar Todos los Datos del Analista:'" + grdAnalista.TextMatrix(i, 2) + "' en la Fila Nro '" + CStr(i) + "'")
                      Exit Sub
                  End If
                  nNroAna = nNroAna + 1
            End If
        Next
        
        Dim objCOMNCredito As COMNCredito.NCOMBPPR
        Dim sMovNro As String
        Dim clsMov As COMNContabilidad.NCOMContFunciones
        
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        Set clsMov = New COMNContabilidad.NCOMContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        If nNroAna = 0 Then
                MsgBox ("Debe Seleccionar al menos un Analista")
                Exit Sub
        Else
            If MsgBox("¿Está seguro de Registrar los Datos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                    
                    objCOMNCredito.actualizarEstadoAnalistaMeta Me.dcAgencia.BoundText, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2), sMovNro
                    For i = 1 To grdAnalista.Rows - 1
                        If grdAnalista.TextMatrix(i, 1) = Chr(254) Then
                           objCOMNCredito.InsertarAnalistaMetaBPP Me.dcAgencia.BoundText, grdAnalista.TextMatrix(i, 2), grdAnalista.TextMatrix(i, 3), Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, grdAnalista.TextMatrix(i, 4), grdAnalista.TextMatrix(i, 5), grdAnalista.TextMatrix(i, 7), grdAnalista.TextMatrix(i, 6), grdAnalista.TextMatrix(i, 8), grdAnalista.TextMatrix(i, 9), grdAnalista.TextMatrix(i, 10), grdAnalista.TextMatrix(i, 11), grdAnalista.TextMatrix(i, 12), grdAnalista.TextMatrix(i, 13), Me.cmbAnio.List(cmbAnio.ListIndex) + Mid(Me.cmbMes.Text, 1, 2), sMovNro
                        End If
                    Next
                    
                    MsgBox ("Se Guardaron Todos los Datos con Exito!")
                    LimpiarControles
            End If
        End If
        
        
End Sub

Private Sub dcAgencia_Change()
    lsFechaMeta = ""
    If dcAgencia.BoundText <> "0" Then
        If Me.dcTipoClasificacion.BoundText <> "" And Me.dcTipoClasificacion.BoundText <> "0" Then
            CargaListaAnalista
            'CargarCordinadoresAgencia
            CargarListaAnalistaMetas Mid(gdFecSis, 7, 4) + Mid(gdFecSis, 4, 2), 1
            
            Dim i As Integer
            cmbMes.ListIndex = val(Mid(gdFecSis, 4, 2))
            i = getIndiceAnio(Mid(gdFecSis, 7, 4))
            cmbAnio.ListIndex = i
            If lsFechaMeta = "" Then
                'Me.cmbAnio.ListIndex = 1
                'Me.cmbMes.ListIndex = 1
                chkAnalista.value = 1
                HabilitarControles , , , , , 1
'            Else
'                Dim i As Integer
'                cmbMes.ListIndex = val((Mid(lsFechaMeta, 5, 2)))
'                i = getIndiceAnio(Mid(lsFechaMeta, 1, 4))
'                cmbAnio.ListIndex = i
            End If
        Else
            MsgBox "Seleccione una Clasificacion de Tipo de Cartera"
        End If
        'Me.dcTipoCartera.BoundText = "0"
        'Me.dcTipoClasificacion.BoundText = "0"
        'Me.cmbAnio.ListIndex = 0
        'Me.cmbMes.ListIndex = 0
        'cmdAnterior.Enabled = False
        
        HabilitarControles , 1, 1, 1
    End If
End Sub
Private Sub CargarListaAnalistaMetas(ByVal sFechaMeta As String, Optional pnOp As Integer = 0)
        Dim rsMetas As New Recordset
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        
        
        Set rsMetas.DataSource = objCOMNCredito.getListaAnalistaMetas(Me.dcAgencia.BoundText, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, sFechaMeta)
        
        
        Dim i As Integer
        Dim J As Integer
        
        If rsMetas.RecordCount > 0 Then
           ' CargaListaAnalista
            
            chkAnalista.value = 0
            desHabilitarControles , , , , 1, , , , 1
            lsFechaMeta = rsMetas!cFechaMeta
'            cmbMes.ListIndex = val((Mid(rsMetas!cFechaMeta, 5, 2)))
'            i = getIndiceAnio(Mid(rsMetas!cFechaMeta, 1, 4))
'            cmbAnio.ListIndex = i
            If pnOp = 0 Then
                grdAnalista.Rows = rsMetas.RecordCount + 1
            End If
            
            If pnOp = 1 Then 'del mes actual
                For i = 1 To grdAnalista.Rows - 1
                    rsMetas.MoveFirst
                    For J = 0 To rsMetas.RecordCount - 1
                        If grdAnalista.TextMatrix(i, 2) = rsMetas!cCodAna Then
                           grdAnalista.TextMatrix(i, 1) = Chr(254)
                           
                           'grdAnalista.TextMatrix(i, 10) = rsMetas!cCodCord + Space(20) + rsMetas!UserCord
                           grdAnalista.TextMatrix(i, 10) = rsMetas!cCodCord
                           grdAnalista.TextMatrix(i, 11) = rsMetas!nMetaClie
                           grdAnalista.TextMatrix(i, 12) = rsMetas!nMetaOpe
                           grdAnalista.TextMatrix(i, 13) = Format(rsMetas!nMetaSaldo, "##,#00")
                           
                           Exit For
                        Else
                           grdAnalista.TextMatrix(i, 1) = Chr(168)
                                           
                        End If
                         rsMetas.MoveNext
                    Next
                   
                Next
                HabilitarControles , , , , , 1
            End If
            If pnOp = 0 Then 'de mese anteriores
                    grdAnalista.Redraw = False
                    For i = 0 To rsMetas.RecordCount - 1
                        
                           grdAnalista.TextMatrix(i + 1, 1) = Chr(254)
                           'convierte  el caracter Chr(254) a check=true
                            grdAnalista.Row = i + 1
                            grdAnalista.Col = 1
                            grdAnalista.CellFontName = "Wingdings"
                            grdAnalista.CellFontSize = 13
                        
                            grdAnalista.TextMatrix(i + 1, 2) = rsMetas!cCodAna
                            grdAnalista.TextMatrix(i + 1, 3) = rsMetas!cUser
                            grdAnalista.TextMatrix(i + 1, 4) = rsMetas!nCred
                            grdAnalista.TextMatrix(i + 1, 5) = rsMetas!nClie
                            grdAnalista.TextMatrix(i + 1, 6) = Format(rsMetas!nSaldo, "##,#00.00")
                            grdAnalista.TextMatrix(i + 1, 7) = rsMetas!nOpe
                            grdAnalista.TextMatrix(i + 1, 8) = Round(rsMetas!nMora, 2)
                            grdAnalista.TextMatrix(i + 1, 9) = Round(rsMetas!nTasaPond, 2)
                        
                           grdAnalista.TextMatrix(i + 1, 10) = rsMetas!cCodCord
                           grdAnalista.TextMatrix(i + 1, 11) = rsMetas!nMetaClie
                           grdAnalista.TextMatrix(i + 1, 12) = rsMetas!nMetaOpe
                           grdAnalista.TextMatrix(i + 1, 13) = Format(rsMetas!nMetaSaldo, "##,#00")
                           '****************************
                         rsMetas.MoveNext
                    Next
                grdAnalista.Redraw = True
                 desHabilitarControles , , , , , 1
            End If
            
'            HabilitarControles , , , , , 1
            desHabilitarControles , , , , , , , , 1
        ElseIf pnOp = 0 Then
            MsgBox "No Existen Registro del Mes Anterior"
'            cmdRegistrar.Enabled = True
'            cmdCancelar.Enabled = True
'        Else
'             desHabilitarControles , , , 1
        End If
End Sub

Private Sub CargarTipoCartera()
    Dim rsTipoCartera As New ADODB.Recordset
    Set objCOMDCredito = New COMDCredito.DCOMBPPR
    Set rsTipoCartera.DataSource = objCOMDCredito.CargarTipoCartera
    dcTipoCartera.BoundColumn = "iTipoCarteraId"
    dcTipoCartera.DataField = "iTipoCarteraId"
    Set dcTipoCartera.RowSource = rsTipoCartera
    dcTipoCartera.ListField = "vTipoCartera"
    dcTipoCartera.BoundText = 0
End Sub

Private Sub dcCordinador_Change()
'    If dcCordinador.Visible = False Then
'        dcCordinador.Visible = True
'        dcCordinador.SetFocus
'    Else
'        grdAnalista.TextMatrix(grdAnalista.Row, 10) = dcCordinador.Text
'        dcCordinador.Visible = False
'        grdAnalista.Col = grdAnalista.Col + 1
'        grdAnalista.CellBackColor = &H80000018
'        grdAnalista.SetFocus
'    End If
End Sub

Private Sub dcCordinador_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            grdAnalista.TextMatrix(grdAnalista.Row, 10) = dcCordinador.Text
            dcCordinador.Visible = False
            grdAnalista.Col = grdAnalista.Col + 1
            grdAnalista.CellBackColor = &H80000018
            grdAnalista.SetFocus
        End If
End Sub

Private Sub dcTipoCartera_Change()

    
    If dcTipoCartera.BoundText <> 0 Then

        limpiarMetas
        
        CargarTipoClasificacion
        Me.dcTipoClasificacion.BoundText = "0"
        Me.cmbAnio.ListIndex = 0
        Me.cmbMes.ListIndex = 0
        
        desHabilitarControles , , , , 1, 1, 1, 1, 1
        'CheckAnalista
    End If
End Sub

Private Sub CargarTipoClasificacion()
    Dim rsTipoClasificacion As New ADODB.Recordset
    Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsTipoClasificacion.DataSource = objCOMNCredito.getCargarTipoClasificacion(dcTipoCartera.BoundText)
    dcTipoClasificacion.BoundColumn = "iTipoClasificacionId"
    dcTipoClasificacion.DataField = "iTipoClasificacionId"
    Set dcTipoClasificacion.RowSource = rsTipoClasificacion
    dcTipoClasificacion.ListField = "vTipoClasificacion"
    dcTipoClasificacion.BoundText = 0
End Sub

Private Sub dcTipoCartera_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Then
'            dcTipoCartera.Visible = False
'            grdAnalista.Col = grdAnalista.Col + 1
'            grdAnalista.SetFocus
'        End If
End Sub

Private Sub dcTipoClasificacion_Change()
        Dim bDatos As Boolean
        bDatos = False
        
        desHabilitarControles , , , , 1, 1, 1, 1, 1
        HabilitarControles 1
        
        dcAgencia.BoundText = 0
        If dcTipoClasificacion.BoundText <> "0" Then
           CargarMetasMinimas
            'If Me.dcAgencia.BoundText <> "0" Then
                'bDatos = ObtenerFechaUlitmoMeta
                'CargarListaAnalistaMetasActuales
            'Else
             '   MsgBox "Selecciones una Agencia", vbInformation
              '  Exit Sub
            'End If
            
        End If
        Dim i As Integer
        'If Not bDatos Then
            For i = 1 To grdAnalista.Rows - 1
                grdAnalista.TextMatrix(i, 10) = ""
                grdAnalista.TextMatrix(i, 11) = ""
                grdAnalista.TextMatrix(i, 12) = ""
                grdAnalista.TextMatrix(i, 13) = ""
    
            Next
        'End If
End Sub
Private Sub CargarMetasMinimas()
        Dim rsMetas As New Recordset
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        Set rsMetas.DataSource = objCOMNCredito.getCargarParametrosTipoClasificacion(Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText)
        
        If rsMetas.RecordCount > 0 Then
            lnNroClieMin = rsMetas!iMetaClieNue
            lnNroOpeMin = rsMetas!iMetaNroOpe
            lnSaldoMin = rsMetas!iMetaSaldo
'            grdAnalista.Enabled = True
'            cmdRegistrar.Enabled = True
'            cmdCancelar.Enabled = True
'            cmdEditar.Enabled = False
        Else
            lnNroClieMin = 0
            lnNroOpeMin = 0
            lnSaldoMin = 0
'            grdAnalista.Enabled = False
'            cmdRegistrar.Enabled = False
'            cmdCancelar.Enabled = False
'            cmdEditar.Enabled = False
            MsgBox "NO se han Registrado los Parametros para esta Clasificacion", vbInformation
            Me.dcTipoClasificacion.BoundText = "0"
        End If
    
End Sub
Private Function ObtenerFechaUlitmoMeta() As Boolean
        ObtenerFechaUlitmoMeta = False
        Dim rsMetas As New Recordset
        Dim i As Integer
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        Set rsMetas.DataSource = objCOMNCredito.getFechaUlitmoMeta(Me.dcAgencia.BoundText, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText)
        
        If Not IsNull(rsMetas!cFechaMeta) Then
            lsFechaMeta = rsMetas!cFechaMeta
            cmbMes.ListIndex = val((Mid(rsMetas!cFechaMeta, 5, 2)))
            i = getIndiceAnio(Mid(rsMetas!cFechaMeta, 1, 4))
            cmbAnio.ListIndex = i
            
                       
            HabilitarControles , 1, 1, , , 1
            desHabilitarControles , , , , 1, , 1, 1
            
            CargarListaAnalistaMetasActuales (lsFechaMeta)
            ObtenerFechaUlitmoMeta = True
        Else
            CargaListaAnalista
            lsFechaMeta = "201103"
            cmbMes.ListIndex = 3
            cmbAnio.ListIndex = 1
            
            HabilitarControles , , , , 1, , 1, 1
            desHabilitarControles , , , , , 1
            'CheckAnalista
        End If
    
End Function

Private Sub CargarListaAnalistaMetasActuales(ByVal sFechaMeta As String)
        Dim rsMetas As New Recordset
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        Set rsMetas.DataSource = objCOMNCredito.getListaAnalistaMetas(Me.dcAgencia.BoundText, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, sFechaMeta)
        
        Dim i As Integer
        Dim J As Integer
        
        If rsMetas.RecordCount > 0 Then
            CargaListaAnalista
            For i = 1 To grdAnalista.Rows - 1
                rsMetas.MoveFirst
                For J = 0 To rsMetas.RecordCount - 1
                    If grdAnalista.TextMatrix(i, 2) = rsMetas!cCodAna Then
                       grdAnalista.TextMatrix(i, 1) = Chr(254)
                       grdAnalista.TextMatrix(i, 5) = rsMetas!nClieNue
                       grdAnalista.TextMatrix(i, 6) = rsMetas!nNroOpe
                       grdAnalista.TextMatrix(i, 7) = Format(rsMetas!nSaldoCar, "##,#00")
                       
                       Exit For
                    Else
                       grdAnalista.TextMatrix(i, 1) = Chr(168)
                                       
                    End If
                     rsMetas.MoveNext
                Next
               
            Next
        End If
End Sub
Private Sub CargarListaAnalistaMetasAnteriores(ByVal sFechaMeta As String)
        Dim rsMetas As New Recordset
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        Set rsMetas.DataSource = objCOMNCredito.getListaAnalistaMetasAnteriores(Me.dcAgencia.BoundText, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, sFechaMeta)
        
        Dim i As Integer
                
        Call LimpiaFlex(grdAnalista)
        If rsMetas.RecordCount > 0 Then
           
            grdAnalista.Rows = grdAnalista.Rows + rsMetas.RecordCount - 1
            'ReDim matris(grdAnalista.Rows, grdAnalista.Cols - 5)
                       
            grdAnalista.Redraw = False
            For i = 0 To rsMetas.RecordCount - 1
                
                grdAnalista.TextMatrix(i + 1, 0) = i + 1
                grdAnalista.TextMatrix(i + 1, 1) = Chr(254)
                    'convierte  el caracter Chr(254) a check=true
                    grdAnalista.Row = i + 1
                    grdAnalista.Col = 1
                    grdAnalista.CellFontName = "Wingdings"
                    grdAnalista.CellFontSize = 13
                               
                grdAnalista.TextMatrix(i + 1, 2) = rsMetas!cCodAna
                grdAnalista.TextMatrix(i + 1, 3) = rsMetas!cUser
                grdAnalista.TextMatrix(i + 1, 4) = PstaNombre(rsMetas!cPersNombre)
                grdAnalista.TextMatrix(i + 1, 5) = rsMetas!nClieNue
                grdAnalista.TextMatrix(i + 1, 6) = rsMetas!nNroOpe
                grdAnalista.TextMatrix(i + 1, 7) = Format(rsMetas!nSaldoCar, "##,#00")
                rsMetas.MoveNext
            Next
            grdAnalista.Redraw = True
        End If
           
        
End Sub

Private Function getIndiceAnio(ByVal sAnio As String) As Integer
    Dim i As Integer
    For i = 1 To Me.cmbAnio.ListCount
        If cmbAnio.List(i) = sAnio Then
            getIndiceAnio = i
            Exit For
        End If
    Next
    
End Function
Private Sub Form_Load()
   ConfigGridAnalista
   CargarTipoCartera
   CargarAgencias
   cargarAnio
    
   'CargarComboCordinador
End Sub
Private Sub cargarAnio()
    Dim i As Integer
'    Dim sFecha As String
    Dim indice As Integer
    indice = 1
    cmbAnio.AddItem "--Año--", 0
    For i = 2011 To val(Year(Date))
        cmbAnio.AddItem i
        cmbAnio.ItemData(indice) = i
        indice = indice + 1
    Next
    
'    sFecha = gdFecSis
    
'    cmbMes.ListIndex = val(Mid(gdFecSis, 4, 2))
'    i = getIndiceAnio(Mid(gdFecSis, 7, 4))
'    cmbAnio.ListIndex = i
    
End Sub
Private Sub ConfigGridAnalista()
    grdAnalista.Clear
    grdAnalista.Rows = 2
    
    With grdAnalista
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Op"
        .TextMatrix(0, 2) = "Codigo"
        .TextMatrix(0, 3) = "User"
        '.TextMatrix(0, 4) = "Nombre" ' se modifico con NºOpe
        .TextMatrix(0, 4) = "Nº Cred"
        .TextMatrix(0, 5) = "Nº Clie"
        .TextMatrix(0, 6) = "Saldo Cart"
        .TextMatrix(0, 7) = "Nº Ope"
        .TextMatrix(0, 8) = "Mora"
        .TextMatrix(0, 9) = "Tasa P"
        .TextMatrix(0, 10) = "Coord"
        .TextMatrix(0, 11) = "Meta Clie"
        .TextMatrix(0, 12) = "Meta Ope"
        .TextMatrix(0, 13) = "Meta Saldo"
       
        
        .ColWidth(0) = 300
        .ColWidth(1) = 350
        .ColWidth(2) = 1350 'Codigo
        .ColWidth(3) = 600
        .ColWidth(4) = 750 'Nº Cred
        .ColWidth(5) = 700 'Nº Clie
        .ColWidth(6) = 1150 'Saldo Car
        .ColWidth(7) = 700 'Nº Ope
        .ColWidth(8) = 550 'Mora
        .ColWidth(9) = 700 'Tasa Pon
        .ColWidth(10) = 1000 'Coord
        .ColWidth(11) = 900
        .ColWidth(12) = 950
        .ColWidth(13) = 1200
        
        .ColAlignment(1) = flexAlignCenterCenter
         
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter 'Saldo Car
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
       
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignRightCenter
         
        .ColAlignmentFixed(2) = flexAlignCenterCenter 'Codigo
        .ColAlignmentFixed(6) = flexAlignCenterCenter 'Saldo Car
        .ColAlignmentFixed(10) = flexAlignCenterCenter 'Coord
        .ColAlignmentFixed(11) = flexAlignCenterCenter 'Nº Cred
        .ColAlignmentFixed(12) = flexAlignCenterCenter 'Nº Clie
        .ColAlignmentFixed(13) = flexAlignCenterCenter 'Saldo
    End With
End Sub
Private Sub CargarAgencias()
    Dim rsAgencia As New ADODB.Recordset
    Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsAgencia.DataSource = objCOMNCredito.getCargarAgencias
    dcAgencia.BoundColumn = "cAgeCod"
    dcAgencia.DataField = "cAgeCod"
    Set dcAgencia.RowSource = rsAgencia
    dcAgencia.ListField = "cAgeDescripcion"
    dcAgencia.BoundText = 0
End Sub
Private Sub CargaListaAnalista()
        Me.MousePointer = vbHourglass
        Dim i As Integer
       
        Dim oGen As COMDConstSistema.DCOMGeneral
        Set oGen = New COMDConstSistema.DCOMGeneral
        Dim clsTC As COMDConstSistema.NCOMTipoCambio
        Dim nTC As Double
        Dim dFecha As Date
        Dim dFechaMesAnt As Date
        Dim dFechaTC As Date
        Set clsTC = New COMDConstSistema.NCOMTipoCambio
        
        dFecha = DateAdd("d", -Day(gdFecSis), gdFecSis)
        dFechaTC = DateAdd("d", 1, dFecha)
        nTC = clsTC.EmiteTipoCambio(dFechaTC, TCFijoDia)

'        Dim sAnalistas As String
'        sAnalistas = oGen.LeeConstSistema(gConstSistRHCargoCodAnalistas)
'        sAnalistas = Replace(sAnalistas, "'", "")
        Dim rsAnalista As New ADODB.Recordset
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        'Set rsAnalista.DataSource = objCOMNCredito.getCargarAnalistasXAgencia(Me.dcAgencia.BoundText, sAnalistas, nTC, Me.dcTipoCartera.BoundText)
        
        'verificar si se ha consolidado del mes anterior
        Dim rsBono As Recordset
        Set rsBono = objCOMNCredito.obtenerBonoAnalistasConsol(Me.dcAgencia.BoundText, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, dFecha)
        If Not rsBono.EOF Or Not rsBono.BOF Then
            Me.cmdConsolidar.Enabled = False
        Else
            Me.cmdConsolidar.Enabled = True
        End If
        '************************************************
        Set rsAnalista.DataSource = objCOMNCredito.getCargarAnalistasXAgencia(Me.dcAgencia.BoundText, nTC, Me.dcTipoCartera.BoundText, Me.dcTipoClasificacion.BoundText, dFecha)
        
        
        Call LimpiaFlex(grdAnalista)
        If Not rsAnalista.EOF Or rsAnalista.BOF Then
            If rsAnalista.RecordCount > 0 Then
                chkAnalista.Enabled = True
                 grdAnalista.Rows = grdAnalista.Rows + rsAnalista.RecordCount - 1
                 'ReDim matris(grdAnalista.Rows, grdAnalista.Cols - 5)
                            
                 grdAnalista.Redraw = False
                 For i = 0 To rsAnalista.RecordCount - 1
                     
                     grdAnalista.TextMatrix(i + 1, 0) = i + 1
                     grdAnalista.TextMatrix(i + 1, 1) = Chr(254)
                         'convierte  el caracter Chr(254) a check=true
                         grdAnalista.Row = i + 1
                         grdAnalista.Col = 1
                         grdAnalista.CellFontName = "Wingdings"
                         grdAnalista.CellFontSize = 13
                                    
                     grdAnalista.TextMatrix(i + 1, 2) = rsAnalista!cPersCod
                     grdAnalista.TextMatrix(i + 1, 3) = rsAnalista!Usuario
                     'grdAnalista.TextMatrix(i + 1, 4) = PstaNombre(rsAnalista!cPersNombre)
                     grdAnalista.TextMatrix(i + 1, 4) = rsAnalista!TOTALCRED
                     grdAnalista.TextMatrix(i + 1, 5) = rsAnalista!TOTALCLIE
                     grdAnalista.TextMatrix(i + 1, 6) = Format(rsAnalista!TOTALSALDO, "##,#00.00")
                     grdAnalista.TextMatrix(i + 1, 7) = rsAnalista!TotalNroOpe
                     grdAnalista.TextMatrix(i + 1, 8) = Round(rsAnalista!TOTALMORA, 2)
                     grdAnalista.TextMatrix(i + 1, 9) = Round(rsAnalista!TOTALTASA, 2)
                     
                     
                     rsAnalista.MoveNext
                 Next
                 grdAnalista.Redraw = True
             End If
        End If
Me.MousePointer = vbDefault
End Sub
Private Sub cmdVisitasEliminar_Click()
    If MsgBox("¿¿Está seguro de eliminar al Analista " + grdAnalista.TextMatrix(grdAnalista.Row, 2) + "??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdAnalista.RemoveItem (grdAnalista.Row)
        Dim i As Integer
        For i = 1 To grdAnalista.Rows - 1
            grdAnalista.TextMatrix(i, 0) = i
        Next
    End If
End Sub


Private Function VerifcarMinimoCliente(ByVal NroClie As Integer, Optional fila As Integer = 0) As Boolean
 VerifcarMinimoCliente = False
        If Me.dcTipoClasificacion.BoundText = "" Or Me.dcTipoCartera.BoundText = 0 Then
            MsgBox ("Debe Seleccionar un Tipo de Cartera")
            VerifcarMinimoCliente = True
            Exit Function
        ElseIf Me.dcTipoClasificacion.BoundText = 0 Then
            MsgBox ("Debe Seleccionar la Clasificacion de la Cartera")
            VerifcarMinimoCliente = True
            Exit Function
        End If
 
                            '-------------------------CARTERA 1
                            If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 1 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 2 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 3 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 4 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            '--------------------------CARTERA 2
                            If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 1 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 2 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 3 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            '--------------------------CARTERA 3
                            If Me.dcTipoCartera.BoundText = 3 And Me.dcTipoClasificacion.BoundText = 1 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            If Me.dcTipoCartera.BoundText = 3 And Me.dcTipoClasificacion.BoundText = 2 And NroClie < lnNroClieMin Then
                                MsgBox ("El Nro Minimo de Clientes Nuevos para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroClieMin) + " en la Fila Nro:" + CStr(fila))
                                VerifcarMinimoCliente = True
                                Exit Function
                            End If
                            
End Function
Private Function VerifcarMinimoOperaciones(ByVal NroOpe As Integer, Optional fila As Integer = 0) As Boolean
     VerifcarMinimoOperaciones = False
     
        If Me.dcTipoClasificacion.BoundText = "" Or Me.dcTipoCartera.BoundText = 0 Then
            MsgBox ("Debe Seleccionar un Tipo de Cartera")
            VerifcarMinimoOperaciones = True
            Exit Function
        ElseIf Me.dcTipoClasificacion.BoundText = 0 Then
            MsgBox ("Debe Seleccionar la Clasificacion de la Cartera")
            VerifcarMinimoOperaciones = True
            Exit Function
        End If
                '-------------------------CARTERA 1
                If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 1 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 2 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 3 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 1 And Me.dcTipoClasificacion.BoundText = 4 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                '--------------------------CARTERA 2
                If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 1 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 2 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 3 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                '--------------------------CARTERA 3
                If Me.dcTipoCartera.BoundText = 3 And Me.dcTipoClasificacion.BoundText = 1 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 3 And Me.dcTipoClasificacion.BoundText = 2 And NroOpe < lnNroOpeMin Then
                    MsgBox ("El Nro Minimo de Operaciones  para este Tipo de Cartera y su Clasificacion es " + CStr(lnNroOpeMin) + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoOperaciones = True
                    Exit Function
                End If
                
            
           
        
End Function
Private Function VerifcarMinimoSaldoCartera(ByVal Saldo As Long, Optional fila As Integer = 0) As Boolean
        
        VerifcarMinimoSaldoCartera = False
        
        If Me.dcTipoClasificacion.BoundText = "" Or Me.dcTipoCartera.BoundText = 0 Then
            MsgBox ("Debe Seleccionar un Tipo de Cartera")
            VerifcarMinimoSaldoCartera = True
            Exit Function
        ElseIf Me.dcTipoClasificacion.BoundText = 0 Then
            MsgBox ("Debe Seleccionar la Clasificacion de la Cartera")
            VerifcarMinimoSaldoCartera = True
            Exit Function
        End If
                '-------------------------CARTERA 1
                If Me.dcTipoCartera.BoundText = 1 And (Me.dcTipoClasificacion.BoundText = 1 Or Me.dcTipoClasificacion.BoundText = 2 Or Me.dcTipoClasificacion.BoundText = 3 Or Me.dcTipoClasificacion.BoundText = 4) And Saldo < lnSaldoMin Then
                    MsgBox ("El Monto Minimo de Saldo  para este Tipo de Cartera y su Clasificacion es S/." + Format(lnSaldoMin, "##,#00") + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoSaldoCartera = True
                    Exit Function
                End If
                
                '--------------------------CARTERA 2
                If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 1 And Saldo < lnSaldoMin Then
                    MsgBox ("El Monto Minimo de Saldo  para este Tipo de Cartera y su Clasificacion es S/." + Format(lnSaldoMin, "##,#00") + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoSaldoCartera = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 2 And Saldo < lnSaldoMin Then
                    MsgBox ("El Monto Minimo de Saldo  para este Tipo de Cartera y su Clasificacion es S/." + Format(lnSaldoMin, "##,#00") + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoSaldoCartera = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 2 And Me.dcTipoClasificacion.BoundText = 3 And Saldo < lnSaldoMin Then
                    MsgBox ("El Monto Minimo de Saldo  para este Tipo de Cartera y su Clasificacion es S/." + Format(lnSaldoMin, "##,#00") + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoSaldoCartera = True
                    Exit Function
                End If
                '--------------------------CARTERA 3
                If Me.dcTipoCartera.BoundText = 3 And Me.dcTipoClasificacion.BoundText = 1 And Saldo < lnSaldoMin Then
                    MsgBox ("El Monto Minimo de Saldo  para este Tipo de Cartera y su Clasificacion es S/." + Format(lnSaldoMin, "##,#00") + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoSaldoCartera = True
                    Exit Function
                End If
                If Me.dcTipoCartera.BoundText = 3 And Me.dcTipoClasificacion.BoundText = 2 And Saldo < lnSaldoMin Then
                    MsgBox ("El Monto Minimo de Saldo  para este Tipo de Cartera y su Clasificacion es S/." + Format(lnSaldoMin, "##,#00") + " en la Fila Nro:" + CStr(fila))
                    VerifcarMinimoSaldoCartera = True
                    Exit Function
                End If
                
End Function



Private Sub grdAnalista_Click()
   Dim s As String
   
   With grdAnalista
        If .Col = 1 Then
            If .TextMatrix(.Row, 1) = Chr(254) Then
                .TextMatrix(.Row, 1) = Chr(168)
                .TextMatrix(.Row, 10) = ""
                .TextMatrix(.Row, 11) = ""
                .TextMatrix(.Row, 12) = ""
                .TextMatrix(.Row, 13) = ""
            ElseIf .TextMatrix(.Row, 1) = Chr(168) Then
                .TextMatrix(.Row, 1) = Chr(254)
            End If
            s = .TextMatrix(.Row, 1)
'        ElseIf .Col = 10 Then        ' Position and size the ListBox, then show it.
'
'            dcCordinador.Width = .CellWidth
'            dcCordinador.Left = .CellLeft + .Left
'            dcCordinador.Top = .CellTop + .Top
'            dcCordinador.Visible = True

        End If
   
   End With
        
End Sub

Private Sub grdAnalista_EnterCell()
    With grdAnalista
        'If (.Col > 4 And .Col < (.Cols)) Or .Col > (.Cols - 4) Then 'JACA 20110330
        If (.Col > 9 And .Col < .Cols) Then
        .CellBackColor = &H80000018
        .Tag = ""
        End If
    End With
End Sub

Private Sub grdAnalista_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdAnalista
     Select Case KeyCode
        Case 46
            .Tag = grdAnalista
            'JACA 20110330 MODIFICADO 4 X 9
            If (.Col > 9 And .Col < (.Cols)) Then
                 'matris(.Row - 1, .col - 5) = ""
                 .TextMatrix(.Row, .Col) = ""
                 grdAnalista = ""
            End If
        

      End Select
    End With
End Sub

Private Sub grdAnalista_KeyPress(KeyAscii As Integer)
 
With grdAnalista
    'si es enter
    If KeyAscii = 13 Then
           ' If .Col = 10 Then
                
'                  If dcCordinador.Visible = False Then
'                      dcCordinador.Width = .CellWidth
'                        dcCordinador.Left = .CellLeft + .Left
'                        dcCordinador.Top = .CellTop + .Top
'                        dcCordinador.Visible = True
'                        dcCordinador.SetFocus
'                  Else
'                    grdAnalista.TextMatrix(grdAnalista.Row, 10) = dcCordinador.Text
'                    dcCordinador.Visible = False
'                    grdAnalista.Col = grdAnalista.Col + 1
'                    grdAnalista.CellBackColor = &H80000018
'                    grdAnalista.SetFocus
'                  End If
'
'                  'dcCordinador.Visible = False
'                  Exit Sub
           ' End If
            'si esta en las columnas 5 a 6 de edicion
            'JACA 20110330 MODIFICADO (4 X10) Y  (7X13)
            If .Col > 9 And .Col < 13 And .TextMatrix(.Row, .Col) <> "" Then
                If .Col = 10 Then
                    If Len(.TextMatrix(.Row, .Col)) < 4 Then
                        MsgBox "Debe Ingresar un Coordinador Correcto", vbInformation
                    End If
                
                End If
                If .Col = 11 Then
                    If .TextMatrix(.Row, .Col) <= 10000 Then ' para evitar desbordamiento
                        If VerifcarMinimoCliente(.TextMatrix(.Row, .Col), .Row) Then
                            .TextMatrix(.Row, .Col) = ""
                            Exit Sub
                         End If
                    Else
                         .TextMatrix(.Row, .Col) = ""
                        MsgBox "Solo esta permitido hasta 10,000", vbInformation
                        Exit Sub
                    End If
                End If
                If .Col = 12 Then
                    If .TextMatrix(.Row, .Col) <= 10000 Then ' para evitar desbordamiento
                        If VerifcarMinimoOperaciones(.TextMatrix(.Row, .Col), .Row) Then
                             .TextMatrix(.Row, .Col) = ""
                            Exit Sub
                        End If
                    Else
                         .TextMatrix(.Row, .Col) = ""
                        MsgBox "Solo esta permitido hasta 10,000", vbInformation
                        Exit Sub
                    End If
                End If
                
                .CellBackColor = &H8000000E
                .Row = .Row
                .Col = .Col + 1
                .CellBackColor = &H80000018
            'si esta en la ultima columna para pasar a la sigte fila
            ElseIf .Row <= .Rows - 1 And .TextMatrix(.Row, .Col) <> "" Then
                If .Col = 13 Then
                    
                    'verifica si es menos de 100,000,000 para evitar desbordamiento
                    If .TextMatrix(.Row, .Col) <= 100000000 Then
                        If VerifcarMinimoSaldoCartera(.TextMatrix(.Row, .Col), .Row) Then
                                 .TextMatrix(.Row, .Col) = ""
                                Exit Sub
                        End If
                    Else
                         .TextMatrix(.Row, .Col) = ""
                        MsgBox "Solo esta permitido hasta S/.100,000,000", vbInformation
                        Exit Sub
                    End If
                End If
                .CellBackColor = &H8000000E
                'verifica si no es la ultima fila y pasa a la siguiente fila en la coumna 5
                
                    If .Row < .Rows - 1 Then
                        '.Row = .Row + 1
                        Dim i As Integer
                        For i = .Row + 1 To .Rows - 1
                            If .TextMatrix(i, 1) = Chr(254) Then
                                  .Col = 10 '5
                                  .Row = i
                                  Exit For
                            End If
                        
                        Next
                    Else
                        .CellBackColor = &H8000000E
                        Me.cmdRegistrar.SetFocus
                    End If
                
                .CellBackColor = &H80000018
            End If
    
    ElseIf KeyAscii = 8 Then 'si es retroceso
            If .TextMatrix(.Row, 1) = Chr(254) Then ' si esta checkeado
                If Len(.TextMatrix(.Row, .Col)) > 0 Then
                
                        .TextMatrix(.Row, .Col) = Mid(.TextMatrix(.Row, .Col), 1, Len(.TextMatrix(.Row, .Col)) - 1)
                        If .Col = 13 Then
                
                            grdAnalista = Format(.TextMatrix(.Row, .Col), "##,##0")
                        Else
                
                            grdAnalista = .TextMatrix(.Row, .Col)
                        End If
                End If
             End If
    ElseIf (.Col > 9 And .Col < (.Cols)) Then
            If .TextMatrix(.Row, 1) = Chr(254) Then ' si esta checkeado
                If (InStr("0123456789", Chr(KeyAscii)) = 0) And .Col > 10 Then 'compara si es un numero
                        KeyAscii = 0
                        .TextMatrix(.Row, .Col) = ""
                
                ElseIf (InStr("0123456789", Chr(KeyAscii)) <> 0) And .Col = 10 Then 'compara si es letra
                        KeyAscii = 0
                        .TextMatrix(.Row, .Col) = ""
                Else 'entra si es un numero
                        .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) + Chr(KeyAscii)
                        
                        If .Col = 10 Then
                            If Len(.TextMatrix(.Row, .Col)) > 4 Then
                                .TextMatrix(.Row, .Col) = Mid(.TextMatrix(.Row, .Col), 1, Len(.TextMatrix(.Row, .Col)) - 1)
                            End If
                            .TextMatrix(.Row, .Col) = UCase(.TextMatrix(.Row, .Col))
                        ElseIf .Col = 13 Then
                            grdAnalista = Format(.TextMatrix(.Row, .Col), "##,##0")
                        Else
                            grdAnalista = .TextMatrix(.Row, .Col)
                        End If
                End If
            End If
    End If
                    
End With
End Sub

Private Sub grdAnalista_LeaveCell()
    With grdAnalista
        Dim s As String
        If (.Col > 9 And .Col < .Cols) Then
              .CellBackColor = &H8000000E
              
              If .TextMatrix(.Row, .Col) <> "" Then
                    If .Col = 11 Then
                         If VerifcarMinimoCliente(.TextMatrix(.Row, .Col), .Row) Then
                             .TextMatrix(.Row, .Col) = ""
                             Exit Sub
                          End If
                     End If
                     If .Col = 12 Then
                         If VerifcarMinimoOperaciones(.TextMatrix(.Row, .Col), .Row) Then
                              .TextMatrix(.Row, .Col) = ""
                             Exit Sub
                         End If
                     End If
                     
                     If .Col = 13 Then
                         'If VerifcarMinimoSaldoCartera(matris(.Row - 1, .col - 5)) Then
                         If VerifcarMinimoSaldoCartera(.TextMatrix(.Row, .Col), .Row) Then
                                  .TextMatrix(.Row, .Col) = ""
                                 Exit Sub
                         End If
                     End If
              End If
               
            
        End If
        
    End With
End Sub

'Private Sub txtFCierre_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdRegistrar.Enabled = True
'        Me.cmdRegistrar.SetFocus
'    End If
'End Sub


'Private Sub LimpiarText()
'    Me.txtCientesNuevos.Text = ""
'    Me.txtSaldoCartera.Text = ""
'    Me.txtNroOperaciones.Text = ""
'    Me.txtFCierre.Text = "__/__/____"
'
'    Me.txtCientesNuevos.Enabled = False
'    Me.txtSaldoCartera.Enabled = False
'    Me.txtNroOperaciones.Enabled = False
'    Me.txtFCierre.Enabled = False
'
'    Me.cmdRegistrar.Enabled = False
'End Sub
Private Sub LimpiarControles()
    Me.dcAgencia.BoundText = 0
    Me.dcTipoCartera.BoundText = 0
    Me.dcTipoClasificacion.BoundText = 0
    Me.cmbAnio.ListIndex = 0
    Me.cmbMes.ListIndex = 0
    
    desHabilitarControles , , , 1, 1, 1, 1, 1, 1
    chkAnalista.value = 1
    Call LimpiaFlex(grdAnalista)
   
End Sub
Sub CargarComboCordinador()
    grdAnalista.RowHeightMin = dcCordinador.Height
    dcCordinador.Visible = False
    dcCordinador.Width = grdAnalista.CellWidth
End Sub
Sub CargarCordinadoresAgencia()
        
        Dim rsCordinador As New ADODB.Recordset
        Set objCOMNCredito = New COMNCredito.NCOMBPPR
        Set rsCordinador.DataSource = objCOMNCredito.getCargarCordinadorXAgencia(Me.dcAgencia.BoundText)
        dcCordinador.BoundColumn = "cPersCod"
        dcCordinador.DataField = "cPersCod"
        Set dcCordinador.RowSource = rsCordinador
       ' dcCordinador.ListField = "cUser" + Space(20) + "cPersCod"
        dcCordinador.ListField = "cUser"
        dcCordinador.BoundText = 0
        
End Sub
Private Sub chkAnalista_Click()
    If chkAnalista.value = 1 Then
            CheckAnalista 254
    Else
            CheckAnalista 168
    End If
End Sub
Private Sub CheckAnalista(ByVal op As Integer)
    'Chr(168)
    Dim i As Integer
        For i = 1 To grdAnalista.Rows - 1
           grdAnalista.TextMatrix(i, 1) = Chr(op)
       Next
End Sub
Private Sub limpiarMetas()
    Dim i As Integer
        For i = 1 To grdAnalista.Rows - 1
            grdAnalista.TextMatrix(i, 10) = ""
            grdAnalista.TextMatrix(i, 11) = ""
            grdAnalista.TextMatrix(i, 12) = ""
            grdAnalista.TextMatrix(i, 13) = ""
        Next
        chkAnalista.value = 1
End Sub
Private Sub HabilitarControles(Optional pnAgencia As Integer = 0, Optional pnAnio As Integer = 0, Optional pnMes As Integer = 0, Optional pnAnterior As Integer = 0, Optional pnGuardar As Integer = 0, Optional pnEditar As Integer = 0, Optional pnCancelar As Integer = 0, Optional pngrdAna As Integer = 0, Optional pnchkAna As Integer = 0, Optional pnConsolidar As Integer = 0)
    
        
    If pnAgencia = 1 Then
        dcAgencia.Enabled = True
    End If
    
    If pnAnio = 1 Then
        cmbAnio.Enabled = True
   End If
    
    If pnMes = 1 Then
        cmbMes.Enabled = True
    End If
    
    If pnAnterior = 1 Then
        cmdAnterior.Enabled = True
    End If
    
    If pnGuardar = 1 Then
        cmdRegistrar.Enabled = True
   End If
      
    If pnEditar = 1 Then
        cmdEditar.Enabled = True
    End If
    
    If pnCancelar = 1 Then
        cmdCancelar.Enabled = True
    End If
    
     If pngrdAna = 1 Then
        Me.grdAnalista.Enabled = True
    End If
    
    If pnchkAna = 1 Then
        Me.chkAnalista.Enabled = True
    End If
    
     If pnConsolidar = 1 Then
        Me.cmdConsolidar.Enabled = True
    End If
    
End Sub
Private Sub desHabilitarControles(Optional pnAgencia As Integer = 0, Optional pnAnio As Integer = 0, Optional pnMes As Integer = 0, Optional pnAnterior As Integer = 0, Optional pnGuardar As Integer = 0, Optional pnEditar As Integer = 0, Optional pnCancelar As Integer = 0, Optional pngrdAna As Integer = 0, Optional pnchkAna As Integer = 0, Optional pnConsolidar As Integer = 0)
    
        
    If pnAgencia = 1 Then
        dcAgencia.Enabled = False
    End If
    
    If pnAnio = 1 Then
        cmbAnio.Enabled = False
    End If
    
    If pnMes = 1 Then
        cmbMes.Enabled = False
    End If
    
    If pnAnterior = 1 Then
        cmdAnterior.Enabled = False
    End If
    
    If pnGuardar = 1 Then
        cmdRegistrar.Enabled = False
    End If
      
    If pnEditar = 1 Then
        cmdEditar.Enabled = False
    End If
    
    If pnCancelar = 1 Then
        cmdCancelar.Enabled = False
    End If
    
    
    If pngrdAna = 1 Then
        Me.grdAnalista.Enabled = False
    End If
    
    
    If pnchkAna = 1 Then
        Me.chkAnalista.Enabled = False
    End If
    
    If pnConsolidar = 1 Then
        Me.cmdConsolidar.Enabled = False
    End If
End Sub
