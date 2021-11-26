VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Administracion de Tarjeta de Debito - CMAC MAYNAS S.A."
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   -1455
   ClientWidth     =   11670
   Icon            =   "MDIMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11994
            MinWidth        =   11994
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "26/05/2010"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "10:56 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "NÚM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuAfiliacion 
      Caption         =   "Afiliacion"
      Begin VB.Menu mnuAfilTarj 
         Caption         =   "Orden de Afiliacion "
         Begin VB.Menu mnuAfilTarjEmi 
            Caption         =   "Carga Archivo Tarjetas Emitidas"
         End
         Begin VB.Menu mnutarrechazadas 
            Caption         =   "Carga Archivo Tarjetas Rechazada"
         End
         Begin VB.Menu mnuafilTarjLote 
            Caption         =   "Carga Archivo de Tarjetas Ordenadas"
         End
      End
      Begin VB.Menu mnuacttarj 
         Caption         =   "Activacion de Tarjeta"
      End
      Begin VB.Menu mnucuenta 
         Caption         =   "Cuenta"
         Begin VB.Menu mnuctareg 
            Caption         =   "Vincular Cuenta"
         End
         Begin VB.Menu mnuctaconsulta 
            Caption         =   "Consulta"
         End
      End
      Begin VB.Menu mnutarjcta 
         Caption         =   "Tarjeta - Cuenta"
         Begin VB.Menu mnutarjctareg 
            Caption         =   "Registro"
         End
         Begin VB.Menu mnutrajctacons 
            Caption         =   "Consulta"
         End
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "Consultas"
      Begin VB.Menu mnuconstarj 
         Caption         =   "Consulta Tarjeta"
      End
      Begin VB.Menu mnuconsesttarj 
         Caption         =   "Consulta de Estado de Tarjeta"
      End
   End
   Begin VB.Menu mnuMantenimiento 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnucambsitu 
         Caption         =   "Bloqueo de Tarjeta"
      End
      Begin VB.Menu mnuasocpers 
         Caption         =   "Asociar Persona"
      End
      Begin VB.Menu mnuCambioClave 
         Caption         =   "Cambio de Clave"
      End
      Begin VB.Menu mnuImpSoliAfi 
         Caption         =   "Impresión Solicitud Afiliación"
      End
      Begin VB.Menu mnumanttarifario 
         Caption         =   "Mantenimiento Tarifario"
      End
      Begin VB.Menu mnumanTipoCta 
         Caption         =   "Mantenimiento Tipos de Cuenta"
      End
      Begin VB.Menu mnutareptar 
         Caption         =   "Tarifario de Reposición de Tarjeta"
      End
      Begin VB.Menu mnulimoper 
         Caption         =   "Limites Operativos"
         Begin VB.Menu mnumonmaxretdia 
            Caption         =   "Monto Maximo de Retiro por Dia"
         End
         Begin VB.Menu mnuoperlib 
            Caption         =   "Operaciones Libres"
         End
         Begin VB.Menu mnunumretpordia 
            Caption         =   "Numero de Retiros por Dia"
         End
         Begin VB.Menu mnumontomaxderet 
            Caption         =   "Montos Maximos de Retiro"
         End
      End
      Begin VB.Menu mnuconfpinpad 
         Caption         =   "Configuracion de PINPAD"
      End
      Begin VB.Menu mnumantpermisos 
         Caption         =   "Mantenimiento de Permisos"
      End
      Begin VB.Menu mnulimrementrans 
         Caption         =   "Limite de Remesas en Transito"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnutarjafil 
         Caption         =   "Tarjetas Afiliadas"
      End
      Begin VB.Menu mnutarjbloq 
         Caption         =   "Tarjetas Bloqueadas"
      End
      Begin VB.Menu mnureptarjcancel 
         Caption         =   "Tarjetas Canceladas"
      End
      Begin VB.Menu mnureptarjact 
         Caption         =   "Tarjetas Activadas"
      End
      Begin VB.Menu mnudettarj 
         Caption         =   "Detalle de Tarjetas"
      End
      Begin VB.Menu mnureptarjrech 
         Caption         =   "Tarjetas Rechazadas"
      End
      Begin VB.Menu mnureptarjord 
         Caption         =   "Tarjetas Ordenadas"
      End
      Begin VB.Menu mnurepcontrol 
         Caption         =   "Control Operaciones ATM"
      End
   End
   Begin VB.Menu mnuStock 
      Caption         =   "Control de Stock"
      Begin VB.Menu mnuInterno 
         Caption         =   "Interno"
         Begin VB.Menu mnuHabilitarTarjetas 
            Caption         =   "Habilitar Tarjetas"
         End
         Begin VB.Menu mnuDevolverTarjetas 
            Caption         =   "Devolver Tarjetas"
         End
         Begin VB.Menu mnuStockActual 
            Caption         =   "Registrar Stock Actual Ventanilla"
         End
         Begin VB.Menu mnuRegStckActBvda 
            Caption         =   "Registrar Stock Actual Boveda"
         End
         Begin VB.Menu mnuCuadrarTarjBvda 
            Caption         =   "Cuadrar Tarjetas Boveda"
         End
         Begin VB.Menu mnuCuadrarTarjetas 
            Caption         =   "Cuadrar Tarjetas Ventanilla"
         End
         Begin VB.Menu mnucuadreconsoltarjeta 
            Caption         =   "Cuadre Consolidado de Tarjetas"
         End
      End
      Begin VB.Menu mnuExterno 
         Caption         =   "Externo"
         Begin VB.Menu mnuIngTar 
            Caption         =   "Ingreso de Tarjetas"
         End
         Begin VB.Menu mnuSalTar 
            Caption         =   "Salida de Tarjetas"
         End
         Begin VB.Menu mnuRemesas 
            Caption         =   "Remesas"
            Begin VB.Menu mnuConfRemesas 
               Caption         =   "Confirma Remesas"
            End
            Begin VB.Menu mnustockconfdevol 
               Caption         =   "Confirmar Devoluciones"
            End
            Begin VB.Menu mnuStckRemesas 
               Caption         =   "Remesar Tarjetas"
            End
            Begin VB.Menu mnuremdevol 
               Caption         =   "Devoluciones Tarjetas"
            End
         End
      End
      Begin VB.Menu mnuregstockbovgen 
         Caption         =   "Registro de Stock Boveda General"
      End
      Begin VB.Menu mnusaldoStockBovGen 
         Caption         =   "Saldo Stock Boveda General"
      End
      Begin VB.Menu mnumonstock 
         Caption         =   "Actualizacion de Stock Minimo"
      End
      Begin VB.Menu mnuCantAct 
         Caption         =   "Consulta de Cantidades Actuales"
      End
      Begin VB.Menu mnuExtornos 
         Caption         =   "Extornos"
         Begin VB.Menu mnuextIngSalidas 
            Caption         =   "Ingresos y Salidas"
         End
         Begin VB.Menu mnuextRemesa 
            Caption         =   "Remesas"
         End
         Begin VB.Menu mnuextdevol 
            Caption         =   "Devoluciones"
         End
         Begin VB.Menu mnuHabDev 
            Caption         =   "Habilitaciones y Devoluciones"
         End
         Begin VB.Menu mnuextconfremesa 
            Caption         =   "Confirmacion de Remesa"
         End
      End
      Begin VB.Menu mnuReportesGen 
         Caption         =   "Reportes Generales"
         Begin VB.Menu mnuAgStckMin 
            Caption         =   "Consulta de Agencias con Stock Minimo"
         End
         Begin VB.Menu mnurepgenlisstockcajero 
            Caption         =   "Listado de Stocks por Cajero"
         End
         Begin VB.Menu mnuestadstockreg 
            Caption         =   "Estadisticas de Registro de Stock"
         End
         Begin VB.Menu mnulistmovbgen 
            Caption         =   "Listado de Movimientos de Boveda General"
         End
         Begin VB.Menu mnurementransito 
            Caption         =   "Listado de Remesas en Transito"
         End
         Begin VB.Menu mnureplimrementra 
            Caption         =   "Remesas en Transito fuera del Limite"
         End
         Begin VB.Menu mnuStockRepTarPorAge 
            Caption         =   "Tarjetas por Agencia"
         End
         Begin VB.Menu mnurepopeconf 
            Caption         =   "Operaciones Confirmadas"
         End
         Begin VB.Menu mnureptarjemiporage 
            Caption         =   "Tarjetas Emitidas por Agencia"
         End
         Begin VB.Menu mnureptarreti 
            Caption         =   "Reporte de Tarjetas Retiradas"
         End
         Begin VB.Menu MnuRepBovedaAgenUsuarios 
            Caption         =   "Movimientos Boveda Agencias y Usuarios"
         End
      End
   End
   Begin VB.Menu mnuConciliacion 
      Caption         =   "Conciliacion"
      Begin VB.Menu mnuCargarArchivoLog 
         Caption         =   "Cargar Archivo Log"
      End
      Begin VB.Menu mnuconfirmaciondeoperaciones 
         Caption         =   "Confirmacion de Operaciones"
      End
      Begin VB.Menu mnuReportesC 
         Caption         =   "Reportes"
         Begin VB.Menu mnuOperacionesRealizadas 
            Caption         =   "Log Caja"
         End
         Begin VB.Menu mnuOperacionesConfirmadas 
            Caption         =   "Log Caja Conciliadas"
         End
         Begin VB.Menu mnuOperacionesPendientes 
            Caption         =   "Log Caja Pendientes"
         End
         Begin VB.Menu mnuSeparador1 
            Caption         =   "_______________________"
         End
         Begin VB.Menu mnuOpeConfirmadas 
            Caption         =   "Log Interbank"
         End
         Begin VB.Menu mnuOperacionesConciliadas 
            Caption         =   "Log Interbank Conciliadas"
         End
         Begin VB.Menu mnuOperacionesConfirmadasnorealizadas 
            Caption         =   "Log Interbank Pendientes"
         End
         Begin VB.Menu mnuSeparador2 
            Caption         =   "_______________________"
         End
         Begin VB.Menu mnuResumenOp 
            Caption         =   "Resumen Operaciones"
         End
         Begin VB.Menu mnuSeparador3 
            Caption         =   "_______________________"
         End
         Begin VB.Menu mnuRepTarjetasRetenidas 
            Caption         =   "Reposicón de Tarjetas"
         End
      End
   End
   Begin VB.Menu mnuPITOpeInterCMACS 
      Caption         =   "Operaciones InterCMACS"
      Begin VB.Menu mnuPITConciliacion 
         Caption         =   "Conciliación"
         Begin VB.Menu mnuPITConciCargaLog 
            Caption         =   "Carga de Archivo LOG"
         End
         Begin VB.Menu mnuPITConciConciliacion 
            Caption         =   "Conciliar Operaciones"
         End
      End
      Begin VB.Menu mnuPITReportes 
         Caption         =   "Reportes"
         Begin VB.Menu mnuPITRepOpeInterCMAC 
            Caption         =   "Operaciones Inter CMAC"
         End
      End
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CargaMenu()
Dim sCadGrupos As String
Dim R As ADODB.Recordset
Dim sSQL As String
Dim Ctl As Control
Dim sTipo As String
                    
        sCadGrupos = CargaGruposDelUsuario(gsCodUser, gsDominio)

'        sSql = "Select DISTINCT M.cNomMenu "
'        sSql = sSql & " from Menu M"
'        sSql = sSql & "     Inner Join Permisos P ON P.cNomMenu = M.cNomMenu"
'        sSql = sSql & "     Where P.cNomgrupo in (" & sCadGrupos & ")"
'
        sSQL = "Select DISTINCT P.cNomMenu "
        sSQL = sSQL & "     From  Permisos P "
        sSQL = sSQL & "     Where P.cNomgrupo in (" & sCadGrupos & ")"
        
On Error Resume Next

            For Each Ctl In MDIMenu.Controls
                sTipo = TypeName(Ctl)
                If sTipo = "Menu" Then
                    If Ctl.Visible = True Then
                        Ctl.Visible = False
                        Ctl.Enabled = False
                    End If
                End If
            Next
            
        Set R = New ADODB.Recordset
        oConec.AbreConexion
        R.Open sSQL, oConec.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not R.EOF

        
            For Each Ctl In MDIMenu.Controls
                sTipo = TypeName(Ctl)
''                If UCase(Ctl.Name) = UCase("mnuCuadrarTarjetas") Then
''                    sTipo = sTipo
''                End If
                If UCase(sTipo) = "MENU" Then
                    If UCase(Ctl.Name) = UCase(R!cNomMenu) Then
                        Ctl.Visible = True
                        Ctl.Enabled = True
                        Exit For
                    End If
                End If
            Next
        
            R.MoveNext
        Loop
        R.Close
        C.Close
        oConec.CierraConexion
End Sub

Private Sub MDIForm_Load()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As New ADODB.Recordset

    Set oConec = New DConecta
    
    gsNomMaquinaUsu = GetMaquinaUsuario
    Call RecuperaConfigPinPad
    
    If gnTipoPinPad = 0 Then
        MsgBox "UD NO Tiene Configurado su PINPAD, Por favor Configure el Tipo de PINPAD a USAR", vbExclamation, "Aviso"
        frmPinpadSelec.Show 1
    End If
        
        
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psUser", adVarChar, adParamInput, 10, gsCodUser)
    Cmd.Parameters.Append Prm
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psAge", adVarChar, adParamOutput, 10)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNomAge", adVarChar, adParamOutput, 50)
    Cmd.Parameters.Append Prm

    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNomUsu", adVarChar, adParamOutput, 150)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_DevuelveAgenciaDeUsuario"
    Cmd.Execute
       
    gsCodAge = IIf(IsNull(Cmd.Parameters(1).Value), "00", Cmd.Parameters(1).Value)
    gsNomAge = IIf(IsNull(Cmd.Parameters(2).Value), "NINGUNA", Cmd.Parameters(2).Value)
    gsNomUser = IIf(IsNull(Cmd.Parameters(3).Value), "", Cmd.Parameters(3).Value)
    
    gsServerName = oConec.ServerName
    gsDatabaseName = oConec.DatabaseName

    Set Cmd = Nothing
    Set Prm = Nothing

    oConec.CierraConexion
    'gsCodAge = "01"
    'gsCodUser = "NSSE"
    'gsNomUser = "NAPOLEON SILVA"


    'FECHA DEL SISTEMA
    
    Set Cmd = New Command
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaFechaSistema"
    Set R = Cmd.Execute
    
    
    gdFecSis = CDate(Format(CDate(R!FechaSistema), "dd/mm/yyyy"))
    
    '**Modificado por DAOR 20100526 ****************************************
    'Me.Caption = Me.Caption & Space(5) & Format(gdFecSis, "dd/mm/yyyy") & Space(2) & gsCodAge & "-" & gsNomAge & Space(2) & gsCodUser & "-" & gsNomUser
    Me.Caption = "Admin Tarjetas - " & Trim(gsNomAge) & " - " & Trim(gsCodUser) & Space(10) & gsServerName & "\" & gsDatabaseName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    '***********************************************************************
       
    '**DAOR 20100406, Control de versión ***********************************
    Me.Caption = Me.Caption & Space(5) & " - Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-20100526" 'Cambiar la fecha cada vez que se compila
    '***********************************************************************
       
    Set Cmd = Nothing
    Set Prm = Nothing
    Set R = Nothing
     
    'CerrarConexion
    oConec.CierraConexion
        
    gsBIN = "810900"
        
    'Cargar Permisos
    Call CargaMenu
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub mnuacttarj_Click()
    frmActTarj.Show 1
End Sub

Private Sub mnuAfilTarjLin_Click()
   frmAfilTarj.Show 1
End Sub

Private Sub mnuAfilTarjEmi_Click()
    frmAfilCargaTarEmitidas.Show 1
End Sub

Private Sub mnuafilTarjLote_Click()
  frmRegistraOrdenesTarj.Show 1
End Sub

Private Sub mnuAgStckMin_Click()
    frmCtrlAgenciasStockMin.Show 1
End Sub

Private Sub mnuasocpers_Click()
    frmAsociaPers.Show 1
End Sub

Private Sub mnuCambioClave_Click()
    frmCambioClave.Show 1

End Sub

Private Sub mnucambsitu_Click()
    frmCambSitu.Show 1
End Sub

Private Sub mnuCantAct_Click()
frmStockCtrlAgeActual.Show 1
End Sub

Private Sub mnuCargarArchivoLog_Click()
    frmCargaLOG.Show 1
End Sub

Private Sub mnuconfirmaciondeoperaciones_Click()
    frmConfOperaConcilia.Show 1
End Sub

Private Sub mnuconfpinpad_Click()
    frmPinpadSelec.Show 1
End Sub

Private Sub mnuConfRemesas_Click()
    frmStockConfRemesa.Show 1
End Sub

Private Sub mnuconsesttarj_Click()
    frmConsEstTarj.Show 1
End Sub

Private Sub mnuconstarj_Click()
    frmConsTarj.Show 1
End Sub

Private Sub mnuctaconsulta_Click()
    frmCtaCons.Show 1
End Sub

Private Sub mnuctareg_Click()
    frmRegCta.Show 1
End Sub

Private Sub mnuCuadrarTarjBvda_Click()
  frmCuadreBvda.Show 1
End Sub

Private Sub mnucuadreconsoltarjeta_Click()
    frmCuadreConsolAge.Show 1
End Sub

Private Sub mnudettarj_Click()
    frmRepDetCtas.Show 1
End Sub

Private Sub mnuestadstockreg_Click()
    Call ListadoDEEstadisticasRegStockBGEN
End Sub

Private Sub mnuextconfremesa_Click()
    frmOpeExtdeConf.Show 1
End Sub

Private Sub mnuextdevol_Click()
frmOPExtDevoluc.Show 1
End Sub

Private Sub mnuextIngSalidas_Click()
    frmOpeExtornosKardex.Show 1
End Sub

Private Sub mnuextRemesa_Click()
    frmOpExtornosRemesas.Show 1
End Sub

Private Sub mnuHabDev_Click()
    frmOpeExtornosHabDev.Show 1
End Sub

Private Sub mnuImpSoliAfi_Click()
    frmImpSolAfi.Show 1
End Sub

Private Sub mnuIngSalidas_Click()

End Sub

Private Sub mnulimrementrans_Click()
frmManLimRemEnTran.Show 1

End Sub

Private Sub mnulistmovbgen_Click()
    frmRepMovBGEN.Show 1
    
End Sub

Private Sub mnumanTipoCta_Click()
    frmMantTipoCuenta.Show 1
End Sub

Private Sub mnumantpermisos_Click()
frmMantPermisos.Show 1

End Sub

Private Sub mnumanttarifario_Click()
    frmMantTarifario.Show 1
End Sub

Private Sub mnumonmaxretdia_Click()
    'frmMantMontoMaxRetXDia.Show 1
    frmLimMontoMaxXDia.Show 1
End Sub

Private Sub mnumonstock_Click()
        frmStockControlAge.Show 1
End Sub

Private Sub mnumontomaxderet_Click()
    frmMantMontoMaDeRetiro.Show 1
End Sub

Private Sub mnunumretpordia_Click()
    frmMantNroRetXDia.Show 1
End Sub


Private Sub mnuOpeConfirmadasDolares_Click()
    frmRepOpeConf.Ini
    Set frmRepOpeConf = Nothing
End Sub

Private Sub mnuOpeConfirmadasSoles_Click()
    frmRepOpeConf.Ini
    Set frmRepOpeConf = Nothing
End Sub

Private Sub mnuOperacionesConciliadasDolares_Click()
    frmRepoOpeNoConcil.Inic
    Set frmRepoOpeNoConcil = Nothing
End Sub

Private Sub mnuOperacionesConciliadasSoles_Click()
    frmRepoOpeNoConcil.Inic
    Set frmRepoOpeNoConcil = Nothing
End Sub

Private Sub mnuOperacionesConfirmadasDolares_Click()
    frmRepoOpRealizadasRet.Ini 2
    Set frmRepoOpRealizadasRet = Nothing
End Sub

Private Sub mnuOpeConfirmadas_Click()
    frmRepOpeConf.Ini
    Set frmRepOpeConf = Nothing
End Sub

Private Sub mnuOperacionesConciliadas_Click()
    frmRepoOpeNoConcil.Inic
    Set frmRepoOpeNoConcil = Nothing
End Sub

Private Sub mnuOperacionesConfirmadas_Click()
    frmRepoOpRealizadasRet.Ini 2
    Set frmRepoOpRealizadasRet = Nothing
End Sub

Private Sub mnuOperacionesConfirmadasnorealizadasDolares_Click()
    frmRepoOpeConfXReg.Ini
    Set frmRepoOpeConfXReg = Nothing
End Sub

Private Sub mnuOperacionesConfirmadasnorealizadasSoles_Click()
    frmRepoOpeConfXReg.Ini
    Set frmRepoOpeConfXReg = Nothing
End Sub

'Private Sub mnuOperacionesConfirmadasSoles_Click()
'    frmRepoOpRealizadasRet.Ini 2, 1
'    Set frmRepoOpRealizadasRet = Nothing
'End Sub


'Private Sub mnuOperacionesPendientesDolares_Click()
'    frmRepoOpRealizadasRet.Ini 3, 2
'    Set frmRepoOpRealizadasRet = Nothing
'End Sub

'Private Sub mnuOperacionesPendientesSoles_Click()
'    frmRepoOpRealizadasRet.Ini 3
'    Set frmRepoOpRealizadasRet = Nothing
'End Sub

Private Sub mnuOperacionesConfirmadasnorealizadas_Click()
    frmRepoOpeConfXReg.Ini
    Set frmRepoOpeConfXReg = Nothing
End Sub

Private Sub mnuOperacionesPendientes_Click()
    frmRepoOpRealizadasRet.Ini 3
    Set frmRepoOpRealizadasRet = Nothing
End Sub

'Private Sub mnuOperacionesRealizadasDolares_Click()
'    frmRepoOpRealizadasRet.Ini 1, 2
'    Set frmRepoOpRealizadasRet = Nothing
'End Sub
'
'Private Sub mnuOperacionesRealizadasSoles_Click()
'    frmRepoOpRealizadasRet.Ini 1, 1
'    Set frmRepoOpRealizadasRet = Nothing
'End Sub

Private Sub mnuOperacionesRealizadas_Click()
        frmRepoOpRealizadasRet.Ini 1
    Set frmRepoOpRealizadasRet = Nothing
End Sub

Private Sub mnuoperlib_Click()
    frmMantNroOpeLibres.Show 1
End Sub





Private Sub mnuPITConciCargaLog_Click()
    frmPITCargaLog.Show vbModal
End Sub

Private Sub mnuPITConciConciliacion_Click()
    frmPITConciliacion.Show vbModal
End Sub

Private Sub mnuPITConsultaMovAut_Click()
    frmPITConsultaMovimientosAut.Show vbModal
End Sub

Private Sub mnuPITConsultaMovCli_Click()
    frmPITConsultaMovimientosCli.Show vbModal
End Sub

Private Sub mnuPITRepOpeInterCMAC_Click()
    frmPITReporteOpeInterCMAC.Show 1
End Sub

'Private Sub mnuOperNoConcDolares_Click()
'    frmRepoOpeNoConcil.Inic 2
'    Set frmRepoOpeNoConcil = Nothing
'End Sub
'
'Private Sub mnuOperNoConcSoles_Click()
'    frmRepoOpeNoConcil.Inic 1
'    Set frmRepoOpeNoConcil = Nothing
'End Sub

Private Sub mnuRegStckActBvda_Click()
    frmStockActualBvda.Show 1
End Sub

Private Sub mnuregstockbovgen_Click()
    frmStockRegBovGen.Show 1
End Sub

Private Sub mnuremdevol_Click()
    frmStockDevol.Show 1
End Sub

Private Sub mnurementransito_Click()
    Call ListadoDERemesasENTransito
End Sub

Private Sub MnuRepBovedaAgenUsuarios_Click()
     frmRepMovimientoBovAgencia.Show 1
End Sub

Private Sub mnurepcontrol_Click()
    frmRangoFechas.Show 1
End Sub

Private Sub mnurepgenlisstockcajero_Click()
    Call ListadoStocksPorReporte
End Sub

Private Sub mnureplimrementra_Click()
    ListadoDERemesasENTransitoFueraDELimite
End Sub

Private Sub mnurepopeconf_Click()
    frmRepOpeConf.Show 1
End Sub

Private Sub mnureptarjact_Click()
    frmRepTarjactiva.Show 1
End Sub

Private Sub mnureptarjcancel_Click()
    frmRepTarjCancel.Show 1
    
End Sub

Private Sub mnureptarjemiporage_Click()
 frmRepTarjEmiPorAge.Show 1
End Sub

Private Sub mnuRepTarjetasRetenidas_Click()
    frmRepTarjetasRetenidas.Inicios
    Set frmRepTarjetasRetenidas = Nothing
End Sub

Private Sub mnureptarjord_Click()
    frmRepTarjOrdenadas.Show 1
End Sub

Private Sub mnureptarjrech_Click()
    frmRepTarRechaz.Show 1
End Sub

Private Sub mnureptarreti_Click()
    Call ListadoDETarjetasRetiradas
End Sub

Private Sub mnuResumenOp_Click()
    frmRepResumenOp.Ini
    Set frmRepResumenOp = Nothing
End Sub

Private Sub mnusaldoStockBovGen_Click()
    frmStockSaldoBovGen.Show 1
End Sub

Private Sub mnuSalir_Click()
    End
End Sub


Private Sub mnustockconfdevol_Click()
    frmStockConfDev.Show 1
End Sub

Private Sub mnuStockRepTarPorAge_Click()
    frmRepTarjPorAge.Show 1
End Sub

Private Sub mnutareptar_Click()
frmMantTarComRep.Show 1

End Sub

Private Sub mnutarjafil_Click()
    frmReportes.Show 1
End Sub

Private Sub mnutarjbloq_Click()
    frmRepTarjetasBloq.Show 1
    
End Sub

Private Sub mnutarjctaelim_Click()
    frmTarjCtaElim.Show 1
End Sub

Private Sub mnutarjctareg_Click()
    frmAdicTarjetaCta.Show 1
End Sub

Private Sub mnutarrechazadas_Click()
    frmAfilCargaTarRechazadas.Show 1
End Sub

Private Sub mnutrajctacons_Click()
    frmTarjCtaCons.Show 1
End Sub
Private Sub mnuSalTar_Click()
    frmStockSalida.Show 1
End Sub

Private Sub mnuStckRemesas_Click()
frmStockRemesas.Show 1
End Sub

Private Sub mnuStockActual_Click()
    frmStockRegTar.Show 1
End Sub
Private Sub mnuIngTar_Click()
    frmStockIngreso.Show 1
End Sub

Private Sub mnuHabilitarTarjeta_Click()
    frmStockHabTarjeta.Show 1
End Sub

Private Sub mnuDevolverTarjetas_Click()
    frmStockDevTarjeta.Show 1
End Sub

Private Sub mnuHabilitarTarjetas_Click()
    frmStockHabTarjeta.Show 1
End Sub
Private Sub mnuCuadrarTarjetas_Click()
    frmStockCuadreTarjeta.Show 1
End Sub
