VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredRelaCta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personas del Credito"
   ClientHeight    =   3780
   ClientLeft      =   3795
   ClientTop       =   3690
   ClientWidth     =   8220
   Icon            =   "frmCredRelaCta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
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
      Left            =   7005
      TabIndex        =   8
      Top             =   3345
      Width           =   1020
   End
   Begin SICMACT.ActXCodCta ActXCtaCod 
      Height          =   420
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   741
      Texto           =   "Crédito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   135
      TabIndex        =   9
      Top             =   600
      Width           =   7860
      Begin VB.ComboBox cboRelacion 
         Height          =   315
         Left            =   5055
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Relacion del Cliente con el Credito"
         Top             =   600
         Width           =   2505
      End
      Begin VB.Frame fractles 
         Height          =   1275
         Left            =   6480
         TabIndex        =   10
         Top             =   1215
         Width           =   1200
         Begin VB.CommandButton cmdcancelar 
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
            Height          =   345
            Left            =   105
            TabIndex        =   5
            Top             =   510
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "A&gregar"
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
            TabIndex        =   1
            Top             =   135
            Width           =   1005
         End
         Begin VB.CommandButton cmdeliminar 
            Caption         =   "Eli&minar"
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
            TabIndex        =   3
            Top             =   855
            Width           =   1005
         End
         Begin VB.CommandButton cmdeditar 
            Caption         =   "&Editar"
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
            TabIndex        =   2
            Top             =   510
            Width           =   1005
         End
      End
      Begin MSComctlLib.ListView ListaRelacion 
         Height          =   1305
         Left            =   225
         TabIndex        =   6
         Top             =   1230
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   2302
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre de Cliente"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Relación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "cCodPers"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ValorRelac"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Estado"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblCodigo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   870
         TabIndex        =   19
         Tag             =   "txtcodigo"
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2145
         TabIndex        =   18
         Tag             =   "txtNombre"
         Top             =   255
         Width           =   5415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Relación :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4305
         TabIndex        =   17
         Top             =   645
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblDocTrib 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   15
         Tag             =   "txttributario"
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "RUC :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2385
         TabIndex        =   14
         Top             =   645
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DNI :"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   660
         Width           =   375
      End
      Begin VB.Label lblDocnat 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   870
         TabIndex        =   12
         Tag             =   "txtdocumento"
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Relaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   975
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Gra&bar"
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
      Left            =   5985
      TabIndex        =   7
      Top             =   3345
      Width           =   1020
   End
End
Attribute VB_Name = "frmCredRelaCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private oCreditoRelac As UCredRelac_Cli
Private cmdEjecutar As Integer

'àgregado por vapi SEGÙN ERS TI-ERS001-2017
Dim bPresol As Boolean
Dim cPersCodPreSol As String
'fin agregado por vapi SEGÙN ERS TI-ERS001-2017

Enum TModiffrmCredRelaCta
    InicioSolicitud = 1
    InicioMantenimiento = 2
End Enum

Enum TCredRelaCtaInicio
    InicioRegistroForm = 1
    InicioConsultaForm = 2
End Enum

Private vTipoInicio As TModiffrmCredRelaCta
Private vValorRelacAnt As String
Private nInicio As TCredRelaCtaInicio
Private vbCreditoAprobado As Boolean
Private vnSolicProceso As Integer
Dim bCambios As Boolean
Dim dPersFecNac As Date

'ADD PTI1 22/08/2018 ERS027-2017
Dim estadoVerifica As Integer
Dim sPercod As String
Dim sPernom As String
'END PTI1

Private Function ExisteTitular() As Boolean
Dim bEnc As Boolean
    bEnc = False
    'Call IniciarMatriz
    oCreditoRelac.IniciarMatriz
    Do While Not oCreditoRelac.EOF
        If oCreditoRelac.ObtenerValorRelac = gColRelPersTitular Then
            bEnc = True
            Exit Do
        End If
        oCreditoRelac.siguiente
    Loop
    ExisteTitular = bEnc
End Function

'*****************************************************************************************
'***     Rutina:                ValidaDatos
'***     Descripcion:           Valida que todos los datos sean correctos
'***     Creado por:            NSSE
'***     Maquina:               07SIST_08
'***     Fecha-Tiempo:          01/06/2001 12:00:05 PM
'***     Ultima Modificacion:   Creacion de la funcion
'*****************************************************************************************

Private Function ValidaDatos() As Boolean
Dim i As Integer
'madm 20110303
Dim bEncGarante As Boolean
Dim bEncRepre As Boolean
Dim bEncRepite As Boolean
Dim bEncRepiteRelac As Boolean

    bEncGarante = False
    bEncRepre = False
    bEncRepite = False
    bEncRepiteRelac = False
'end madm

    'Valida que se ingrese Relacion
    ValidaDatos = True
    If cboRelacion.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de Reclacion que tiene la Persona con el Credito", vbInformation, "Aviso"
        ValidaDatos = False
        cboRelacion.SetFocus
        Exit Function
    End If
    
    'Valida que se halla seleccionado una Persona
    If Trim(lblCodigo.Caption) = "" Then
        MsgBox "Nesecita Seleccionar una Persona", vbInformation, "Aviso"
        ValidaDatos = False
        cmdNuevo.SetFocus
        Exit Function
    End If
    
    'Valida que se Ingrese Primero el Titular
    If oCreditoRelac.NroRelaciones = 0 Then
        If CInt(Trim(Right(cboRelacion.Text, 3))) <> gColRelPersTitular Then
            MsgBox "Debera Ingresar primero el Titular de la Cuenta", vbInformation, "Aviso"
            ValidaDatos = False
            cboRelacion.SetFocus
            Exit Function
        End If
    End If
    
    'Valida que se Ingrese un Solo Titular
    If CInt(Trim(Right(cboRelacion.Text, 5))) = gColRelPersTitular Then
        oCreditoRelac.IniciarMatriz
        Do While Not oCreditoRelac.EOF
            If oCreditoRelac.ObtenerValorRelac = gColRelPersTitular Then
                MsgBox "El Titular de la Cuenta ya ha sido Ingresado", vbInformation, "Aviso"
                ValidaDatos = False
                cboRelacion.SetFocus
                Exit Function
            End If
            oCreditoRelac.siguiente
        Loop
    End If
    
    'INICIO ORCR-20140913*********
    'Valida que se Ingrese un Solo Conyugue
    If CInt(Trim(Right(cboRelacion.Text, 5))) = gColRelPersConyugue Then
        oCreditoRelac.IniciarMatriz
        Do While Not oCreditoRelac.EOF
            If oCreditoRelac.ObtenerValorRelac = gColRelPersConyugue Then
                MsgBox "El Conyugue de la Cuenta ya ha sido Ingresado", vbInformation, "Aviso"
                ValidaDatos = False
                cboRelacion.SetFocus
                Exit Function
            End If
            oCreditoRelac.siguiente
        Loop
    End If
    'FIN ORCR-20140913************
    'Valida que el Titular solo sea Titular y No Otra Relacion a la vez
''    If cmdEjecutar <> 2 Then
''        oCreditoRelac.IniciarMatriz
''        Do While Not oCreditoRelac.EOF
''            If oCreditoRelac.ObtenerCodigo = Trim(lblCodigo.Caption) Then
''                MsgBox "La Persona ya posee una Relacion en el Credito", vbInformation, "Aviso"
''                ValidaDatos = False
''                cboRelacion.SetFocus
''                Exit Function
''            End If
''            oCreditoRelac.siguiente
''        Loop
''    End If
'madm 20110303
    If cmdEjecutar <> 2 Then
        oCreditoRelac.IniciarMatriz
        Do While Not oCreditoRelac.EOF
            If oCreditoRelac.ObtenerCodigo = Trim(lblCodigo.Caption) Then
                bEncRepite = True
                bEncGarante = IIf(oCreditoRelac.ObtenerValorRelac = 25, True, False)
                bEncRepre = IIf(oCreditoRelac.ObtenerValorRelac = 23, True, False)
                    If oCreditoRelac.ObtenerValorRelac = CInt(Trim(Right(Me.cboRelacion.Text, 2))) Then
                        bEncRepiteRelac = True
                    End If
            End If
            oCreditoRelac.siguiente
        Loop
        
        If bEncRepite Then
            If (bEncGarante = False And bEncRepre = False) Or (bEncRepiteRelac) Then
                MsgBox "La Persona ya posee una Relacion en el Credito", vbInformation, "Aviso"
                cboRelacion.SetFocus
                ValidaDatos = False
                Exit Function
            Else
                If bEncRepre Then
                    If Not CInt(Trim(Right(Me.cboRelacion.Text, 2))) = 25 Then
                        MsgBox "La Persona ya posee una Relacion en el Credito", vbInformation, "Aviso"
                        cboRelacion.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
                
                If bEncGarante Then
                    If Not CInt(Trim(Right(Me.cboRelacion.Text, 2))) = 23 Then
                        MsgBox "La Persona ya posee una Relacion en el Credito", vbInformation, "Aviso"
                        cboRelacion.SetFocus
                        ValidaDatos = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
'end madm

    
    'Valida que el Titular Tenga un Conyugue de Diferente Sexo o No Tenga Conyugue si es Persona Juridica
    If CInt(Right(cboRelacion.Text, 2)) = gColRelPersConyugue Then
        Dim nPos As Integer
        
        'Verficamos si existe el titular
        If Not oCreditoRelac.ExisteTitular Then
                MsgBox "No Se puede Agregar al Conyugue porque No Existe el Titular", vbInformation, "Aviso"
                ValidaDatos = False
                cboRelacion.SetFocus
                Exit Function
        End If
        nPos = oCreditoRelac.PosicionTitular
        oCreditoRelac.nPuntMat = nPos
        If Len(Trim(oCreditoRelac.ObtenerValorSexo)) <= 0 Then
            MsgBox "No Se puede Agregar al Conyugue porque No se ha Definido el Sexo del Titular", vbInformation, "Aviso"
            ValidaDatos = False
            cmdcancelar.SetFocus
            Exit Function
        End If
        
        nPos = oCreditoRelac.PosicionTitular
        oCreditoRelac.nPuntMat = nPos
        If Trim(oCreditoRelac.ObtenerValorSexo) = Trim(lblNombre.Tag) Then
            MsgBox "No Se puede Agregar al Conyugue porque Posee el mismo Sexo que el Titular", vbInformation, "Aviso"
            ValidaDatos = False
            cmdcancelar.SetFocus
            Exit Function
        End If
    End If
    
    'JOEP20190117 CP
    'Cuando ingresa el titular en desorden Validad que el Titular y el Conyuge o Codeudor sea del mismo sexo
    If CInt(Right(cboRelacion.Text, 2)) = gColRelPersTitular Or CInt(Right(cboRelacion.Text, 2)) = gColRelPersConyugue Or CInt(Right(cboRelacion.Text, 2)) = gColRelPersCodeudor Then
        oCreditoRelac.IniciarMatriz
        Do While Not oCreditoRelac.EOF
                If oCreditoRelac.ObtenerValorSexo = Trim(lblNombre.Tag) Then
                    MsgBox "No se puede agregar al " & IIf(CInt(Right(cboRelacion.Text, 2)) = 20, "TITULAR", IIf(CInt(Right(cboRelacion.Text, 2)) = 21, "CONYUGE", "CODEUDOR")) & " porque posee el mismo sexo que el " & oCreditoRelac.ObtenerRelac, vbInformation, "Aviso"
                    ValidaDatos = False
                    cboRelacion.SetFocus
                    Exit Function
                End If
                oCreditoRelac.siguiente
            Loop
    End If
    'JOEP20190117 CP
    
    'EJVG20130202 ***
    If tieneGarantiasPendiente(lblCodigo.Caption) Then
        MsgBox "No se puede agregar a la persona " & lblNombre.Caption & Chr(10) & "por contar con Garantías pendiente de Constituir o en Segunda Preferente", vbInformation, "Aviso"
        ValidaDatos = False
        cmdcancelar.SetFocus
        Exit Function
    End If
    'END EJVG *******
    'RECO20150609 ERS033-2015*******************
    If CInt(Trim(Right(cboRelacion.Text, 5))) = gColRelPersCodeudor Then
        oCreditoRelac.IniciarMatriz
        Do While Not oCreditoRelac.EOF
            If oCreditoRelac.ObtenerValorRelac = gColRelPersCodeudor Then
                MsgBox "El codeudor de la Cuenta ya ha sido Ingresado", vbInformation, "Aviso"
                ValidaDatos = False
                cboRelacion.SetFocus
                Exit Function
            End If
            oCreditoRelac.siguiente
        Loop
    End If
    
    Dim oCons As NConstSistemas
    Set oCons = New NConstSistemas
    
    If oCons.LeeConstSistema(507) = 1 Then
        If cmdEjecutar <> 2 Then
            If CInt(Trim(Right(cboRelacion.Text, 5))) = gColRelPersConyugue Or CInt(Trim(Right(cboRelacion.Text, 5))) = gColRelPersCodeudor Then
                oCreditoRelac.IniciarMatriz
                Do While Not oCreditoRelac.EOF
                    If CInt(Trim(Right(cboRelacion.Text, 5))) = gColRelPersConyugue Then
                        If oCreditoRelac.ObtenerValorRelac = gColRelPersCodeudor Then
                            MsgBox "No se puede agregar el cónyuge en la relación del crédito, debido a que ya se agregó previamente un codeudor", vbInformation, "Aviso"
                            ValidaDatos = False
                            cboRelacion.SetFocus
                            Exit Function
                        End If
                    Else
                        If oCreditoRelac.ObtenerValorRelac = gColRelPersConyugue Then
                            MsgBox "No se puede agregar el codeudor en la relación del crédito, debido a que ya se agregó previamente un cónyuge", vbInformation, "Aviso"
                            ValidaDatos = False
                            cboRelacion.SetFocus
                            Exit Function
                        End If
                    End If
                    oCreditoRelac.siguiente
                Loop
            End If
        End If
    End If
    'RECO FIN***********************************
End Function

Private Sub LimpiaPantalla()
    ActXCtaCod.NroCuenta = ""
    ActXCtaCod.CMAC = gsCodCMAC
    ActXCtaCod.Age = gsCodAge
    ListaRelacion.ListItems.Clear
    Set oCreditoRelac = Nothing
    cmdNuevo.Enabled = False
    cmdeditar.Enabled = False
    cmdeliminar.Enabled = False
    cmdAceptar.Enabled = False
End Sub

Private Sub ActualizarRelacion()
    Call oCreditoRelac.ActualizaRelacion(lblCodigo.Caption, Trim(Left(cboRelacion.Text, Len(cboRelacion.Text) - 10)), Trim(Right(cboRelacion.Text, 10)), Trim(Right(vValorRelacAnt, 10)))
    Call CargaObjeto
End Sub

Public Sub ControlPermiso()
    If nInicio = InicioRegistroForm Then
        If vTipoInicio = InicioSolicitud Then
            cmdNuevo.Enabled = True
            cmdeditar.Enabled = True
            cmdeliminar.Enabled = True
            cmdAceptar.Enabled = True
        Else
            cmdNuevo.Enabled = False
            cmdeditar.Enabled = False
            cmdeliminar.Enabled = False
            cmdAceptar.Enabled = False
        End If
    Else
        cmdNuevo.Enabled = False
        cmdeditar.Enabled = False
        cmdeliminar.Enabled = False
        cmdAceptar.Enabled = True
    End If
End Sub

Public Sub Inicio(ByRef oCredPersRela As UCredRelac_Cli, ByVal pInicio As TModiffrmCredRelaCta, Optional ByVal pnInicio As TCredRelaCtaInicio = InicioRegistroForm, Optional ByVal pnSolicProceso As Integer = 2, Optional ByVal pbPresol As Boolean = False, Optional ByVal pcPersCodPreSol As String) ' vapi agrego el paràmetro pbPresol SEGÙN ERS TI-ERS001-2017
    'agregado por vapi SEGÙN ERS TI-ERS001-2017
    bPresol = pbPresol
    cPersCodPreSol = pcPersCodPreSol
    'fin agregado por vapi
    nInicio = pnInicio
    Set oCreditoRelac = oCredPersRela
    vTipoInicio = pInicio
    vnSolicProceso = pnSolicProceso
    If vTipoInicio = InicioSolicitud Then
        ActXCtaCod.top = 660
        ActXCtaCod.Left = 210
        ActXCtaCod.Enabled = False
        ActXCtaCod.Visible = False
        fraCliente.top = 45
        fraCliente.Left = 150
        cmdAceptar.top = 2790
        cmdAceptar.Left = 6000
        cmdSalir.top = 2790
        cmdSalir.Left = 7020
        Me.Height = 3675
    Else
        ActXCtaCod.top = 105
        ActXCtaCod.Left = 150
        ActXCtaCod.Visible = True
        ActXCtaCod.Enabled = True
        ActXCtaCod.EnabledAge = True
        ActXCtaCod.EnabledCMAC = True
        ActXCtaCod.EnabledCta = True
        ActXCtaCod.EnabledProd = True
        ActXCtaCod.CMAC = gsCodCMAC
        ActXCtaCod.Age = gsCodAge
        fraCliente.top = 600
        fraCliente.Left = 135
        cmdAceptar.top = 3345
        cmdAceptar.Left = 5985
        cmdSalir.top = 3345
        cmdSalir.Left = 7005
        Me.Height = 4185
        '***marg ers046-2016***
        gsOpeCod = "190280"
        '***end marg***********
    End If
    Call ControlPermiso
    
      'AGREGADO PTI1 INICIO 22/08/2018 ERS027-2017
      Dim bCredSol As Boolean
      Dim bAmpliacion As Integer
      bCredSol = frmCredSolicitud.getbfCredSolicitud
      bAmpliacion = frmCredSolicitud.getbAmpliacion
      
      If bCredSol Then
        sPercod = lblCodigo.Caption
        sPernom = ""
        If estadoVerifica = 0 And cPersCodPreSol = "" And sPercod <> "" Then
            Dim oHojaRutaVerifica As COMDCredito.DCOMhojaRuta
            Set oHojaRutaVerifica = New COMDCredito.DCOMhojaRuta
            Dim datosUsuario As ADODB.Recordset
            Set datosUsuario = Nothing
            Set datosUsuario = oHojaRutaVerifica.existePresolicitud(sPercod, gsCodUser, bAmpliacion)
            If datosUsuario.RecordCount > 0 Then
                estadoVerifica = 1
                sPernom = datosUsuario!cPersNombre
                MsgBox "El cliente tiene una pre solicitud pendiente, por favor gestionarlo por el modulo de Pre Solicitud", vbInformation, "Aviso"
                frmCredSolicitud.cmdpresolicitud_Click
                bPresol = True
                cPersCodPreSol = frmCredSolicitud.getcPersCodPreSol
                cmdCancelar_Click
                Unload Me
            Else
                estadoVerifica = 0
                Me.Show 1
            End If
        Else
            Unload Me
            estadoVerifica = 0
            If bPresol Then
                cmdAceptar_Click
            End If
        End If
      Else
        If bPresol Then
            cmdAceptar_Click
        Else
            Me.Show 1
        End If
      End If
    
    sPercod = ""
    sPernom = ""
    'FIN AGREGADO PTI1
              
'    'àgregado por vapi SEGÙN ERS TI-ERS001-2017 'COMENTADO POR PTI1 22/08/2018 ERS027-2017
'    If bPresol Then
'        cmdAceptar_Click
'    Else
'        Me.Show 1
'    End If
'    'fin agregado por vapi 'FIN COMENTADO PTI1
End Sub

Public Sub RelacionesCredito(ByRef oCredRela As UCredRelac_Cli)
    Set oCreditoRelac = oCredRela
    Me.Show 1
End Sub

Private Sub AdicionaRelacion()
Dim L As ListItem
'Dim oPers As COMDPersona.UCOMPersona

Dim oCreditos As COMDCredito.DCOMCreditos
'Dim oCon As COMConecta.DCOMConecta
'Dim ssql As String
'Dim R As ADODB.Recordset
'Dim sCtaCod As String
Dim MatRelaciones As Variant
Dim i As Integer
Dim bTieneRelaciones As Boolean
Dim lsPersNombreGarantPend As String

    'Se Adiciona el Titular
    Call oCreditoRelac.AdicionaRelacion(lblCodigo.Caption, lblNombre.Caption, Trim(Left(cboRelacion.Text, 40)), Trim(Right(cboRelacion.Text, 5)), Trim(Right(cboRelacion.Text, 5)), lblDocnat.Caption, lblDocTrib.Caption, NuevaRegistro, CInt(Trim(Right(cboRelacion.Text, 2))), lblNombre.Tag, CInt(lblCodigo.Tag), dPersFecNac)
    Set L = ListaRelacion.ListItems.Add(, , lblNombre.Caption)
    L.SubItems(1) = Trim(Left(cboRelacion.Text, 40))
    L.SubItems(2) = lblCodigo.Caption
    L.SubItems(3) = Trim(Right(cboRelacion.Text, 5))
    
    lsPersNombreGarantPend = ""
    'Si es Titular Adicionamos el conyugue y las otras personas relacionadas al Credito
    If Trim(Right(cboRelacion.Text, 5)) = gColRelPersTitular Then
        Set oCreditos = New COMDCredito.DCOMCreditos
        MatRelaciones = oCreditos.AdicionaRelacionesCredito(lblCodigo.Caption, bTieneRelaciones)
        Set oCreditos = Nothing
        
        If bTieneRelaciones = False Then Exit Sub
        
        For i = 0 To UBound(MatRelaciones) - 1
            If Not tieneGarantiasPendiente(MatRelaciones(i, 0)) Then 'EJVG20130204
            Call oCreditoRelac.AdicionaRelacion(MatRelaciones(i, 0), PstaNombre(CStr(MatRelaciones(i, 1))), MatRelaciones(i, 2), MatRelaciones(i, 3), MatRelaciones(i, 4), MatRelaciones(i, 5), MatRelaciones(i, 6), MatRelaciones(i, 7), MatRelaciones(i, 8), MatRelaciones(i, 9), MatRelaciones(i, 10), MatRelaciones(i, 11))
            Set L = ListaRelacion.ListItems.Add(, , PstaNombre(CStr(MatRelaciones(i, 1))))
            L.SubItems(1) = MatRelaciones(i, 2)
            L.SubItems(2) = MatRelaciones(i, 0)
            L.SubItems(3) = MatRelaciones(i, 3)
            Else
                lsPersNombreGarantPend = lsPersNombreGarantPend & Chr(10) & "- " & Trim(MatRelaciones(i, 1))
            End If
        Next
        'EJVG20130204 ***
        If lsPersNombreGarantPend <> "" Then
            MsgBox "Las sgtes personas que tienen relación con el titular no se pueden agregar" & Chr(10) & "por contar con Garantías pendiente de Constituir o en Segunda Preferente:" & Chr(10) & lsPersNombreGarantPend, vbInformation, "Aviso"
        End If
        'END EJVG *******
    End If
    '    Set oPers = New COMDPersona.UCOMPersona
    '    Call oPers.ObtieneConyugeDePersona(lblCodigo.Caption)
    '    If oPers.sPersCod <> "" Then
    '        Call oCreditoRelac.AdicionaRelacion(oPers.sPersCod, PstaNombre(oPers.sPersNombre), "Conyuge", Trim(Str(gColRelPersConyugue)), Trim(Str(gColRelPersConyugue)), oPers.sPersIdnroDNI, oPers.sPersIdnroRUC, NuevaRegistro, 1, oPers.sPersNatSexo, CInt(oPers.sPersPersoneria), oPers.dPersNacCreac)
    '        Set L = ListaRelacion.ListItems.Add(, , PstaNombre(oPers.sPersNombre))
    '        L.SubItems(1) = "Conyuge"
    '        L.SubItems(2) = oPers.sPersCod
    '        L.SubItems(3) = Trim(Str(gColRelPersConyugue))
    '    End If
    '    Set oPers = Nothing
        
         'Obtenemos Otras Relaciones de Creditos Anteriores
'        Set oCon = New COMConecta.DCOMConecta
'        oCon.AbreConexion
'
'        ssql = "Select top 1 C.dVigencia, PP.* "
'        ssql = ssql & " from ProductoPersona PP"
'        ssql = ssql & " Inner Join Colocaciones C ON C.cCtaCod = PP.cCtaCod"
'        ssql = ssql & " Inner Join ColocacCred CC ON CC.cCtaCod = PP.cCtaCod"
'        ssql = ssql & " Inner Join Producto P ON P.cCtaCod = PP.cCtaCod"
'        ssql = ssql & " where PP.cPersCod = '" & lblCodigo.Caption & "' AND P.nPrdEstado in (2020,2021,2022,2030,2031,2032)"
'        ssql = ssql & " Order by C.dVigencia DESC"
'
'        Set R = oCon.CargaRecordSet(ssql)
'
'        If R.RecordCount > 0 Then
'            sCtaCod = R!cCtaCod
'            R.Close
'
'            ssql = " select PP.nPrdPersRelac,PP.cPersCod,C.cConsDescripcion, C.nConsValor"
'            ssql = ssql & " from ProductoPersona PP"
'            ssql = ssql & " Inner Join Constante C ON C.nConsValor = PP.nPrdPersRelac AND C.nConsCod = 3002"
'            ssql = ssql & " Where cCtaCod = '" & sCtaCod & "' and nPrdPersRelac not in (20,28,29,21)"
'
'            Set R = oCon.CargaRecordSet(ssql)
     '       Set oCreditos = New COMDCredito.DCOMCreditos
            
     '       Set R = oCreditos.ObtenerOtrasRelaciones(lblCodigo.Caption)
     '       Set oCreditos = Nothing
            
     '       Do While Not R.EOF
     '           Set oPers = New COMDPersona.UCOMPersona
     '           Call oPers.ObtieneClientexCodigo(R!cPersCod)
     '           If oPers.sPersCod <> "" Then
     '               Call oCreditoRelac.AdicionaRelacion(oPers.sPersCod, PstaNombre(oPers.sPersNombre), Trim(R!cConsDescripcion), Trim(Str(R!nPrdPersRelac)), Trim(Str(R!nPrdPersRelac)), oPers.sPersIdnroDNI, oPers.sPersIdnroRUC, NuevaRegistro, 1, oPers.sPersNatSexo, CInt(oPers.sPersPersoneria), oPers.dPersNacCreac)
     '               Set L = ListaRelacion.ListItems.Add(, , PstaNombre(oPers.sPersNombre))
     '               L.SubItems(1) = R!cConsDescripcion
     '               L.SubItems(2) = oPers.sPersCod
     '               L.SubItems(3) = Trim(Str(R!nPrdPersRelac))
     '           End If
     '           Set oPers = Nothing
                
     '           R.MoveNext
     '       Loop
     '       R.Close
            
     '   End If
                
        'oCon.CierraConexion
        'Set oCon = Nothing
        
   ' End If
End Sub

Private Sub CargaControlRelaCred()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaControlRelaCred
    
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(gColocRelacPers)
    Do While Not R.EOF
        If R!nConsValor >= 20 And R!nConsValor <= 25 Then
            cboRelacion.AddItem Trim(R!cConsDescripcion) & Space(80) & R!nConsValor
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing

    Exit Sub

ErrorCargaControlRelaCred:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub CargaPersona(ByVal pPersona As COMDPersona.UCOMPersona)
    
    If Not pPersona Is Nothing Then
        lblCodigo.Caption = Trim(pPersona.sPersCod)
        lblNombre.Caption = Trim(pPersona.sPersNombre)
        lblDocnat.Caption = Trim(pPersona.sPersIdnroDNI)
        lblDocTrib.Caption = Trim(pPersona.sPersIdnroRUC)
        lblNombre.Tag = Trim(pPersona.sPersNatSexo)
        lblCodigo.Tag = Trim(pPersona.sPersPersoneria)
        dPersFecNac = pPersona.dPersNacCreac
    End If
End Sub

Private Sub CargaObjeto()
Dim L As ListItem
    If oCreditoRelac.NroRelaciones >= 0 Then
        oCreditoRelac.IniciarMatriz
        If oCreditoRelac.ObtenerValorEstado >= gColocEstAprob Then
            vbCreditoAprobado = True
        Else
            vbCreditoAprobado = False
        End If
    End If
    ListaRelacion.ListItems.Clear
    oCreditoRelac.IniciarMatriz
    Do While Not oCreditoRelac.EOF
        If Not (oCreditoRelac.ObtenerValorRelac = gColRelPersEstudioJuridico Or _
                oCreditoRelac.ObtenerValorRelac = gColRelPersAbogadoReponsable Or _
                oCreditoRelac.ObtenerValorRelac = gColRelPersJuzgado Or _
                oCreditoRelac.ObtenerValorRelac = gColRelPersJuez Or _
                oCreditoRelac.ObtenerValorRelac = gColRelPersSecretario) Then
            
            Set L = ListaRelacion.ListItems.Add(, , oCreditoRelac.ObtenerNombre)
            L.SubItems(1) = oCreditoRelac.ObtenerRelac
            L.SubItems(2) = oCreditoRelac.ObtenerCodigo
            L.SubItems(3) = oCreditoRelac.ObtenerValorRelac
            L.SubItems(4) = Trim(str(oCreditoRelac.ObtenerValorEstado))
        
        End If
        oCreditoRelac.siguiente
    Loop
End Sub

Private Sub ActXCtaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set oCreditoRelac = Nothing
        Set oCreditoRelac = New UCredRelac_Cli
        Call oCreditoRelac.CargaRelacPersCred(ActXCtaCod.NroCuenta)
        Call CargaObjeto
        If oCreditoRelac.NroRelaciones = 0 Then
            MsgBox "Cuenta No Existe o Credito ya Esta Vigente", vbInformation, "Aviso"
        Else
            ActXCtaCod.Enabled = False
            If vTipoInicio = InicioMantenimiento Then
                If nInicio = InicioConsultaForm Then
                    cmdNuevo.Enabled = False
                    cmdeditar.Enabled = False
                    cmdeliminar.Enabled = False
                    cmdAceptar.Enabled = False
                Else
                    cmdNuevo.Enabled = True
                    cmdeditar.Enabled = True
                    cmdeliminar.Enabled = True
                    cmdAceptar.Enabled = True
                    If gsCodCargo <> "002009" Then
                        If vbCreditoAprobado Then
                            cmdeliminar.Enabled = False
                            cmdeditar.Enabled = False
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboRelacion_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        cmdNuevo.SetFocus
     End If
End Sub

Private Sub cmdAceptar_Click()
Dim oCredito As COMDCredito.DCOMCredito
    If vTipoInicio = InicioMantenimiento Then
        If Not ExisteTitular Then
            MsgBox "Debe Ingresar el Titular del Credito", vbInformation, "Aviso"
            Exit Sub
        End If
        Call existeAlertasRelaCta(oCreditoRelac.ObtenerMatrizRelaciones) 'FRHU 20160702 ERS002-2016
        If MsgBox("Se va a Grabar la Informacion, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            bCambios = False
            Set oCredito = New COMDCredito.DCOMCredito
            Call oCredito.ActuializaPersRelacCred(oCreditoRelac.ObtenerMatrizRelaciones, ActXCtaCod.NroCuenta, True)
            Set oCredito = Nothing
            ActXCtaCod.Enabled = True
            Call LimpiaPantalla
        End If
    Else
        Unload Me
    End If
Set oCredito = Nothing
End Sub

Private Sub cmdCancelar_Click()
    lblCodigo.Caption = ""
    lblNombre.Caption = ""
    lblCodigo.Caption = ""
    lblDocTrib.Caption = ""
    lblDocnat.Caption = ""
    cmdNuevo.Caption = "A&gregar"
    cmdeditar.Visible = True
    If ListaRelacion.ListItems.count = 0 Then
        cmdeditar.Enabled = False
    Else
        cmdeditar.Enabled = True
    End If
    cmdcancelar.Visible = False
    cmdeliminar.Enabled = True
    cboRelacion.ListIndex = -1
    cboRelacion.Enabled = False
    cmdAceptar.Enabled = True
    cmdSalir.Enabled = True
    ListaRelacion.Enabled = True
    If vbCreditoAprobado Then
        cmdeliminar.Enabled = False
        cmdeditar.Enabled = False
    End If
End Sub

Private Sub cmdEditar_Click()
Dim nPos As Integer

    'Verifica que no sea Titular
    If vTipoInicio <> InicioSolicitud Then
        If Trim(ListaRelacion.SelectedItem.SubItems(4)) <> "" Then
            If CInt(ListaRelacion.SelectedItem.SubItems(4)) <> gColocEstSolic And CInt(ListaRelacion.SelectedItem.SubItems(3)) = gColRelPersTitular Then
                MsgBox "No se Puede Modificar al Titular", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If

    cmdEjecutar = 2
    nPos = ListaRelacion.SelectedItem.Index - 1
    oCreditoRelac.nPuntMat = nPos
    lblCodigo.Caption = oCreditoRelac.ObtenerCodigo
    lblNombre.Caption = oCreditoRelac.ObtenerNombre
    lblNombre.Tag = oCreditoRelac.ObtenerValorSexo
    lblDocnat.Caption = oCreditoRelac.ObtenerDNI
    lblDocTrib.Caption = oCreditoRelac.ObtenerRUC
    cboRelacion.ListIndex = IndiceListaCombo(cboRelacion, oCreditoRelac.ObtenerValorRelac)
    cboRelacion.Enabled = True
    cboRelacion.SetFocus
    vValorRelacAnt = cboRelacion.Text
    ListaRelacion.Enabled = False
    cmdNuevo.Caption = "A&ceptar"
    cmdeditar.Visible = False
    cmdcancelar.Visible = True
    cmdeliminar.Enabled = False
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = False
End Sub



Private Sub CmdEliminar_Click()
    Dim oCredito As COMDCredito.DCOMCredito
    Dim lsMsg As String
    'Verifica que no sea Titular
    'If nInicio = InicioRegistroForm And CInt(ListaRelacion.SelectedItem.SubItems(3)) = gColRelPersTitular Then
    '    MsgBox "No se Puede Modificar al Titular", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    
    'EJVG20160607 ***
    If vTipoInicio = InicioMantenimiento Then
        Set oCredito = New COMDCredito.DCOMCredito
        lsMsg = oCredito.CadenaPermiteModificarInterviniente(ActXCtaCod.NroCuenta)
        If Len(lsMsg) > 0 Then
            MsgBox lsMsg, vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    'END EJVG *******
    
    If ListaRelacion.ListItems.count > 0 Then
        'If MsgBox("Se Va ha Eliminar el Registro?, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Call oCreditoRelac.EliminarRelacion(ListaRelacion.SelectedItem.SubItems(2), Trim(ListaRelacion.SelectedItem.SubItems(3)))
            Call ListaRelacion.ListItems.Remove(ListaRelacion.SelectedItem.Index)
            bCambios = True
        'End If
    Else
        MsgBox "No Existe Registros para Eliminar", vbInformation, "Aviso"
    End If
    Set oCredito = Nothing
End Sub


Private Sub cmdNuevo_Click()
   
   
   
   If cmdNuevo.Caption = "A&gregar" Then
      cmdNuevo.Caption = "A&ceptar"
      cmdeditar.Visible = False
      cmdcancelar.Visible = True
      cmdeliminar.Enabled = False
      
      'Call CargaPersona(frmBuscaPersona.Inicio) comentado por vapi
      
      'AGREGADO POR VAPI SEGÙN ERS TI-ERS001-2017
      
      If bPresol Then
        Call CargaPersona(frmBuscaPersona.InicioAutomatico(cPersCodPreSol)) 'vapi SEGÙN ERS TI-ERS001-2017
      Else
        Call CargaPersona(frmBuscaPersona.Inicio)
      End If
      
      'fin agregado por vapi
      
      cmdEjecutar = 1
      cboRelacion.ListIndex = 0
      cboRelacion.Enabled = True
      If cboRelacion.Visible And cboRelacion.Enabled Then
        cboRelacion.SetFocus
      End If
      cmdAceptar.Enabled = False
      cmdSalir.Enabled = False
      cboRelacion.Enabled = True
   Else
        If ValidaDatos Then
            bCambios = True
            If cmdEjecutar = 1 Then
                Call AdicionaRelacion
                Call cmdCancelar_Click
            Else
                Call ActualizarRelacion
                Call cmdCancelar_Click
            End If
            ListaRelacion.Enabled = True
            cmdeditar.Visible = True
            cmdcancelar.Visible = False
            cmdeliminar.Enabled = True
            cmdAceptar.Enabled = True
            cmdSalir.Enabled = True
            cboRelacion.Enabled = False
            
            If vbCreditoAprobado Then
                cmdeliminar.Enabled = False
                cmdeditar.Enabled = False
            End If
        Else
            cboRelacion.Enabled = True
        End If
        'cmdAceptar.SetFocus
   End If
   
   
   'agregado por vapi SEGÙN ERS TI-ERS001-2017
   If bPresol Then
        If ValidaDatos Then
            bCambios = True
            If cmdEjecutar = 1 Then
                Call AdicionaRelacion
                Call cmdCancelar_Click
            Else
                Call ActualizarRelacion
                Call cmdCancelar_Click
            End If
            ListaRelacion.Enabled = True
            cmdeditar.Visible = True
            cmdcancelar.Visible = False
            cmdeliminar.Enabled = True
            cmdAceptar.Enabled = True
            cmdSalir.Enabled = True
            cboRelacion.Enabled = False
            
            If vbCreditoAprobado Then
                cmdeliminar.Enabled = False
                cmdeditar.Enabled = False
            End If
        Else
            cboRelacion.Enabled = True
        End If
   End If
   'fin agregado por vapi
End Sub

Private Sub cmdsalir_Click()
    If bCambios Then
        If MsgBox("Se han realizado cambios, Desea Salir sin Grabar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraForm Me
   Call CargaControlRelaCred
   Call CargaObjeto
   cboRelacion.Enabled = False
   cmdEjecutar = -1
   If vTipoInicio = InicioSolicitud And vnSolicProceso = 1 Then
        Call cmdNuevo_Click
   End If
End Sub

'Eliminar la referencia al objeto
Private Sub Form_Unload(Cancel As Integer)
    Set oCreditoRelac = Nothing
    '***marg ers046-2016***
    gsOpeCod = ""
    '***end marg***********
End Sub
'EJVG20130202 *** Valida no tenga garantías Pendientes de Constituir o en 2° Preferente
Private Function tieneGarantiasPendiente(ByVal psPersCod As String) As Boolean
    Dim oGarantia As New COMDCredito.DCOMGarantia
    Dim rsGarantia As New ADODB.Recordset

    Set rsGarantia = oGarantia.RecuperaGarantPendConstituirySgdoPreferente(psPersCod)
    If rsGarantia.RecordCount > 0 Then
        tieneGarantiasPendiente = True
    Else
        tieneGarantiasPendiente = False
    End If
  
    Set rsGarantia = Nothing
    Set oGarantia = Nothing
End Function
'END EJVG *******
'FRHU 20160702 ERS002-2016
Private Sub existeAlertasRelaCta(ByVal poRelPersCred As Variant)
    Dim i As Integer
    Dim lsPersCodTitular As String
    Dim lsPersCodConyuCodeu As String
    
    For i = 0 To UBound(poRelPersCred) - 1
        If poRelPersCred(i, 1) = gColRelPersTitular Then
            lsPersCodTitular = poRelPersCred(i, 0)
        End If
        If poRelPersCred(i, 1) = gColRelPersConyugue Then
            lsPersCodConyuCodeu = poRelPersCred(i, 0)
            If verificarAlertas(lsPersCodTitular, lsPersCodConyuCodeu) Then Exit Sub
        End If
        If poRelPersCred(i, 1) = gColRelPersCodeudor Then
            lsPersCodConyuCodeu = poRelPersCred(i, 0)
            If verificarAlertas(lsPersCodTitular, lsPersCodConyuCodeu) Then Exit Sub
        End If
    Next
End Sub
Private Function verificarAlertas(ByVal psPersCodTitular As String, ByVal psPersCodConyuCodeu As String) As Boolean
    Dim oCred As New COMDCredito.DCOMCredito
    Dim rs As New ADODB.Recordset
    
    verificarAlertas = False
    If psPersCodTitular <> "" Then
        Set rs = oCred.verificarAlertasRelaCred(psPersCodTitular, psPersCodConyuCodeu)
        If Not (rs.BOF And rs.EOF) Then
            If IIf(IsNull(rs!nAlerta), 0, rs!nAlerta) <> 0 Then
                MsgBox IIf(IsNull(rs!cTpoDesc), "", rs!cTpoDesc), vbInformation, "AVISO"
                verificarAlertas = True
            End If
        End If
    End If
End Function
'FIN FRHU 20160702


'ADD PTI1 23/08/2018 ERS027-2017
Property Get getnEstadoVerifica() As String
getnEstadoVerifica = estadoVerifica
End Property

Property Get getsPersona() As String
getsPersona = sPernom
End Property

Public Sub ActualizaEstadoVerifica(nestadoverifica As Integer)
estadoVerifica = nestadoverifica
End Sub
'FIN PTI1

