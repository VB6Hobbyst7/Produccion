VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredNewNivAprPorNivel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización por Niveles"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   Icon            =   "frmCredNewNivAprPorNivel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistorial 
      Caption         =   "Historial"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   3840
      Width           =   1410
   End
   Begin VB.CommandButton cmdResolvCred 
      Caption         =   "Resolver Autorización"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1965
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
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
      Left            =   8760
      TabIndex        =   0
      Top             =   3840
      Width           =   1170
   End
   Begin MSComctlLib.ListView LstCreditos 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   531
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Nº Credito"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Titular"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Moneda"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Monto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fecha Solicitud"
         Object.Width           =   2469
      EndProperty
   End
End
Attribute VB_Name = "frmCredNewNivAprPorNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprPorNivel
'** Descripción : Formulario para la Aprobación por Niveles creado segun RFC110-2012
'** Creación : JUEZ, 20121206 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim rs As ADODB.Recordset
Dim fsNivAprCod As String
Dim fnTipo As Integer

'RECO20160625 ERS002-2016************************
Enum TipoAprobacion
    nAprobacionCred = 1000
    nAprobacionAuto = 2000
End Enum

Dim nTipoAprobacion As TipoAprobacion
'RECO FIN ***************************************


Public Sub Inicio(Optional ByVal pnTipoOpe As TipoAprobacion = TipoAprobacion.nAprobacionCred)   'RECO20160625 ERS002-2016 AGREGO PARAMETRO TIPO OPE

    nTipoAprobacion = pnTipoOpe 'RECO20160625 ERS002-2016
    '*** FRHU 20160824
    If nTipoAprobacion = nAprobacionAuto Then
        Me.Caption = "Autorización por Niveles"
        cmdResolvCred.Caption = "Resolver Autorización"
    Else
        Me.Caption = "Aprobación por Niveles"
        cmdResolvCred.Caption = "Resolver Crédito"
    End If
    '*** FIN FRHU 20160824
    cmdHistorial.Visible = IIf(nTipoAprobacion = nAprobacionAuto, False, True)
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    If oDNiv.VerificaUsuarioSiTieneNivel(gsCodUser, , , fsNivAprCod) Then
        fnTipo = 1
        ListaCreditosPorAprobar (1)
        If LstCreditos.ListItems.Count > 0 Then
            Me.Show 1
        Else
            Unload Me
        End If
    Else
        If oDNiv.VerificaUsuarioSiTieneNivelDelegado(gsCodUser) Then
            fsNivAprCod = oDNiv.ObtieneNivelAprobacionDelegacion(gsCodUser)
            fnTipo = 2
            ListaCreditosPorAprobar (2)
            If LstCreditos.ListItems.Count > 0 Then
                Me.Show 1
            Else
                Unload Me
            End If
        Else
            MsgBox "Ud. no tiene nivel de aprobación, no puede acceder a esta opción", vbInformation, "Aviso"
            Unload Me
        End If
    End If
End Sub

Private Sub ListaCreditosPorAprobar(ByVal pnTipo As Integer)
    Dim i As Integer, lnFila As Integer
    Dim L As ListItem
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    If nTipoAprobacion = TipoAprobacion.nAprobacionCred Then
        Set rs = oDNiv.RecuperaCreditosPorAprobar(gsCodUser, pnTipo)
    ElseIf nTipoAprobacion = TipoAprobacion.nAprobacionAuto Then
        Set rs = oDNiv.ObtieneAutorizacionesPendientes(gsCodUser)
    End If
    
    Set oDNiv = Nothing
    LstCreditos.ListItems.Clear
    If Not rs.EOF Then
        If Not (rs.EOF And rs.BOF) Then
            LstCreditos.Enabled = True
        End If
        
        Do While Not rs.EOF
            Set L = LstCreditos.ListItems.Add(, , rs.Bookmark)
            L.SubItems(1) = rs!cCtaCod
            L.SubItems(2) = rs!cPersNombre
            L.SubItems(3) = rs!cMoneda
            L.SubItems(4) = Format(rs!nMonto, "#,##0.00")
            L.SubItems(5) = Format(rs!dPrdEstado, "dd/mm/yyyy")
            rs.MoveNext
        Loop
    Else
        MsgBox "No existen créditos pendientes de aprobación para su nivel", vbInformation, "Aviso"
        Unload Me
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub LstCreditos_DblClick()
    cmdResolvCred_Click
End Sub

Private Sub cmdResolvCred_Click()
    Dim bExiste As Boolean
    Dim bVerificarIR As Boolean
    If Me.LstCreditos.SelectedItem.SubItems(1) <> "" Then
        If nTipoAprobacion = TipoAprobacion.nAprobacionCred Then 'RECO20160625 ERS002-2016
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
            bExiste = oDNiv.ExisteRegistroResultadoUserNivApr(LstCreditos.SelectedItem.SubItems(1), fsNivAprCod, gsCodUser)
            
            If bExiste Then
                Set oDNiv = Nothing
                MsgBox "Ya se registraron todas la aprobaciones, favor de verificar en el historial", vbInformation, "Aviso"
            Else
                If Not IntervinientesSonVinculados(LstCreditos.SelectedItem.SubItems(1)) Then
                    Exit Sub
                End If
                
                bVerificarIR = oDNiv.VerificarInformeRiesgoXNivApr(LstCreditos.SelectedItem.SubItems(1))
                Set oDNiv = Nothing
                
                If bVerificarIR Then
                    If Not GenerarDataExposicionRiesgoUnico(LstCreditos.SelectedItem.SubItems(1)) Then
                        Exit Sub
                    End If
                    If Not EmiteInformeRiesgo(eProcesoEmiteInformeRiesgo.NivelAprobacion, LstCreditos.SelectedItem.SubItems(1)) Then
                        Exit Sub
                    End If
                End If
                
                Call frmCredNewNivAprResolvCred.ResolverCredito(LstCreditos.SelectedItem.SubItems(1), fsNivAprCod)
                ListaCreditosPorAprobar (fnTipo)
            End If
        ElseIf nTipoAprobacion = TipoAprobacion.nAprobacionAuto Then 'RECO20160625 ERS002-2016
            frmCredNewNivAutorizaResolver.inicia (LstCreditos.SelectedItem.SubItems(1))
            ListaCreditosPorAprobar (fnTipo)
        End If
    End If
End Sub

Private Sub cmdHistorial_Click()
    If Me.LstCreditos.SelectedItem.SubItems(1) <> "" Then
        Call frmCredNewNivAprHist.InicioCredito(LstCreditos.SelectedItem.SubItems(1))
    End If
End Sub

Private Sub lstCreditos_KeyPress(KeyAscii As Integer)
    Dim n As Integer
    Dim letra As String
    
    If KeyAscii = 13 Then
        EnfocaControl cmdResolvCred
    End If
    
    letra = UCase(Chr(KeyAscii))
    For n = 1 To LstCreditos.ListItems.Count
        If letra = UCase(Left(LstCreditos.ListItems(n).SubItems(2), 1)) Then
            LstCreditos.ListItems(n).Selected = True
            LstCreditos.ListItems(n).EnsureVisible
            Exit For
        End If
    Next
End Sub
