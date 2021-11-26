VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCFNivelApr 
   AutoRedraw      =   -1  'True
   Caption         =   "Carta Fianza - Niveles de Aprobación"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   Icon            =   "frmCFNivelApr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   5760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Editar"
      Height          =   390
      Left            =   180
      TabIndex        =   9
      Top             =   4380
      Width           =   870
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   3780
      TabIndex        =   8
      Top             =   4380
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   4680
      TabIndex        =   4
      Top             =   4380
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   2880
      TabIndex        =   5
      Top             =   4380
      Width           =   870
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   1980
      TabIndex        =   3
      Top             =   4380
      Width           =   870
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   1080
      TabIndex        =   2
      Top             =   4380
      Width           =   870
   End
   Begin VB.Frame frmDetalle 
      Height          =   2115
      Left            =   180
      TabIndex        =   7
      Top             =   2070
      Width           =   5385
      Begin VB.Frame fraNivelAprob 
         Caption         =   "Nivel Aprobacion"
         Height          =   1725
         Left            =   2700
         TabIndex        =   13
         Top             =   270
         Width           =   2445
         Begin VB.CheckBox chkNiv 
            Caption         =   "Nivel IV - Comite Gerenc"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   19
            Tag             =   "4"
            Top             =   1350
            Width           =   2175
         End
         Begin VB.CheckBox chkNiv 
            Caption         =   "Nivel III - Gerencia Cred"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   18
            Tag             =   "3"
            Top             =   990
            Width           =   2175
         End
         Begin VB.CheckBox chkNiv 
            Caption         =   "Nivel II - Jefe Creditos"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   17
            Tag             =   "2"
            Top             =   630
            Width           =   2175
         End
         Begin VB.CheckBox chkNiv 
            Caption         =   "Nivel I - Administrador"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   16
            Tag             =   "1"
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame fraProductos 
         Caption         =   "Productos"
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   2445
         Begin VB.CheckBox chkProd 
            Caption         =   "221 MICROEMPRESA"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   15
            Tag             =   "221"
            Top             =   720
            Width           =   2085
         End
         Begin VB.CheckBox chkProd 
            Caption         =   "121 COMERCIAL"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   14
            Tag             =   "121"
            Top             =   270
            Width           =   2085
         End
      End
   End
   Begin VB.Frame frmAnalista 
      Caption         =   "Analista"
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
      Height          =   2085
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox CboAnalista 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   4995
      End
      Begin MSComctlLib.ListView lvwMetas 
         Height          =   1350
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   2381
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Producto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Monto Inicial"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Monto Final"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label LblAnalista 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   225
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   345
      Left            =   5550
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   609
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCFNivelApr.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCFNivelApr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFNivelApr
'*  CREACION: 25/01/2002      AUTOR :LAYG
'*  MODIFICACION
'***************************************************************************
'*  RESUMEN: Registro de Niveles para aprobación de Cartas Fianza
'***************************************************************************
Option Explicit

Dim lsNueMod As String * 1

Private Sub CboAnalista_Click()
    lblAnalista = Right(cboAnalista.Text, 13)
    Limpiar
    ' Llena Lista de Metas
    LlenaListaNiveles (lblAnalista)
End Sub

Private Sub cboAnalista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CboAnalista_Click
End If
End Sub

Private Sub cmdCancelar_Click()

    Me.frmAnalista.Enabled = True
    Me.frmDetalle.Enabled = False

    Call HabilitarComandos(True, False, False, False, False, True)

End Sub

Private Sub cmdGrabar_Click()
Dim lsSQL As String
Dim lrs As New ADODB.Recordset
Dim lsTipoMeta As String * 1
Dim Item As Integer
Dim lnPropuesto As Currency
Dim lnProd As Integer
Dim lnNiv As Integer
Dim lnRef As Integer
Dim lsCodNivApr As String
Dim loConecta As DConecta

If MsgBox(" Grabar los Niveles de Aprobacion de Carta Fianza ? ", vbQuestion + vbYesNo, "Aviso ") = vbYes Then
    
    Set loConecta = New DConecta
        loConecta.AbreConexion
    
        lsSQL = "DELETE ColocCredPersNivelesApr Where cPersCod = '" & lblAnalista & "' " _
              & "And substring(cCodNiv,2,2) = '21' "
        loConecta.Ejecutar (lsSQL)
   
        For lnProd = 1 To 2
            For lnNiv = 1 To 4
                If Chkprod(lnProd).value = 1 And chkNiv(lnNiv).value = 1 Then
                    lsCodNivApr = Chkprod(lnProd).Tag & "1" & chkNiv(lnNiv).Tag & "1"
                    lsSQL = "INSERT INTO ColocCredPersNivelesApr(cCodNiv,cPersCod) " _
                          & " VALUES('" & lsCodNivApr & "','" & lblAnalista & "') "
                    loConecta.Ejecutar (lsSQL)
                End If
            Next lnNiv
        Next lnProd
    
        loConecta.CierraConexion
    Set loConecta = Nothing
    
    Call HabilitarComandos(True, True, False, False, False, True)
    
    Me.frmAnalista.Enabled = True
    Me.frmDetalle.Enabled = False
    Limpiar
    LlenaListaNiveles (lblAnalista)
End If

End Sub

Private Sub CmdModificar_Click()
    lsNueMod = "M"
    Call HabilitarComandos(False, False, False, True, True, False)
    Me.frmAnalista.Enabled = False
    Me.frmDetalle.Enabled = True
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CargaAnalistas
    Call HabilitarComandos(True, True, False, False, False, True)
    Me.frmAnalista.Enabled = True
    Me.frmDetalle.Enabled = False
End Sub

Private Sub lvwMetas_Click()
    Call HabilitarComandos(True, True, True, False, False, True)
End Sub

Private Sub CargaAnalistas()
Dim lr As ADODB.Recordset
Dim lsSQL As String, sApoderados As String
Dim loConecta As DConecta
Dim oGen As DGeneral

On Error GoTo ERRORCargaAnalistas
    
    Set oGen = New DGeneral
    sApoderados = oGen.LeeConstSistema(gConstSistRHCargoCodApoderados)
    Set oGen = Nothing
    
    lsSQL = "Select R.cPersCod, P.cPersNombre from RRHH R inner join Persona P ON R.cPersCod = P.cpersCod " _
        & " AND nRHEstado = 201 " _
        & " Inner join RHCargos RC ON R.cPersCod = RC.cPersCod " _
        & " Where  RC.cRHCargoCod in (" & sApoderados & ") AND RC.dRHCargoFecha = (select MAX(dRHCargoFecha) from RHCargos RHC2 where RHC2.cPersCod = RC.cPersCod) " _
        & " Order by P.cPersNombre "
    
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set lr = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
    
    cboAnalista.Clear
    Do While Not lr.EOF
        cboAnalista.AddItem PstaNombre(lr!cPersNombre, False) & Space(100) & lr!cPersCod
        lr.MoveNext
    Loop
    lr.Close
    Set lr = Nothing
    Exit Sub

ERRORCargaAnalistas:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Limpiar()
Dim lnCont As Integer
For lnCont = 1 To 2
    Me.Chkprod(lnCont).value = 0
Next lnCont
For lnCont = 1 To 4
    Me.chkNiv(lnCont).value = 0
Next lnCont

End Sub

Private Sub HabilitarComandos(pcmdNuevo As Boolean, pcmdModificar As Boolean, _
                              pcmdeliminar As Boolean, pcmdGrabar As Boolean, pcmdCancelar As Boolean, pcmdImprimir As Boolean)
    Me.cmdnuevo.Enabled = pcmdNuevo
    Me.cmdModificar.Enabled = pcmdModificar
    Me.cmdGrabar.Enabled = pcmdGrabar
    Me.cmdCancelar.Enabled = pcmdCancelar
    Me.cmdImprimir.Enabled = pcmdImprimir
End Sub

Private Sub LlenaListaNiveles(ByVal pAnalista As String)
Dim L As ListItem
Dim sSQLMet As String
Dim rsNivel As New ADODB.Recordset
Dim loConecta As DConecta

    lvwMetas.ListItems.Clear
    If Len(pAnalista) = 0 Then
      Exit Sub
    End If
    
    sSQLMet = " SELECT N.cProduct, N.nNivel,  N.nMontoMin, N.nMontoMax FROM ColocCredPersNivelesApr U " _
            & " INNER JOIN ColocCredNivelesApr N ON U.cCodNiv = N.cCodNiv " _
            & " WHERE U.cPersCod = '" & Me.lblAnalista & "' and Substring(N.cCodNiv,2,2) = '21' " _
            & " ORDER BY N.cProduct,  N.nMontoMin "

    Set loConecta = New DConecta
    Call loConecta.AbreConexion
        Set rsNivel = loConecta.CargaRecordSet(sSQLMet)
    Call loConecta.CierraConexion
    Set loConecta = Nothing
    
    If rsNivel.BOF And rsNivel.EOF Then
       MsgBox " Usuario no tiene permisos para APROBAR Cartas Fianza ", vbInformation, " Aviso "
       Exit Sub
    Else
       Do While Not rsNivel.EOF
            Set L = lvwMetas.ListItems.Add(, , Trim(rsNivel!cProduct))
                'L.SubItems(1) = rsNivel!cRefinan
                L.SubItems(2) = Format(rsNivel!nMontoMin, "#0.00")
                L.SubItems(3) = Format(rsNivel!nMontoMax, "#0.00")
            
            Select Case rsNivel!cProduct
                Case "121"
                    Chkprod(1).value = 1
                Case "221"
                    Chkprod(2).value = 1
            End Select
            
            Select Case rsNivel!nNivel
                Case "1"
                    chkNiv(1).value = 1
                Case "2"
                    chkNiv(2).value = 1
                Case "3"
                    chkNiv(3).value = 1
                Case "4"
                    chkNiv(4).value = 1
            End Select
        
            rsNivel.MoveNext
       Loop
    End If
    rsNivel.Close
    Set rsNivel = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

