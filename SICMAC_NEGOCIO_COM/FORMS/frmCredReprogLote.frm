VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredReprogLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprogramacion en Lote"
   ClientHeight    =   2805
   ClientLeft      =   1605
   ClientTop       =   3075
   ClientWidth     =   8805
   Icon            =   "frmCredReprogLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2730
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8670
      Begin VB.Frame Frame2 
         Height          =   2010
         Left            =   210
         TabIndex        =   3
         Top             =   585
         Width           =   7995
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   390
            Left            =   4845
            TabIndex        =   11
            Top             =   780
            Width           =   2370
         End
         Begin VB.CommandButton CmdReprogramar 
            Caption         =   "&Reprogramar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4845
            TabIndex        =   10
            Top             =   345
            Width           =   2370
         End
         Begin MSComctlLib.ProgressBar PBBarra 
            Height          =   240
            Left            =   225
            TabIndex        =   9
            Top             =   1635
            Width           =   7530
            _ExtentX        =   13282
            _ExtentY        =   423
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Reprogramacion"
            Height          =   1305
            Left            =   120
            TabIndex        =   4
            Top             =   210
            Width           =   4185
            Begin VB.OptionButton OptTipoReprog 
               Caption         =   "Por Fecha"
               Height          =   345
               Index           =   0
               Left            =   465
               TabIndex        =   7
               Top             =   255
               Value           =   -1  'True
               Width           =   1140
            End
            Begin VB.OptionButton OptTipoReprog 
               Caption         =   " Total (Nuevo Credito)"
               Height          =   345
               Index           =   1
               Left            =   2070
               TabIndex        =   6
               Top             =   255
               Width           =   1920
            End
            Begin MSMask.MaskEdBox TxtFecha 
               Height          =   345
               Left            =   2520
               TabIndex        =   5
               Top             =   825
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   609
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Line Line1 
               X1              =   165
               X2              =   4065
               Y1              =   660
               Y2              =   660
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Reprogramacion :"
               Height          =   195
               Left            =   435
               TabIndex        =   8
               Top             =   885
               Width           =   2010
            End
         End
         Begin VB.CommandButton CmdAceptar 
            Caption         =   "&Aceptar"
            Height          =   390
            Left            =   4845
            TabIndex        =   12
            Top             =   780
            Width           =   1170
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   390
            Left            =   6045
            TabIndex        =   13
            Top             =   780
            Width           =   1170
         End
      End
      Begin VB.ComboBox CmbInstitucion 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   6735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Institucion :"
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
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmCredReprogLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CargaInstituciones()
Dim oPersonas As COMDPersona.DCOMPersonas
Set oPersonas = New COMDPersona.DCOMPersonas
Dim R As ADODB.Recordset

Set R = oPersonas.RecuperaPersonasTipo(Trim(str(gPersTipoConvenio)))
Set oPersonas = Nothing

'Call CargaComboPersonasTipo(gPersTipoConvenio, cmbInstitucion)
    
Do While Not R.EOF
    CmbInstitucion.AddItem PstaNombre(R!cPersNombre) & Space(250) & R!cPersCod
    R.MoveNext
Loop

Set R = Nothing
End Sub

'JUEZ 20130417 *****************************************************
'Private Sub CmbInstitucion_Click()
'    Dim rsCred As ADODB.Recordset
'    Dim lsCad As String
'    Dim lsCabecera  As String
'    Dim loPrevio As previo.clsprevio
'    Dim oImpre As COMFunciones.FCOMImpresion
'    Dim oCredito As COMDCredito.DCOMCreditos
'    Dim oCred As COMDCredito.DCOMCredito
'
'    Set oCredito = New COMDCredito.DCOMCreditos
'    Set oCred = New COMDCredito.DCOMCredito
'    Set oImpre = New COMFunciones.FCOMImpresion
'    Set rsCred = oCredito.RecuperaCreditosVigxInstitucion(Trim(Right(CmbInstitucion.Text, 20)))
'    Do While Not rsCred.EOF
'        lsCabecera = oImpre.CabeceraPagina("LISTA DE CREDITOS SIN PAGO DE COMISION POR REPROGRAMACION", 0, 1, gsNomAge, "CMAC MAYNAS SA", gdFecSis, , False)
'        If oCred.ExisteComisionVigente(rsCred!cCtaCod, gComisionReprogCredito) = False Then
'            lsCad = lsCad & rsCred!cCtaCod & Chr$(10)
'        End If
'        rsCred.MoveNext
'    Loop
'    If Trim(lsCad) <> "" Then
'        Set loPrevio = New previo.clsprevio
'        MsgBox "Para realizar la Reprogramación en Lote es necesario que todos los créditos realicen el pago de la Comisión por Reprogramación en ventanilla", vbInformation, "Aviso"
'        loPrevio.Show lsCabecera & lsCad, "Lista de Créditos sin comisión por reprogramación", True
'        Set loPrevio = Nothing
'        CmbInstitucion.ListIndex = -1
'    End If
'    Set oCredito = Nothing
'    Set oCred = Nothing
'End Sub
'END JUEZ **********************************************************

Private Sub CmdReprogramar_Click()
'Dim odCredito As COMDCredito.DCOMCreditos
Dim oCredito As COMNCredito.NCOMCredito
Dim loVistoElectronico As SICMACT.frmVistoElectronico 'JUEZ 20130417
Dim lbVistoVal As Boolean 'JUEZ 20130417
'Dim R As ADODB.Recordset
'Dim dFecha As Date
'Dim MatCalend As Variant

    On Error GoTo ErrorCmdReprogramar_Click
    
    'JUEZ 20130417 ************************************************************
    If Trim(Right(CmbInstitucion.Text, 20)) <> "" Then
        If IsDate(TxtFecha.Text) Then
            Set loVistoElectronico = New SICMACT.frmVistoElectronico
            lbVistoVal = loVistoElectronico.Inicio(2, "", " ")
            If lbVistoVal Then
                Set oCredito = New COMNCredito.NCOMCredito
                Call oCredito.ReprogramarCreditoLote(Trim(Right(CmbInstitucion.Text, 20)), IIf(OptTipoReprog(0).value, 1, 2), gdFecSis, gsCodUser, gsCodAge, TxtFecha.Text)
                Set oCredito = Nothing
    
'    If OptTipoReprog(0).value Then
'        If ValidaFecha(txtFecha.Text) <> "" Then
'            MsgBox ValidaFecha(txtFecha.Text), vbInformation, "Aviso"
'            Exit Sub
'        End If
'    End If
'    Set odCredito = New COMDCredito.DCOMCreditos
'    Set oNCredito = New COMNCredito.NCOMCredito
    
'    Set R = odCredito.RecuperaCreditosVigxInstitucion(Trim(Right(cmbInstitucion.Text, 20)))
'    Do While Not R.EOF
'        If OptTipoReprog(0).value Then
'            dFecha = CDate(txtFecha.Text)
'            If dFecha > R!dFecVencPend Then
'                MatCalend = oNCredito.ReprogramarCreditoenMemoria(R!cCtaCod, R!nTasaInteres, R!dFecVencPend, dFecha, R!nCuotaPend - 1, 1)
'                Call oNCredito.ReprogramarCredito(R!cCtaCod, MatCalend, 1, , , gdFecSis, , gsCodUser, gsCodAge)
'            End If
'        Else
'            MatCalend = oNCredito.ReprogramarCreditoenMemoriaTotal(R!cCtaCod, gdFecSis)
'            Call oNCredito.ReprogramarCredito(R!cCtaCod, MatCalend, 1, , , gdFecSis, , gsCodUser, gsCodAge)
'        End If
'        PBBarra.value = (R.Bookmark / R.RecordCount) * 100
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    Set oCredito = Nothing
'    Set oNCredito = Nothing
                MsgBox "Proceso de Reprogramacion en Lote Terminado", vbInformation, "Aviso"
                CmbInstitucion.ListIndex = -1
                TxtFecha.Text = "__/__/____"
                Exit Sub
            Else
                Exit Sub
            End If
        Else
            MsgBox "Debe insertar correctamente la fecha", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        MsgBox "Debe seleccionar una institución", vbInformation, "Aviso"
        Exit Sub
    End If
    'END JUEZ *****************************************************************
ErrorCmdReprogramar_Click:
    MsgBox err.Description, vbCritical, "Aviso"


End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Call CargaInstituciones
End Sub

Private Sub OptTipoReprog_Click(Index As Integer)
    If Index = 0 Then
        TxtFecha.Enabled = True
    Else
        TxtFecha.Enabled = False
    End If
    TxtFecha.Text = "__/__/____"
End Sub
