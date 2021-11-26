VERSION 5.00
Begin VB.Form frmCajaChicaConfEncargatura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caja Chica: Confirmación de Encargatura"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSaldo 
      Caption         =   "Saldo Actual Caja Chica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   8355
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   4320
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "Rechazar"
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSaldoActual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         Caption         =   "S/. :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Frame fraEncActual 
      Caption         =   "Datos: Nuevo Encargado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1545
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   8355
      Begin Sicmact.Usuario Usu 
         Left            =   7800
         Top             =   240
         _ExtentX        =   820
         _ExtentY        =   820
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   4545
      End
      Begin VB.Label lblPerscod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lblDni 
         Caption         =   "DNI :"
         Height          =   255
         Left            =   6120
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblNroDni 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   330
         Left            =   6720
         TabIndex        =   6
         Top             =   960
         Width           =   1380
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   825
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8355
      Begin VB.Label lblAreAge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7665
         TabIndex        =   4
         Top             =   270
         Width           =   525
      End
      Begin VB.Label lblNro 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   210
         Left            =   7380
         TabIndex        =   3
         Top             =   330
         Width           =   255
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   270
         Width           =   4485
      End
      Begin VB.Label lblAreaAge 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   210
         Left            =   90
         TabIndex        =   1
         Top             =   330
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmCajaChicaConfEncargatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Persona As UPersona
Dim oCajaCH As nCajaChica
Dim oArendir As NARendir

Private Sub cmdConfirmar_Click()
Dim oCon As NContFunciones
Dim lsMovNro As String
Dim lsTexto As String
Dim oContImp As NContImprimir

If MsgBox("¿Desea Confirmar ser el Responsable de Caja Chica?", vbYesNo + vbQuestion, "Aviso") = vbYes Then

    Set oCon = New NContFunciones
    Set oContImp = New NContImprimir
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    lsTexto = oCajaCH.GrabaResponsableNuevoCH(lsMovNro, gsOpeCod, "Cambio Responsable Caja Chica - Confirmación", Mid(lblAreAge, 1, 3), Mid(lblAreAge, 4, 2), Val(lblNroProcCH), "", lblPerscod, lblSaldoActual, 1)

    If lsTexto = "" Then
        MsgBox "No se realizo la operación" & vbCrLf & "Verifique los datos del nuevo responsable", vbInformation, "Aviso"
    Else
        MsgBox "El proceso de confirmación se realizo con exito", vbInformation, "Aviso"
        Unload Me
    End If
End If
End Sub

Private Sub cmdRechazar_Click()
Dim oCon As NContFunciones
Dim lsMovNro As String
Dim lsTexto As String
Dim oContImp As NContImprimir

If MsgBox("Ud. no podra ser asignado como responsable a menos que se vuelva a realizar la designación en Contabilidad ¿Desea Rechazar ser el Responsable de Caja Chica?", vbYesNo + vbQuestion, "Aviso") = vbYes Then

    Set oCon = New NContFunciones
    Set oContImp = New NContImprimir
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    lsTexto = oCajaCH.GrabaResponsableNuevoCH(lsMovNro, gsOpeCod, "Cambio Responsable Caja Chica - Rechazo", Mid(lblAreAge, 1, 3), Mid(lblAreAge, 4, 2), Val(lblNroProcCH), "", lblPerscod, lblSaldoActual, 2)

    If lsTexto = "" Then
        MsgBox "No se realizo la operación" & vbCrLf & "Verifique los datos del nuevo responsable", vbInformation, "Aviso"
    Else
        MsgBox "El proceso de rechazo se realizo con exito", vbInformation, "Aviso"
        Unload Me
    End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim ClsPersona As DPersonas
Dim R As ADODB.Recordset
Dim rsDatosResp As ADODB.Recordset
Dim rsAreaAge As ADODB.Recordset
Dim lsCodigo As String
Dim lsAgencia As String
Set ClsPersona = New DPersonas
Set oCajaCH = New nCajaChica
Set oArendir = New NARendir

     Set R = ClsPersona.RecuperaAreaAge(gsCodUser)
     If R!cAreaCodActual = "023" Or R!cAreaCodActual = "042" Or R!cAreaCodActual = "043" Or R!cAreaCodActual = "064" Or R!cAreaCodActual = "050" Then  'ande 20170801 agregué 050 para que reconozca el area de asesoría legal
        lsCodigo = R!cAreaCodActual
        lsAgencia = ""
     Else
        lsCodigo = R!cAreaCodActual + R!cAgenciaActual
        lsAgencia = R!cAgenciaActual
     End If
     If Not R.EOF Then
        Set rsDatosResp = oCajaCH.devolverDatosRespCajaChica(R!cAreaCodActual, lsAgencia, Format(gdFecSis, "yyyyMMdd"))
            If Not rsDatosResp.EOF Then
                    If rsDatosResp!nEstado = 0 Then
                        If LCase(gsCodUser) = LCase(rsDatosResp!cUser) Then
                            Set rsAreaAge = oArendir.EmiteCajasChicas
                                Do While Not (rsAreaAge.EOF Or rsAreaAge.BOF)
                                    If lsCodigo = rsAreaAge!Codigo Then
                                        lblAreAge = rsAreaAge!Codigo
                                        lblCajaChicaDesc = rsAreaAge!Descripcion
                                        lblNroProcCH = rsDatosResp!nProNro
                                        lblPerscod = rsDatosResp!cPersCodNuevo
                                        lblPersNombre = rsDatosResp!cPersNombre
                                        lblNroDni = IIf(rsDatosResp!cPersIDnroDNI = "", rsDatosResp!cPersIDnroRUC, rsDatosResp!cPersIDnroDNI)
                                        lblSaldoActual = rsDatosResp!nSaldoActual
                                    End If
                                    rsAreaAge.MoveNext
                                Loop
                            Set rsAreaAge = Nothing
                        Else
                            MsgBox "Ud. no fue definido como encargado de Caja Chica", vbInformation, "Aviso"
                            Me.cmdConfirmar.Enabled = False
                            Me.cmdRechazar.Enabled = False
                            Exit Sub
                        End If
                    ElseIf rsDatosResp!nEstado = 1 Then
                        MsgBox "Ud. ya confirmo ser responsable de Caja Chica", vbInformation, "Aviso"
                        Me.cmdConfirmar.Enabled = False
                        Me.cmdRechazar.Enabled = False
                    ElseIf rsDatosResp!nEstado = 2 Then
                        MsgBox "Ud. rechazo ser responsable de Caja Chica", vbInformation, "Aviso"
                        Me.cmdConfirmar.Enabled = False
                        Me.cmdRechazar.Enabled = False
                    End If
            Else
                MsgBox "No existen usuarios para cambio de responsables Caja Chica", vbInformation, "Aviso"
                Me.cmdConfirmar.Enabled = False
                Me.cmdRechazar.Enabled = False
            End If
        Set rsDatosResp = Nothing
     End If
     Set R = Nothing
End Sub
