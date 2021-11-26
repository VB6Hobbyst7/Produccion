VERSION 5.00
Begin VB.Form frmCredEvalSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Formato de Evaluación"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "frmCredEvalSeleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEvaluar 
      Caption         =   "&Evaluar"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame fraBuscarSolicitud 
      Caption         =   "Buscar Solicitud"
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
      Height          =   1365
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7515
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   1065
         Left            =   5040
         TabIndex        =   5
         Top             =   120
         Width           =   2355
         Begin VB.ListBox LstCred 
            Height          =   645
            Left            =   75
            TabIndex        =   6
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   741
         Texto           =   "Crédito:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredEvalSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredEvalSeleeccion
'** Descripción : Administración  de seleccion de formato
'**               segun RFC090-2012
'** Creación : WIOR, 20120903 09:00:00 AM
'**********************************************************************************************
Option Explicit
Dim fnTipo As Integer

Private Sub ActxCta_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(Me.ActxCta.NroCuenta)) < 18 Then
        Me.cmdEvaluar.Enabled = False
    End If
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdEvaluar.Enabled = True
        cmdEvaluar.SetFocus
    End If
End Sub

Private Sub cmdBuscar_Click()
'Dim oCredito As COMDCredito.DCOMCredito PTI120170530 segun ERS014-2017
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
Call CargarControles
    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
         'Set oCredito = New COMDCredito.DCOMCredito
         Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval 'PTI120170530 segun ERS014-2017
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval 'PTI120170530 segun ERS014-2017
        'Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstSolic, gColocEstSolic))
         Set R = oDCOMFormatosEval.RecuperaCreditosEvaluadosFormatos(oPers.sPersCod, fnTipo) 'PTI120170530 segun ERS014-2017
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
            
           
            LstCred.Selected(0) = True
        Loop
        LstCred.SetFocus
        R.Close
        Set R = Nothing
        Set oDCOMFormatosEval = Nothing 'PTI120170530 segun ERS014-2017
        
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Solicitados", vbInformation, "Aviso"
        cmdBuscar.SetFocus
    End If

    Set oPers = Nothing
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub


'Private Sub cmdEvaluar_Click()
'If Len(ActxCta.NroCuenta) < 18 Then
'    MsgBox "Nº de Crédito Incompleto.", vbInformation, "Aviso"
'Else
'    Call EvaluarCredito(ActxCta.NroCuenta)
'End If
'End Sub

Private Sub cmdEvaluar_Click()
    Dim oEval As COMDCredito.DCOMFormatosEval
    Dim oCred As COMDCredito.DCOMCredito
    Dim bConsultar As Boolean
    Dim lsCtaCod As String
    Dim nEstado As Integer

    If Len(ActxCta.NroCuenta) < 18 Then
        MsgBox "Nº de Crédito Incompleto.", vbInformation, "Aviso"
    Else
        Set oEval = New COMDCredito.DCOMFormatosEval
        Set oCred = New COMDCredito.DCOMCredito
        
        lsCtaCod = ActxCta.NroCuenta
        nEstado = oCred.RecuperaEstadoCredito(lsCtaCod)
        
        If nEstado = 0 Then
            MsgBox "Nº de Crédito no existe.", vbInformation, "Aviso"
            Exit Sub
        End If
        
        If fnTipo = 1 Then
            If oEval.RecuperaFormatoEvaluacion(lsCtaCod).RecordCount > 0 Then
                MsgBox "El crédito ya cuenta con Evaluación, verifique en la opción de Mantenimiento o Consulta.", vbInformation, "Aviso"
                Exit Sub
            End If
            If Not (nEstado = 2000 Or nEstado = 2001) Then
                MsgBox "El crédito tiene un estado diferente a SOLICITADO y/o SUGERIDO, no se podrá registrar Evaluación.", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf fnTipo = 2 Then
            If oEval.RecuperaFormatoEvaluacion(lsCtaCod).RecordCount = 0 Then
                MsgBox "El crédito no cuenta con Evaluación, ingréselo en la opción de Registro.", vbInformation, "Aviso"
                Exit Sub
            End If
            If Not (nEstado = 2000 Or nEstado = 2001) Then
                MsgBox "El crédito tiene un estado diferente a SOLICITADO y/o SUGERIDO, no se podrá editar la Evaluación", vbInformation, "Aviso"
                Exit Sub
            End If
        ElseIf fnTipo = 3 Then
            If oEval.RecuperaFormatoEvaluacion(lsCtaCod).RecordCount = 0 Then
                MsgBox "El crédito no cuenta con Evaluación.", vbInformation, "Aviso"
                Exit Sub
            End If
            bConsultar = True
        End If
        
        Call EvaluarCredito(lsCtaCod, True, , , , , bConsultar)
    End If
End Sub

Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub CargarControles()
ActxCta.NroCuenta = ""
Me.cmdEvaluar.Enabled = False
ActxCta.CMAC = gsCodCMAC
ActxCta.Age = gsCodAge
'PTI120170530 segun ERS014-2017
If fnTipo = 3 Then
ActxCta.EnabledCMAC = True
ActxCta.EnabledAge = True
Else
ActxCta.EnabledCMAC = False
ActxCta.EnabledAge = False
End If
'FIN PTI120170530 segun ERS014-2017
ActxCta.EnabledCMAC = False
ActxCta.EnabledAge = False
End Sub

Public Sub Inicio(ByVal pnTipo As Integer)
Dim NCredito As COMNCredito.NCOMCredito
Dim nExisteAgencia As Integer

Set NCredito = New COMNCredito.NCOMCredito
nExisteAgencia = NCredito.ObtieneAgenciaCredEval(gsCodAge)

'If nExisteAgencia = 0 Then
'   MsgBox "Agencia No configurada para este proceso.", vbInformation, "Aviso"
'   Exit Sub
'Else
    fnTipo = pnTipo
    Call CargarControles
    Me.Show 1
'End If
End Sub

'Private Sub EvaluarCredito(ByVal pcCtaCod As String)
'Dim DCredito As COMDCredito.DCOMCredito
'Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
'Dim rsDCredito As ADODB.Recordset
'Dim nEstado As Integer
'Dim nFormato As Integer
'Dim nMonto As Double
'Dim cPrd As String
'Dim cSPrd As String
'Dim sAgeCod As String
'Dim NCredito As COMNCredito.NCOMCredito
'Dim nExisteAgencia As Integer
'
'Set NCredito = New COMNCredito.NCOMCredito
'Set DCredito = New COMDCredito.DCOMCredito
'Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
'
'
'nEstado = DCredito.RecuperaEstadoCredito(pcCtaCod)
'
'If nEstado = 0 Then
'    MsgBox "Nº de Crédito no existe.", vbInformation, "Aviso"
'    Exit Sub
'Else
'    If nEstado = 2000 Then
'        Set rsDCredito = DCredito.RecuperaSolicitudDatoBasicos(pcCtaCod)
'        If rsDCredito.RecordCount > 0 Then
'            nMonto = CDbl(Trim(rsDCredito!nMonto))
'            cSPrd = Trim(rsDCredito!cTpoProdCod)
'            cPrd = Mid(cSPrd, 1, 1) & "00"
'            sAgeCod = Trim(rsDCredito!cAgeCodAct)
'
'            nExisteAgencia = NCredito.ObtieneAgenciaCredEval(sAgeCod)
'            If nExisteAgencia = 0 Then
'                MsgBox "Credito no Pertenece a una configurada Agencia para este proceso.", vbInformation, "Aviso"
'                Call CargarControles
'                Exit Sub
'            End If
'
'            If Mid(pcCtaCod, 9, 1) = "2" Then
'                nMonto = nMonto * CDbl(oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia))
'            End If
'        End If
'
'        nFormato = DCredito.AsignarFormato(cPrd, cSPrd, nMonto)
'
'        Select Case nFormato
'            Case 0, 4, 5: MsgBox "Crédito no se adecua para este Proceso.", vbInformation, "Aviso"
'            Case 1: Call frmCredEvalFormato1.Inicio(pcCtaCod, fnTipo)
'            Case 2: Call frmCredEvalFormato2.Inicio(pcCtaCod, fnTipo)
'            Case 3: Call frmCredEvalFormato3.Inicio(pcCtaCod, fnTipo)
'            'Case 4: Call frmCredEvalFormato4.Inicio(pcCtaCod, fnTipo)
'            'Case 5: Call frmCredEvalFormato5.Inicio(pcCtaCod, fnTipo)
'        End Select
'    Else
'        MsgBox "Nº de Crédito no se encuentra en estado Solicitado.", vbInformation, "Aviso"
'        Exit Sub
'    End If
'End If
'
'End Sub

