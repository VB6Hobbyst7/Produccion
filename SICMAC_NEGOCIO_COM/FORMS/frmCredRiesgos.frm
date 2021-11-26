VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos: Informe de Riesgos"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16065
   Icon            =   "frmCredRiesgos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   16065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Créditos"
      TabPicture(0)   =   "frmCredRiesgos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "feCredInfRiesgo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFechExp"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdVinculados"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdInfRiesgos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdActualizar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCerrar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtBuscar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSalidaObs"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdRIngresoObs"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdAutorizar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   9720
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         Begin VB.OptionButton optCampana 
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   16
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton optCampana 
            Caption         =   "No"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   15
            Top             =   200
            Width           =   615
         End
         Begin VB.OptionButton optCampana 
            Caption         =   "Si"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   220
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdAutorizar 
         Caption         =   "Autorizar"
         Height          =   375
         Left            =   7920
         TabIndex        =   11
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmdRIngresoObs 
         Caption         =   "Reingreso x Obs."
         Height          =   375
         Left            =   6360
         TabIndex        =   10
         Top             =   5880
         Width           =   1580
      End
      Begin VB.CommandButton cmdSalidaObs 
         Caption         =   "Salida x Obs."
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   5880
         Width           =   1580
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   6855
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   14400
         TabIndex        =   6
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   13200
         TabIndex        =   5
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdInfRiesgos 
         Caption         =   "Inf. Riesgos"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   5880
         Width           =   1580
      End
      Begin VB.CommandButton cmdVinculados 
         Caption         =   "Ver Vinculados"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   5880
         Width           =   1580
      End
      Begin VB.CommandButton cmdFechExp 
         Caption         =   "Fecha Exp."
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   5880
         Width           =   1580
      End
      Begin SICMACT.FlexEdit feCredInfRiesgo 
         Height          =   4740
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   15570
         _ExtentX        =   27464
         _ExtentY        =   8361
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCredRiesgos.frx":0326
         EncabezadosAnchos=   "400-2000-1800-2500-2200-2200-1000-1200-1800-0-1800-1800-1500"
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-R-C-C-C-C-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0"
         CantEntero      =   10
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label2 
         Caption         =   "Campañas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   8760
         TabIndex        =   13
         Top             =   550
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Buscar Por Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCredRiesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fnCampana As Integer 'JOEP-ERS064 20161104

'WIOR 20150917 ***
Private rsCreditoTotal As ADODB.Recordset
'WIOR FIN ********

Private rsCredConCampTotal As ADODB.Recordset 'JOEP-ERS064 20161104
Private rsCredSinCampTotal As ADODB.Recordset 'JOEP-ERS064 20161104

'RECO20161018 ERS060-2016 *********************************
Dim oNCOMColocEval As NCOMColocEval
Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
Dim lnFilaAct As Integer
'RECO FIN *************************************************

Private Sub cmdActualizar_Click()
'JOEP-ERS064 20161104
    If fnCampana = 1 Then
        optCampana(1).value = False
        fnCampana = 0
    ElseIf fnCampana = 2 Then
        optCampana(2).value = False
        fnCampana = 0
    End If
'JOEP-ERS064 20161104

    'WIOR 20150917 ***
    txtBuscar.Text = ""
    txtBuscar.SetFocus
    'WIOR FIN *********
    Call LLenarGrilla
    Call feCredInfRiesgo_Click 'RECO20161019 ERS060-2016
End Sub

'JOEP-ERS064 20161019
Private Sub cmdAutorizar_Click()
Dim lnCodForm As Integer
Dim rsFormatoEval As New ADODB.Recordset
Dim oDCOMFormatosEval As New COMDCredito.DCOMFormatosEval
Dim lcMovNro As String 'ARLO20170926 ERS060-2016
Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones 'ARLO20171125
    
Dim rsExpRU As ADODB.Recordset
Dim oDCredExpRsgUn As New COMDCredito.DCOMNivelAprobacion

Set rsExpRU = oDCredExpRsgUn.ObtineRiesgoUnicoCred(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)))
Set rsFormatoEval = oDCOMFormatosEval.RecuperaFormatoEvaluacion(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)))

If fnCampana = 1 Then
        
    If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2) = "" Then
        MsgBox "Seleccione el credito correctamente.", vbInformation, "Aviso"
    Else
    
        If Not (rsFormatoEval.BOF Or rsFormatoEval.EOF) Then
            lnCodForm = rsFormatoEval!nCodForm
        Else
            MsgBox "No se ha registrado la evaluación del crédito" & Chr(10) & " - Por favor registrar el formato de evaluación correspondiente.", vbInformation, "Alerta"
            Screen.MousePointer = 0
            Exit Sub
        End If
    
        If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 8) = "" Then
            MsgBox "Aun no registro la fecha de recepcion del expediente.", vbInformation, "Aviso"
            Call LLenarGrilla
        ElseIf CDbl(rsExpRU!nMonto) >= 30001 Then
            If MsgBox("Desea Registrar el Informe de Riesgo" & Chr(10) & "Exp. Riesgo Unico :" & rsExpRU!nMonto & "", vbYesNo + vbInformation, "Confirmación") = vbYes Then
                Call frmCredRiesgoInformeMontos.Inicio(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2))
            Else
                Exit Sub
            End If
        Else
             Call frmCredRiesgoInformeMontos.Inicio(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2))
        End If
    End If
ElseIf fnCampana = 2 Then
    If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2) = "" Then
        MsgBox "Seleccione el credito correctamente.", vbInformation, "Aviso"
    Else
        If Not (rsFormatoEval.BOF Or rsFormatoEval.EOF) Then
            lnCodForm = rsFormatoEval!nCodForm
        Else
            MsgBox "No se ha registrado la evaluación del crédito" & Chr(10) & " - Por favor registrar el formato de evaluación correspondiente.", vbInformation, "Alerta"
            Screen.MousePointer = 0
            Exit Sub
        End If
    
        If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 8) = "" Then
            MsgBox "Aun no registro la fecha de recepcion del expediente.", vbInformation, "Aviso"
            Call LLenarGrilla
        ElseIf CDbl(rsExpRU!nMonto) >= 30001 Then
            If MsgBox("Desea Registrar el Informe de Riesgo" & Chr(10) & "Exp. Riesgo Unico : " & Format(rsExpRU!nMonto, "#,##0.00") & "", vbYesNo + vbInformation, "Confirmación") = vbYes Then
                Call frmCredRiesgoInformeMontos.Inicio(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2))
            Else
                Exit Sub
            End If
        Else
             Call frmCredRiesgoInformeMontos.Inicio(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2))
        End If
    End If
Else
        
    If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2) = "" Then
        MsgBox "Seleccione el credito correctamente.", vbInformation, "Aviso"
    Else
    
        If Not (rsFormatoEval.BOF Or rsFormatoEval.EOF) Then
            lnCodForm = rsFormatoEval!nCodForm
        Else
            MsgBox "No se ha registrado la evaluación del crédito" & Chr(10) & " - Por favor registrar el formato de evaluación correspondiente.", vbInformation, "Alerta"
            Screen.MousePointer = 0
            Exit Sub
        End If
    
        If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 8) = "" Then
            MsgBox "Aun no registro la fecha de recepcion del expediente.", vbInformation, "Aviso"
            Call LLenarGrilla
        ElseIf CDbl(rsExpRU!nMonto) >= 30001 Then
            If MsgBox("Desea Registrar el Informe de Riesgo" & Chr(10) & "Exp. Riesgo Unico :" & Format(rsExpRU!nMonto, "#,##0.00") & "", vbYesNo + vbInformation, "Confirmación") = vbYes Then
                Set oNCOMColocEval = New NCOMColocEval 'ARLO20170926 ERS060-2016
                lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'ARLO20170926 ERS060-2016
                Call oNCOMColocEval.insEstadosExpediente(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)), "Gerencia de Riesgos", "", "", "", lcMovNro, 1, 2001, gTpoRegCtrlRiesgos) 'ARLO20170926 ERS060-2016
                Call frmCredRiesgoInformeMontos.Inicio(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2))
            Else
                'Call LLenarGrilla
                Exit Sub
            End If
        Else
             Call frmCredRiesgoInformeMontos.Inicio(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2))
        End If
    End If
End If
End Sub
'JOEP-ERS064 20161019 ******************************

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdFechExp_Click()
    Dim oCredito As COMDCredito.DCOMCredito
    Dim lsMovNro As String 'EJVG20160531
    Dim lcMovNro As String 'RECO20161018 ERS060-2016
    Set oCredito = New COMDCredito.DCOMCredito
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones 'RECO20161018 ERS060-2016
    Set oNCOMColocEval = New NCOMColocEval 'RECO20161018 ERS060-2016
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'RECO20161018 ERS060-2016
    lnFilaAct = feCredInfRiesgo.row 'RECO20161019 ERS060-2016
    If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2) = "" Then
        MsgBox "Seleccione el credito correctamente.", vbInformation, "Aviso"
    Else
        If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 8) = "" Then
            If MsgBox("¿Está seguro de grabar la fecha de recepción del Expediente?", vbYesNo + vbInformation, "Confirmación") = vbYes Then
                lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) 'EJVG20160531
                'Call oCredito.OpeInformeRiesgo(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)), 2, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), , , , 1)
                oCredito.ActualizarInformeRiesgo Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)), 0, 1, , , , lsMovNro, lsMovNro  'EJVG20160531
                Call oNCOMColocEval.insEstadosExpediente(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)), "Gerencia de Riesgos", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlRiesgos) 'RECO20161018 ERS060-2016
                Call LLenarGrilla
                Set oNCOMColocEval = Nothing
                feCredInfRiesgo.row = lnFilaAct 'RECO20161019 ERS060-2016
                feCredInfRiesgo.TopRow = lnFilaAct 'RECO20161019 ERS060-2016
                Call feCredInfRiesgo_Click 'RECO20161019 ERS060-2016
            End If
        Else
            MsgBox "Ya se recepcionó el Expediente.", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub cmdInfRiesgos_Click()
    If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2) = "" Then
        MsgBox "Seleccione el credito correctamente.", vbInformation, "Aviso"
    Else
        If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 8) = "" Then
            MsgBox "Aun no registro la fecha de recepcion del expediente.", vbInformation, "Aviso"
            lnFilaAct = feCredInfRiesgo.row 'RECO20161019 ERS060-2016
            Call LLenarGrilla
            feCredInfRiesgo.row = lnFilaAct 'RECO20161019 ERS060-2016
            feCredInfRiesgo.TopRow = lnFilaAct 'RECO20161019 ERS060-2016
            Call feCredInfRiesgo_Click 'RECO20161019 ERS060-2016
        Else
            Screen.MousePointer = 11 'RECO20161019 ERS060-2016
            Call frmCredRiesgosInforme.Inicio(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)), True)
            'WIOR 20141010 True
        End If
    End If
End Sub
'RECO20161018 ERS060-2016*****************************************
Private Sub cmdRIngresoObs_Click()
    Set oNCOMColocEval = New NCOMColocEval
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lcMovNro As String
    lnFilaAct = feCredInfRiesgo.row
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2), gTpoRegCtrlRiesgos) 'BY ARLO MODIFY 20171027
    Call oNCOMColocEval.insEstadosExpediente(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2), "Gerencia de Riesgos", "", "", lcMovNro, "", 1, 2001, gTpoRegCtrlRiesgos)
    MsgBox "Re Ingreso de Expediente a Gerencia de Riesgos", vbInformation, "Aviso"
    Set oNCOMColocEval = Nothing
    Call LLenarGrilla
    feCredInfRiesgo.row = lnFilaAct
    feCredInfRiesgo.TopRow = lnFilaAct
    Call feCredInfRiesgo_Click
End Sub

Private Sub cmdSalidaObs_Click()
    Set oNCOMColocEval = New NCOMColocEval
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    
    Dim lcMovNro As String
    lnFilaAct = feCredInfRiesgo.row
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2), gTpoRegCtrlRiesgos) 'BY ARLO MODIFY 20171027
    Call oNCOMColocEval.insEstadosExpediente(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2), "Analista de Creditos", "", lcMovNro, "", "", 1, 2001, gTpoRegCtrlRiesgos)
           
    MsgBox "Expediente Salio por Observación de Gerencia de Riesgos", vbInformation, "Aviso"
    Set oNCOMColocEval = Nothing
    Call LLenarGrilla
    feCredInfRiesgo.row = lnFilaAct
    feCredInfRiesgo.TopRow = lnFilaAct
    Call feCredInfRiesgo_Click
End Sub
'RECO FIN ***********************************************************

Private Sub cmdVinculados_Click()
If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2) = "" Then
    MsgBox "Seleccione el credito correctamente.", vbInformation, "Aviso"
Else
    Call frmCredRiesgosVinculados.Inicio(Trim(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 2)), CDbl(feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 7))) 'WIOR 20120928 AGREGO MONTO SUGERIDO
End If
End Sub

'RECO20161019 ERS060-2016******************************
Private Sub feCredInfRiesgo_Click()
    Call ActivaBotones
End Sub
'RECO FIN *********************************************

Private Sub Form_Load()
    fnCampana = 0 'JOEP-ERS064 20161104
    
    'Call CentraForm(Me)'WIOR 20150917 COMENTADO
    'Me.Icon = LoadPicture(App.path & gsRutaIcono)'WIOR 20150917 COMENTADO
    txtBuscar.Text = "" 'WIOR 20150917
    Call LLenarGrilla
    Call feCredInfRiesgo_Click 'RECO20161019 ERS060-2016
End Sub

Public Sub LLenarGrilla()
    Dim oCredito As COMDCredito.DCOMCredito
    Dim rsCredito As ADODB.Recordset
    'Dim i As Integer'WIOR 20150917 COMENTO
            
    Set oCredito = New COMDCredito.DCOMCredito
    
    Set rsCredito = oCredito.CreditoInformeRiesgoPend
    Set rsCreditoTotal = rsCredito.Clone 'WIOR 20150917
    
    Call MostrarLista(rsCredito) 'WIOR 20150917
    'Call CampanaLista 'JOEP-ERS064 20161104
    'WIOR 20150917 COMENTO
    'Call LimpiaFlex(feCredInfRiesgo)
    'If rsCredito.RecordCount > 0 Then
    '    For i = 0 To rsCredito.RecordCount - 1
    '        feCredInfRiesgo.AdicionaFila
    '        feCredInfRiesgo.TextMatrix(i + 1, 0) = i + 1
    '        feCredInfRiesgo.TextMatrix(i + 1, 1) = rsCredito!Agencia
    '        feCredInfRiesgo.TextMatrix(i + 1, 2) = rsCredito!cCtaCod
    '        feCredInfRiesgo.TextMatrix(i + 1, 3) = PstaNombre(rsCredito!cPersNombre, True)
    '        feCredInfRiesgo.TextMatrix(i + 1, 4) = rsCredito!TpoProd
    '        feCredInfRiesgo.TextMatrix(i + 1, 5) = rsCredito!TpoCred
    '        feCredInfRiesgo.TextMatrix(i + 1, 6) = rsCredito!Moneda
    '        feCredInfRiesgo.TextMatrix(i + 1, 7) = rsCredito!nMonto
    '        feCredInfRiesgo.TextMatrix(i + 1, 8) = rsCredito!FechaExp
    '        rsCredito.MoveNext
    '    Next i
    'End If
    'feCredInfRiesgo.TopRow = 1 'WIOR 20130514
    'WIOR 20150917 ***
    If Trim(txtBuscar.Text) <> "" Then
        Call txtBuscar_Change
    End If
'WIOR FIN ********
End Sub

'WIOR 20150917 ***
Private Sub MostrarLista(ByVal pRs As ADODB.Recordset)

'JOEP-ERS064 20161103
    If fnCampana = 1 Then
        Call CampanaLista
        Exit Sub
    ElseIf fnCampana = 2 Then
        Call CampanaLista
        Exit Sub
    End If
'JOEP-ERS064 20161103

    Dim i As Integer
    Call LimpiaFlex(feCredInfRiesgo)
    If pRs.RecordCount > 0 Then
        For i = 0 To pRs.RecordCount - 1
            feCredInfRiesgo.AdicionaFila
            feCredInfRiesgo.TextMatrix(i + 1, 0) = i + 1
            feCredInfRiesgo.TextMatrix(i + 1, 1) = pRs!Agencia
            feCredInfRiesgo.TextMatrix(i + 1, 2) = pRs!cCtaCod
            feCredInfRiesgo.TextMatrix(i + 1, 3) = PstaNombre(pRs!cPersNombre, True)
            feCredInfRiesgo.TextMatrix(i + 1, 4) = pRs!TpoProd
            feCredInfRiesgo.TextMatrix(i + 1, 5) = pRs!TpoCred
            feCredInfRiesgo.TextMatrix(i + 1, 6) = pRs!Moneda
            feCredInfRiesgo.TextMatrix(i + 1, 7) = Format(pRs!nMonto, "###," & String(15, "#") & "#0.00")
            feCredInfRiesgo.TextMatrix(i + 1, 8) = pRs!FechaExp
            feCredInfRiesgo.TextMatrix(i + 1, 10) = pRs!FechaSaldiaObs
            feCredInfRiesgo.TextMatrix(i + 1, 11) = pRs!FechaReingresoObs
            feCredInfRiesgo.TextMatrix(i + 1, 12) = Format(pRs!ExpRU, "###," & String(15, "#") & "#0.00") 'JOEP20170523 ERS-064
            pRs.MoveNext
        Next i
    End If
    feCredInfRiesgo.TopRow = 1
    Call feCredInfRiesgo_Click 'RECO20161019 ERS060-2016
End Sub

'JOEP-ERS064 20161103
Private Sub optCampana_Click(Index As Integer)
    fnCampana = Index
    Call CampanaLista
End Sub

Private Sub CampanaLista()
Dim oCredCampLis As COMDCredito.DCOMCredito
Dim rsCredCampLisCon As ADODB.Recordset 'ERS-064 JOEP20170523
Dim rsCredCampLisSin As ADODB.Recordset 'ERS-064 JOEP20170523

Set oCredCampLis = New COMDCredito.DCOMCredito 'ERS-064 JOEP20170523
Set oCredCampLis = New COMDCredito.DCOMCredito 'ERS-064 JOEP20170523

'Inicio ERS-064 JOEP20170523
If fnCampana = 1 Then
    Set rsCredCampLisCon = oCredCampLis.CreditoInformeRiesgoPendConCampana
    Set rsCredConCampTotal = rsCredCampLisCon.Clone
ElseIf fnCampana = 2 Then
    Set rsCredCampLisSin = oCredCampLis.CreditoInformeRiesgoPendSinCampana
    Set rsCredSinCampTotal = rsCredCampLisSin.Clone
End If

    If fnCampana = 1 Then
        Call MostrarListaCampana(rsCredCampLisCon)
    ElseIf fnCampana = 2 Then
        Call MostrarListaCampana(rsCredCampLisSin)
    End If
'Fin ERS-064 JOEP20170523

End Sub

Private Sub MostrarListaCampana(ByVal pRs As ADODB.Recordset)
    Dim i As Integer
    Call LimpiaFlex(feCredInfRiesgo)
    If pRs.RecordCount > 0 Then
        For i = 0 To pRs.RecordCount - 1
            feCredInfRiesgo.AdicionaFila
            feCredInfRiesgo.TextMatrix(i + 1, 0) = i + 1
            feCredInfRiesgo.TextMatrix(i + 1, 1) = pRs!Agencia
            feCredInfRiesgo.TextMatrix(i + 1, 2) = pRs!cCtaCod
            feCredInfRiesgo.TextMatrix(i + 1, 3) = PstaNombre(pRs!cPersNombre, True)
            feCredInfRiesgo.TextMatrix(i + 1, 4) = pRs!TpoProd
            feCredInfRiesgo.TextMatrix(i + 1, 5) = pRs!TpoCred
            feCredInfRiesgo.TextMatrix(i + 1, 6) = pRs!Moneda
            feCredInfRiesgo.TextMatrix(i + 1, 7) = Format(pRs!nMonto, "###," & String(15, "#") & "#0.00")
            feCredInfRiesgo.TextMatrix(i + 1, 8) = pRs!FechaExp
            feCredInfRiesgo.TextMatrix(i + 1, 10) = pRs!FechaSaldiaObs
            feCredInfRiesgo.TextMatrix(i + 1, 11) = pRs!FechaReingresoObs
            feCredInfRiesgo.TextMatrix(i + 1, 12) = Format(pRs!ExpRU, "###," & String(15, "#") & "#0.00") 'JOEP20170523 ERS-064
            pRs.MoveNext
        Next i
    End If
    feCredInfRiesgo.TopRow = 1
    Call feCredInfRiesgo_Click 'RECO20161019 ERS060-2016
End Sub
'JOEP-ERS064 20161103

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras(KeyAscii, True)
End Sub

Private Sub txtBuscar_Change()
    Dim rsFiltro As ADODB.Recordset
    
    Dim rsFiltroConCamp As ADODB.Recordset 'JOEP-ERS064 20161104
    Dim rsFiltroSinCamp As ADODB.Recordset 'JOEP-ERS064 20161104
    
    Set rsFiltro = rsCreditoTotal.Clone
    
    
    If fnCampana = 1 Then
        Set rsFiltroConCamp = rsCredConCampTotal.Clone 'JOEP-ERS064 20161104
    ElseIf fnCampana = 2 Then
        Set rsFiltroSinCamp = rsCredSinCampTotal.Clone 'JOEP-ERS064 20161104
    End If
          
    If Trim(txtBuscar.Text) <> "" Then
        If fnCampana = 1 Then 'JOEP-ERS064 20161104
            rsFiltroConCamp.Filter = " cPersNombre LIKE '*" + Trim(txtBuscar.Text) + "*'" 'JOEP-ERS064 20161104
        ElseIf fnCampana = 2 Then 'JOEP-ERS064 20161104
            rsFiltroSinCamp.Filter = " cPersNombre LIKE '*" + Trim(txtBuscar.Text) + "*'" 'JOEP-ERS064 20161104
        Else
            rsFiltro.Filter = " cPersNombre LIKE '*" + Trim(txtBuscar.Text) + "*'"
        End If
    End If
    
    If fnCampana = 1 Then
        Call MostrarListaCampana(rsFiltroConCamp) 'JOEP-ERS064 20161104
    ElseIf fnCampana = 2 Then
        Call MostrarListaCampana(rsFiltroSinCamp) 'JOEP-ERS064 20161104
    Else
        Call MostrarLista(rsFiltro)
    End If
        
End Sub
'WIOR FIN ********
'RECO20161019 ERS060-2016 ************************************
Private Sub ActivaBotones()
    cmdFechExp.Enabled = True
    If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 8) <> "" Then
        If feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 10) <> "" Then
            cmdFechExp.Enabled = False
            cmdSalidaObs.Enabled = False
            cmdRIngresoObs.Enabled = True
        ElseIf feCredInfRiesgo.TextMatrix(feCredInfRiesgo.row, 11) <> "" Then
            cmdFechExp.Enabled = False
            cmdSalidaObs.Enabled = True
            cmdRIngresoObs.Enabled = False
        Else
            cmdFechExp.Enabled = False
            cmdSalidaObs.Enabled = True
            cmdRIngresoObs.Enabled = False
        End If
    Else
        cmdSalidaObs.Enabled = False
        cmdRIngresoObs.Enabled = False
    End If
End Sub
'RECO FIN ****************************************************


