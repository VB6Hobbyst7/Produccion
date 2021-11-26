VERSION 5.00
Begin VB.Form frmLogVehiculoTipoReg 
   ClientHeight    =   4155
   ClientLeft      =   1425
   ClientTop       =   2775
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7965
   Begin VB.Frame fraReg 
      Caption         =   "Tipos de Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3945
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7845
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   480
         Width           =   6195
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5100
         TabIndex        =   14
         Top             =   3420
         Width           =   1185
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   6360
         TabIndex        =   13
         Top             =   3420
         Width           =   1185
      End
      Begin VB.CheckBox chkEtiqueta 
         Caption         =   "Etiqueta Descripcion / Persona"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   12
         Top             =   1260
         Width           =   2595
      End
      Begin VB.CheckBox chkSeleY 
         Caption         =   "Solicitar Documento"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2700
         Width           =   1935
      End
      Begin VB.CheckBox chkSele0 
         Caption         =   "Selector de Persona"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1260
         Width           =   1995
      End
      Begin VB.CheckBox chkSele1 
         Caption         =   "Selector de UbiGeo 1"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1620
         Width           =   1995
      End
      Begin VB.CheckBox chkSele2 
         Caption         =   "Selector de Ubigeo 2"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1980
         Width           =   1995
      End
      Begin VB.CheckBox chkSeleX 
         Caption         =   "Campo Descripción"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   2340
         Width           =   1935
      End
      Begin VB.TextBox txtEtiqueta 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   1995
      End
      Begin VB.CheckBox chkEtiqueta 
         Caption         =   "Etiqueta Selector de UbiGeo 1"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   1620
         Width           =   2535
      End
      Begin VB.CheckBox chkEtiqueta 
         Caption         =   "Etiqueta Selector de Ubigeo 2"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   1980
         Width           =   2535
      End
      Begin VB.TextBox txtEtiqueta 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   1995
      End
      Begin VB.TextBox txtEtiqueta 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1920
         Width           =   1995
      End
      Begin VB.CommandButton cmdCtrl 
         Caption         =   "Grabar distribución de Controles"
         Height          =   375
         Left            =   2820
         TabIndex        =   1
         Top             =   2520
         Width           =   4665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registro"
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
         Left            =   300
         TabIndex        =   15
         Top             =   540
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmLogVehiculoTipoReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
CentraForm Me
CargaTipos
End Sub

Sub CargaTipos()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod =9024 and nConsCod<>nConsValor"
If oConn.AbreConexion Then
Set rs = oConn.CargaRecordSet(sSQL)

If Not rs.EOF Then
   Do While Not rs.EOF
      cboTipo.AddItem rs!cConsDescripcion
      cboTipo.ItemData(cboTipo.ListCount - 1) = rs!nConsValor
      rs.MoveNext
   Loop
End If
End If
End Sub

'/*************************************************
Private Sub chkEtiqueta_Click(Index As Integer)
If chkEtiqueta(Index).value = 1 Then
   txtEtiqueta(Index).BackColor = "&H80000005"
   txtEtiqueta(Index).Locked = False
Else
   txtEtiqueta(Index).BackColor = "&H8000000F"
   txtEtiqueta(Index).Locked = True
End If
End Sub

Private Sub cmdCtrl_Click()
Dim oConn As New DConecta
Dim nCod As Integer, sSQL As String

If cboTipo.ListIndex < 0 Then Exit Sub

nCod = cboTipo.ItemData(cboTipo.ListIndex)

If MsgBox("Esta seguro de Grabar", vbQuestion + vbYesNo, "AVISO") = vbYes Then
   If oConn.AbreConexion Then
   
          sSQL = "DELETE FROM LogVehiculoCtrl Where nTipoReg = " & nCod & " "
          oConn.Ejecutar sSQL
      
            
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'lblValor0','" & txtEtiqueta(0).Text & "'," & chkEtiqueta(0).value & ")"
          oConn.Ejecutar sSQL
          
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'lblValor1','" & txtEtiqueta(1).Text & "'," & chkEtiqueta(1).value & ")"
          oConn.Ejecutar sSQL
          
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'lblValor2','" & txtEtiqueta(2).Text & "'," & chkEtiqueta(2).value & ")"
          oConn.Ejecutar sSQL
      
         ' Selector de Persona ---------------------------------------------------
         
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtValor0',''," & chkSele0.value & ")"
          oConn.Ejecutar sSQL
          
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'cmdBusq0',''," & chkSele0.value & ")"
          oConn.Ejecutar sSQL
         
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtValDesc0',''," & chkSele0.value & ")"
          oConn.Ejecutar sSQL
      
         ' Selector de Ubigeo 1 ---------------------------------------------------
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtValor1',''," & chkSele1.value & ")"
          oConn.Ejecutar sSQL
          
          If nCod = 4 Then
             sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                    "   VALUES (" & nCod & ",'cmdBusq1','',0)"
             oConn.Ejecutar sSQL
         
             sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                    "   VALUES (" & nCod & ",'txtValDesc1','',0)"
             oConn.Ejecutar sSQL
          Else
             sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                    "   VALUES (" & nCod & ",'cmdBusq1',''," & chkSele1.value & ")"
             oConn.Ejecutar sSQL
         
             sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                    "   VALUES (" & nCod & ",'txtValDesc1',''," & chkSele1.value & ")"
             oConn.Ejecutar sSQL
          End If
      
         ' Selector de Ubigeo 2 ---------------------------------------------------
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtValor2',''," & chkSele2.value & ")"
          oConn.Ejecutar sSQL
          
          If nCod = 4 Then
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'cmdBusq2','',0)"
          oConn.Ejecutar sSQL
         
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtValDesc2','',0)"
          oConn.Ejecutar sSQL
          
          Else
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'cmdBusq2',''," & chkSele2.value & ")"
          oConn.Ejecutar sSQL
         
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtValDesc2',''," & chkSele2.value & ")"
          oConn.Ejecutar sSQL
          End If
         ' Campo Descripcion ---------------------------------------------------
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'txtDescripcion',''," & chkSeleX.value & ")"
          oConn.Ejecutar sSQL
      
         ' Campo Documento   ---------------------------------------------------
          sSQL = "INSERT INTO LogVehiculoCtrl (nTipoReg,cCtrlName,cCtrlText,nVisible) " & _
                 "   VALUES (" & nCod & ",'fraDoc',''," & chkSeleY.value & ")"
          oConn.Ejecutar sSQL
      
      oConn.CierraConexion
      Unload Me
   End If
End If
End Sub

