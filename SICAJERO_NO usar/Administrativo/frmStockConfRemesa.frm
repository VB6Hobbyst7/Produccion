VERSION 5.00
Begin VB.Form frmStockConfRemesa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmación de Remesa"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
   Icon            =   "frmStockConfRemesa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   15
      TabIndex        =   15
      Top             =   5880
      Width           =   11070
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   420
         Left            =   9480
         TabIndex        =   8
         Top             =   240
         Width           =   1410
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   420
         Left            =   90
         TabIndex        =   7
         Top             =   195
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4680
      Left            =   30
      TabIndex        =   14
      Top             =   1200
      Width           =   11025
      Begin VB.ListBox lstRemesasXConf 
         Height          =   1860
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   600
         Width           =   10575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Glosa"
         Height          =   1455
         Left            =   135
         TabIndex        =   16
         Top             =   3150
         Width           =   10665
         Begin VB.TextBox txtGlosa 
            Height          =   1065
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   240
            Width           =   10335
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9600
         TabIndex        =   23
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7680
         TabIndex        =   22
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Area Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Area Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Busqueda"
      Height          =   1095
      Left            =   15
      TabIndex        =   9
      Top             =   30
      Width           =   11040
      Begin VB.OptionButton Option2 
         Caption         =   "Devolucion"
         Height          =   315
         Left            =   4875
         TabIndex        =   26
         Top             =   615
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Enabled         =   0   'False
         Height          =   585
         Left            =   9840
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtFechaFinal 
         Height          =   300
         Left            =   7995
         TabIndex        =   3
         Text            =   "30/11/2008"
         Top             =   615
         Width           =   1125
      End
      Begin VB.TextBox txtFechaInicial 
         Height          =   300
         Left            =   7995
         TabIndex        =   2
         Text            =   "20/11/2008"
         Top             =   270
         Width           =   1125
      End
      Begin VB.ComboBox cboDestino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   615
         Width           =   3135
      End
      Begin VB.ComboBox cboOrigen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Remesa"
         Height          =   270
         Left            =   4830
         TabIndex        =   25
         Top             =   270
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Final   :"
         Height          =   240
         Left            =   6900
         TabIndex        =   13
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Inicial :"
         Height          =   240
         Left            =   6915
         TabIndex        =   12
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Destino :"
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Origen   :"
         Height          =   285
         Left            =   390
         TabIndex        =   10
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Nro. Remesa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmStockConfRemesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta


Private Sub cmdConfirmar_Click()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim v As Integer, i As Integer

    oConec.AbreConexion
    
    For i = 0 To Me.lstRemesasXConf.ListCount - 1
    If Me.lstRemesasXConf.Selected(i) = True Then
        
         Set Prm = New ADODB.Parameter
         Set Prm = Cmd.CreateParameter("@NumTran", adInteger, adParamInput, , Mid(lstRemesasXConf.List(i), 1, InStr(lstRemesasXConf.List(i), "_") - 1))
         Cmd.Parameters.Append Prm
                  
         Set Prm = New ADODB.Parameter
         Set Prm = Cmd.CreateParameter("@cObservacion", adVarChar, adParamInput, 100, txtGlosa.Text)
         Cmd.Parameters.Append Prm
        
         
         Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
         Cmd.CommandType = adCmdStoredProc
         Cmd.CommandText = "ATM_ActualizaEstadoRemesas"
                
         Call Cmd.Execute
         Set Prm = Nothing
         Set Cmd = Nothing
                        
        End If
    Next
    Me.txtGlosa.Text = ""
    MsgBox "Remesas se actualizaron Correctamente", vbInformation, "Confirmacion de Remesa"
    
    oConec.CierraConexion
    
    Call CargaDatosGridView
    
End Sub

Private Sub CmdProcesar_Click()
        
    If Not IsDate(Me.txtFechaInicial.Text) Then
        MsgBox "Fecha Inicial Incorrecta", vbExclamation, "Validacion"
        Exit Sub
    End If
    If Not IsDate(Me.txtFechaFinal.Text) Then
        MsgBox "Fecha Final Incorrecta", vbExclamation, "Validacion"
        Exit Sub
    End If
    

Call CargaDatosGridView

If Me.lstRemesasXConf.ListCount > 0 Then
    Me.cmdConfirmar.Enabled = True
Else
    Me.cmdConfirmar.Enabled = False
End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    txtFechaInicial.Text = gdFecSis
    txtFechaFinal.Text = gdFecSis
    Call CargaDatos
    'Call CargaDatos2
    'cboOrigen.Text = cboOrigen.List(0)
    Me.cmdConfirmar.Enabled = False
    
    CmdProcesar.Enabled = True
    cboOrigen.Text = cboOrigen.List(0)
    cboDestino.List(1) = gsCodAge & "-" & gsNomAge
    cboDestino.Text = cboDestino.List(1)
    cboOrigen.Enabled = False
    cboDestino.Enabled = False
    Me.lstRemesasXConf.Clear
    Me.cmdConfirmar.Enabled = False
    
    
End Sub

Private Sub CargaDatosGridView()
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter
Dim R As ADODB.Recordset

        Set R = New ADODB.Recordset
        
        oConec.AbreConexion
        Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        
        If (Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1)) = (Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1)) Then
            MsgBox "Los valores de Origen y Destino deben ser diferentes"
            Exit Sub
        End If
    
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nOrigen", adInteger, adParamInput, 8, Mid(cboOrigen.Text, 1, InStr(cboOrigen.Text, "-") - 1))
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nDestino", adInteger, adParamInput, 8, Mid(cboDestino.Text, 1, InStr(cboDestino.Text, "-") - 1))
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFechaInicial", adDate, adParamInput, 16, txtFechaInicial.Text)
        Cmd.Parameters.Append Prm
        
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFechaFinal", adDate, adParamInput, 16, txtFechaFinal.Text)
        Cmd.Parameters.Append Prm
        
        Cmd.CommandText = "ATM_RecuperaRemesasXConfirmar"

        R.CursorType = adOpenStatic
        R.LockType = adLockReadOnly
        Set R = Cmd.Execute
        lstRemesasXConf.Clear
        Do While Not R.EOF
            lstRemesasXConf.AddItem Left(Space(5) & R!NroRemesa & "_" & Space(13), 15) & Left(R!Fecha & Space(8), 25) & Left(R!AreaOrigen & Space(20), 20) & Left(R!AreaDestino & Space(25), 25) & Left(R!Cantidad & Space(25), 25) & Left(R!Del & Space(25), 25) & Left(R!Al & Space(20), 20)
            R.MoveNext
        Loop
        Me.cmdConfirmar.Enabled = True
        If Me.lstRemesasXConf.ListCount = 0 Then
            cmdConfirmar.Enabled = False
        End If
        'CmdAct.Enabled = True
        oConec.CierraConexion
        
End Sub


Private Sub CargaDatos()
 
    Dim Cmd As New Command
    Dim Prm As New ADODB.Parameter
    Dim R As ADODB.Recordset
    Set Prm = New ADODB.Parameter
    'Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 50, LblNumTarj.Caption)
    'Cmd.Parameters.Append Prm

    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    Cmd.CommandText = "ATM_RecuperaAgencias"
    
        
    Set R = Cmd.Execute
    'CboCtas.Clear
    Do While Not R.EOF
         cboOrigen.AddItem R!Codigo & "-" & R!Agencia
        cboDestino.AddItem R!Codigo & "-" & R!Agencia
       R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    oConec.CierraConexion
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

'Private Sub CargaDatos2()
'    cboDestino.List(0) = gsCodAge & "-" & gsNomAge
'    cboDestino.Text = cboDestino.List(0)
'End Sub

Private Sub Option1_Click()
    CmdProcesar.Enabled = True
    cboOrigen.Text = cboOrigen.List(0)
    cboDestino.List(1) = gsCodAge & "-" & gsNomAge
    cboDestino.Text = cboDestino.List(1)
    cboOrigen.Enabled = False
    cboDestino.Enabled = False
    Me.lstRemesasXConf.Clear
    Me.cmdConfirmar.Enabled = False
End Sub

Private Sub Option2_Click()
    CmdProcesar.Enabled = True
    cboOrigen.Enabled = True
    cboDestino.Enabled = False
    'cboOrigen.List(1) = gsCodAge & "-" & gsNomAge
    cboOrigen.Text = cboOrigen.List(1)
    cboDestino.Text = cboDestino.List(0)
    Me.lstRemesasXConf.Clear
    Me.cmdConfirmar.Enabled = False
End Sub
