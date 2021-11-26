VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmConsolidado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Consolidada"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3420
      TabIndex        =   2
      Top             =   4020
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe"
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
      Height          =   3810
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   7290
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msh 
         Height          =   3285
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5794
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   2220
      TabIndex        =   0
      Top             =   4020
      Width           =   1050
   End
   Begin VB.Image ImgX 
      Height          =   375
      Left            =   1050
      Picture         =   "FrmConsolidado.frx":0000
      Stretch         =   -1  'True
      Top             =   3990
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image ImgQuestion 
      Height          =   375
      Left            =   570
      Picture         =   "FrmConsolidado.frx":02B5
      Stretch         =   -1  'True
      Top             =   3990
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image ImgBien 
      Height          =   375
      Left            =   120
      Picture         =   "FrmConsolidado.frx":05B8
      Stretch         =   -1  'True
      Top             =   3990
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "FrmConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nposFinal As Integer
Private Sub CmdAceptar_Click()
    Dim bEstado As Boolean
 
 If MsgBox("Esta seguro que desea tranferir data", vbInformation + vbYesNo, "AVISO") = vbYes Then
    'PB.Visible = True
    'Timer1.Enabled = True
    With Msh
    .TextMatrix(1, 0) = "1"
    .TextMatrix(1, 1) = "Transferencia a Consolidada"
    .TextMatrix(1, 2) = "Procesando..."
    
'    .ColWidth(0) = 700
'    .ColWidth(1) = 2000
'    .ColWidth(2) = 3500
'    .Row = 1
'    .Col = 3
'    Set .CellPicture = ImgQuestion.Picture
'    .CellPictureAlignment = flexAlignCenterCenter
'    .Row = 2
'    .Col = 3
'    Set .CellPicture = ImgQuestion.Picture
'    .CellPictureAlignment = flexAlignCenterCenter
 End With
    bEstado = EjeuctarPackage_ProcesoEjecucion
    If bEstado = True Then
        With Msh
            .TextMatrix(nposFinal, 0) = "1"
            .TextMatrix(nposFinal, 1) = "Calculos Consolidada"
            .TextMatrix(nposFinal, 2) = "Procesando..."
        End With
        bEstado = Ejecutar_PakageCalculosConsolidada
        'Timer1.Enabled = False
        'PB.Visible = False
    Else
        'Timer1.Enabled = False
        'PB.Visible = False
    End If
 End If
End Sub

Private Sub cmdCancelar_Click()
 Configurar_MSH
 Msh.Cols = 4
 Msh.Rows = 3
 'Timer1.Enabled = False
 'PB.Visible = False
End Sub

Private Sub Form_Load()
    Configurar_MSH
  '  PB.Min = 0
   ' PB.Max = 100
    'PB.value = 0
    'PB.Visible = False
End Sub
Sub Configurar_MSH()
 With Msh
    .TextMatrix(0, 0) = "Nro.DTS"
    .TextMatrix(0, 1) = "Nombre DTS"
    .TextMatrix(0, 2) = "Proceso en Ejecucion"
    .TextMatrix(0, 3) = "Estado"
    
    .TextMatrix(1, 0) = "1"
    .TextMatrix(2, 0) = "2"
    
    .TextMatrix(1, 1) = "Transferencia a Consolidada"
    .TextMatrix(2, 1) = "Calculos Consolidada"
    .TextMatrix(1, 2) = "Espera..."
    .TextMatrix(2, 2) = "Espera..."
    .ColWidth(0) = 700
    .ColWidth(1) = 2000
    .ColWidth(2) = 3500
    .Row = 1
    .Col = 3
    Set .CellPicture = ImgQuestion.Picture
    .CellPictureAlignment = flexAlignCenterCenter
    .Row = 2
    .Col = 3
    Set .CellPicture = ImgQuestion.Picture
    .CellPictureAlignment = flexAlignCenterCenter
 End With
End Sub

Function EjeuctarPackage_ProcesoEjecucion() As Boolean
    Dim objPkg As DTS.Package
    Dim strServerName As String
    Dim strUserSQL As String
    Dim strPasswordSQL As String
    Dim strComentario As String
    Dim objSteps As DTS.Step
    Dim i As Integer
    Dim nError As Long
    Dim sSource As String
    Dim sDesc As String
    Dim bEstado As Boolean
        On Error GoTo ErrHandler
    
    
    
        
        strServerName = ObtenerServidorDatos
        strUserSQL = ObtenerUsuarioDatos
        strPasswordSQL = ObtenerPassword
        
        Set objPkg = New DTS.Package
        objPkg.LoadFromSQLServer strServerName, strUserSQL, strPasswordSQL, DTSSQLStgFlag_UseTrustedConnection, _
                    , , , "Transferencia a Consolidada"
        

         objPkg.Execute
        i = 1
        For Each objSteps In objPkg.Steps
            i = i + 1
            Msh.Rows = Msh.Rows + 1
           objSteps.ExecuteInMainThread = True
         

            If objSteps.ExecutionResult = DTSStepExecResult_Success Then
               Msh.TextMatrix(i, 0) = i
               Msh.TextMatrix(i, 1) = "Transferencia a Consolidada"
               Msh.TextMatrix(i, 2) = "OK:" & objSteps.Description & " " & objSteps.ExecutionTime
               Msh.Col = 3
               Msh.Row = i
               Set Msh.CellPicture = ImgBien.Picture
               Msh.CellPictureAlignment = flexAlignCenterCenter
               bEstado = True
            Else
                objSteps.GetExecutionErrorInfo nError, sSource, sDesc
                Msh.TextMatrix(i, 0) = i
                Msh.TextMatrix(i, 1) = "Transferencia a Consolidada"
                Msh.TextMatrix(i, 2) = "FAIL:" & objSteps.Description & " " & objSteps.ExecutionTime
                Msh.Col = 3
                Msh.Row = i
                Set Msh.CellPicture = ImgX.Picture
                Msh.CellPictureAlignment = flexAlignCenterCenter
                bEstado = False
            End If
        Next
        'Configurando de nuevo el primer dts
        ' Configurando de nuevo el segundo DTS
                Msh.TextMatrix(1, 0) = 1
                Msh.TextMatrix(1, 1) = "Transferencia a Consolidada"
                Msh.TextMatrix(1, 2) = "Proceso terminado..."
                Msh.Col = 3
                Msh.Row = 1
                If bEstado = True Then
                    Set Msh.CellPicture = ImgBien.Picture
                Else
                    Set Msh.CellPicture = ImgX.Picture
                End If
                Msh.CellPictureAlignment = flexAlignCenterCenter
        
        i = i + 1
                nposFinal = i
                Msh.TextMatrix(i, 0) = i
                Msh.TextMatrix(i, 1) = "Calculos Consolidada"
                Msh.TextMatrix(i, 2) = "Esperando..."
                Msh.Col = 3
                Msh.Row = i
                Set Msh.CellPicture = ImgQuestion.Picture
                Msh.CellPictureAlignment = flexAlignCenterCenter
                
        objPkg.UnInitialize
        Set objSteps = Nothing
        Set objPkg = Nothing
        EjeuctarPackage_ProcesoEjecucion = True
    Exit Function
ErrHandler:
    If Not objPkg Is Nothing Then Set objPkg = Nothing
    EjeuctarPackage_ProcesoEjecucion = False
End Function

Function Ejecutar_PakageCalculosConsolidada() As Boolean
    Dim objPkg As DTS.Package
    Dim strServerName As String
    Dim strUserSQL As String
    Dim strPasswordSQL As String
    Dim strComentario As String
    Dim objSteps As DTS.Step
    Dim i As Integer
    Dim nError As String
    Dim sSource As String
    Dim sDesc As String
    Dim bEstado As Boolean
    On Error GoTo ErrHandler
    strServerName = ObtenerServidorDatos
        strUserSQL = ObtenerUsuarioDatos
        strPasswordSQL = ObtenerPassword
        
        Set objPkg = New DTS.Package
        objPkg.LoadFromSQLServer strServerName, strUserSQL, strPasswordSQL, DTSSQLStgFlag_UseTrustedConnection, _
                    , , , "Calculos Consolidada"
        objPkg.Execute
        i = Msh.Rows
        For Each objSteps In objPkg.Steps
            i = i + 1
            Msh.Rows = Msh.Rows + 1
           ' objSteps.Execute
            objSteps.ExecuteInMainThread = True
            If objSteps.ExecutionResult = DTSStepExecResult_Success Then
               Msh.TextMatrix(i - 1, 0) = i
               Msh.TextMatrix(i - 1, 1) = "Calculos Consolidada"
               Msh.TextMatrix(i - 1, 2) = "OK:" & objSteps.Description & " " & objSteps.ExecutionTime
               Msh.Col = 3
               Msh.Row = i - 1
               Set Msh.CellPicture = ImgBien.Picture
               Msh.CellPictureAlignment = flexAlignCenterCenter
               bEstado = True
            Else
                'objSteps.GetExecutionErrorInfo nError, sSource, sDesc
                Msh.TextMatrix(i - 1, 0) = i - 1
                Msh.TextMatrix(i - 1, 1) = "Calculos Consolidada"
                Msh.TextMatrix(i - 1, 2) = "FAIL:" & objSteps.Description & " " & objSteps.ExecutionTime
                Msh.Col = 3
                Msh.Row = i - 1
                Set Msh.CellPicture = ImgX.Picture
                Msh.CellPictureAlignment = flexAlignCenterCenter
                bEstado = False
            End If
        Next
                i = nposFinal
                Msh.TextMatrix(i, 0) = i
                Msh.TextMatrix(i, 1) = "Calculos Consolidada"
                Msh.TextMatrix(i, 2) = "Proceso terminado.."
                Msh.Col = 3
                Msh.Row = i
                If bEstado = True Then
                    Set Msh.CellPicture = ImgBien.Picture
                Else
                    Set Msh.CellPicture = ImgX.Picture
                End If
                Msh.CellPictureAlignment = flexAlignCenterCenter
                
        
        objPkg.UnInitialize
        Set objSteps = Nothing
        Set objPkg = Nothing
        Ejecutar_PakageCalculosConsolidada = True
    Exit Function
ErrHandler:
    If Not objPkg Is Nothing Then Set objPkg = Nothing
    Ejecutar_PakageCalculosConsolidada = False
End Function
Function ObtenerServidorDatos() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPos As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPos = InStr(1, strCadenaConexion, "Data Source=")
    intPosF = InStr(intPos, strCadenaConexion, ";")
    ObtenerServidorDatos = Mid(strCadenaConexion, intPos + Len("Data Source="), intPosF - (intPos + Len("Data Source=")))
    
Set oConec = Nothing
End Function

Function ObtenerUsuarioDatos() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPos As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPos = InStr(1, strCadenaConexion, "User ID=")
    intPosF = InStr(intPos, strCadenaConexion, ";")
    ObtenerUsuarioDatos = Mid(strCadenaConexion, intPos + Len("User ID="), intPosF - (intPos + Len("User ID=")))
Set oConec = Nothing
End Function

Function ObtenerPassword() As String
Dim oConec As DConecta
Dim strCadenaConexion As String
Dim intPosI As Integer
Dim intPosF As Integer
Set oConec = New DConecta
    oConec.AbreConexion
    strCadenaConexion = oConec.CadenaConexion
    oConec.CierraConexion
    intPosI = InStr(1, strCadenaConexion, "Password")
    intPosF = InStr(intPosI, strCadenaConexion, ";")
    ObtenerPassword = Mid(strCadenaConexion, intPosI + Len("Password="), intPosF - (intPosI + Len("Password=")))
    Set oConec = Nothing
End Function

'Private Sub Timer1_Timer()
'    Static i As Integer
'        If i = 100 Then
'            i = 0
'        End If
'        i = i + 1
'        PB.value = i
'End Sub
