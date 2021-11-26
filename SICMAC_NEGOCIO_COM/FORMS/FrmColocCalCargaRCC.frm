VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmColocCalCargaRCC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga RCC"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Actualizar MaestroRCC"
      Height          =   660
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Carga CD RCC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   11
         ToolTipText     =   "Primero Cargar el CD del RCC"
         Top             =   750
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   330
         Left            =   5880
         TabIndex        =   10
         ToolTipText     =   "Buscar el Archivo "
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox txtruta 
         Enabled         =   0   'False
         Height          =   330
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   435
      End
   End
   Begin ComctlLib.ProgressBar PrgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1605
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   1905
      Width           =   6375
      Begin VB.CommandButton cmdCargar 
         Caption         =   "&Cargar"
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
         Left            =   2085
         TabIndex        =   13
         Top             =   200
         Width           =   1455
      End
      Begin VB.CommandButton CmdActualiza1 
         Caption         =   "&Actualiza"
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
         Left            =   3690
         TabIndex        =   5
         ToolTipText     =   "Actualliza la Data de la BD"
         Top             =   200
         Width           =   1215
      End
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
         Left            =   4905
         TabIndex        =   4
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   650
      Left            =   150
      TabIndex        =   0
      Top             =   825
      Width           =   2655
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   6
         Format          =   "######"
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Data RCC"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   -960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   615
      Left            =   3300
      Top             =   900
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Este proceso tendra una duracion de 2 a 3 horas. "
      Height          =   600
      Left            =   3450
      TabIndex        =   6
      Top             =   900
      Width           =   2490
   End
End
Attribute VB_Name = "FrmColocCalCargaRCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sServConsol As String
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2

Dim User As String
Dim Pss As String
Dim Ruta As String
Dim FechaTabla As String
Dim BaseConsolidada As String
Dim ServidorInstitucion As String

Dim strNameDTS As String, strServerName As String, strUsuarioSQL As String, strPasswordSQL As String

Dim fcConexConsol As ADODB.Connection

Dim nError As Long
Dim sSource As String, sDesc As String
Dim bExito As Boolean
Dim fsRCCServidor As String, fsRCCBaseDatos As String

Dim objDTS As DTS.Package
Dim objSteps As DTS.Step


Private Sub cmdActualiza_Click()

Dim Co As COMConecta.DCOMConecta
Dim sCn As String
Set Co = New COMConecta.DCOMConecta

Co.AbreConexion
sCn = Co.CadenaConexion
Co.CierraConexion
Set Co = Nothing

Call ExtraeCadenaUsu(sCn, User)
Call ExtraeCadenaPss(sCn, Pss)
If Len(txtruta) = 0 Then
    MsgBox "Ruta no valida", vbInformation, "AVISO"
    Exit Sub
End If
Ruta = Me.txtruta
FechaTabla = Left(Right(Ruta, 10), 6)

        Set goPackage = goPackageOld

        goPackage.Name = "C1"
        goPackage.Description = "DTS package description"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0
        
Dim oConnProperty As DTS.OleDBProperty
'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2
'------------- a new connection defined below.

Set oConnection = goPackage.Connections.New("DTSFlatFile")

        oConnection.ConnectionProperties("Data Source") = Ruta
        
        oConnection.ConnectionProperties("Mode") = 1
        oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
        oConnection.ConnectionProperties("File Format") = 1
        oConnection.ConnectionProperties("Column Delimiter") = "        "
        oConnection.ConnectionProperties("File Type") = 1
        oConnection.ConnectionProperties("Skip Rows") = 0
        oConnection.ConnectionProperties("Text Qualifier") = """"
        oConnection.ConnectionProperties("First Row Column Name") = False
        
        oConnection.Name = "Connection 1"
        oConnection.id = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = Ruta
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = User
        oConnection.ConnectionProperties("Password") = Pss
        oConnection.ConnectionProperties("Initial Catalog") = BaseConsolidada
        oConnection.ConnectionProperties("Data Source") = ServidorInstitucion
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 2"
        oConnection.id = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ServidorInstitucion
        oConnection.UserID = User
        oConnection.Password = Pss
        oConnection.ConnectionTimeout = 60
        'oConnection.Catalog = NOMBRESERVERCONSOL
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.Description = "Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from rcc" & FechaTabla & " to [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.Description = "Copy Data from rcc" & FechaTabla & " to [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from rcc" & FechaTabla & " to [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from rcc" & FechaTabla & " to [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step")
        oPrecConstraint.StepName = "Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Create Table [" & BaseConsolidada & "].[dbo].[rcc200401] Task (Create Table [" & BaseConsolidada & "].[dbo].[rcc200401] Task)
Call Task_Sub1(goPackage)

'------------- call Task_Sub2 for task Copy Data from rcc200401 to [" & BaseConsolidada & "].[dbo].[rcc200401] Task (Copy Data from rcc200401 to [" & BaseConsolidada & "].[dbo].[rcc200401] Task)
Call Task_Sub2(goPackage)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'-****************

goPackage.Execute
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing

mskFecha.Text = FechaTabla
End Sub

Private Sub cmdCargar_Click()
Dim strNameDTS As String, strServerName As String
Dim strUsuarioSQL As String, strPasswordSQL As String

Dim Co As COMConecta.DCOMConecta
Dim sCn As String
Set Co = New COMConecta.DCOMConecta

Co.AbreConexion
sCn = Co.CadenaConexion
Co.CierraConexion
Set Co = Nothing

Call ExtraeCadenaUsu(sCn, User)
Call ExtraeCadenaPss(sCn, Pss)

strServerName = ServidorInstitucion '"01srvSicmac01"
strUsuarioSQL = User  '"sa"
strPasswordSQL = Pss '"cmacica"

'** Fecha de Data a Consolidar
'Set lcCon = New BDConsol.ClsConsolida
'    Me.lblFechaConsol.Caption = Format(lcCon.ObtieneFechaConsolida, "dd/mm/yyyy")
'Set lcCon = Nothing

    strNameDTS = "Carga RCC"

Set objDTS = New DTS.Package

    objDTS.LoadFromSQLServer strServerName, strUsuarioSQL, strPasswordSQL, DTSSQLStgFlag_Default, _
                              , , , strNameDTS
    objDTS.GlobalVariables("gFechaHora").value = Format(gdFecSis, "mm/dd/yyyy")
    objDTS.Execute
    bExito = True
    For Each objSteps In objDTS.Steps
        objSteps.ExecuteInMainThread = True
        If objSteps.ExecutionResult = DTSStepExecResult_Failure Then
            objSteps.GetExecutionErrorInfo nError, sSource, sDesc
            MsgBox "Error en Transferencia :" & objSteps.Description & " " & sDesc, vbExclamation, "Error"
            bExito = False
            Exit For
        End If
    Next
    objDTS.UnInitialize
    Set objDTS = Nothing
End Sub

Private Sub cmdOpen_Click()
 ' Establecer CancelError a True
   
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.InitDir = App.path
    ' Establecer los filtros
    
    CommonDialog1.Filter = "Data Rcc (*.ope)|*.ope|"
    ' Especificar el filtro predeterminado
    'cmdlOpen.FilterIndex = 2
    
    ' Presentar el cuadro de diálogo Abrir
    CommonDialog1.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    txtruta = CommonDialog1.FileName
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub ExtraeCadenaUsu(ByVal Cn As String, ByRef Usu As String)
Dim X As String
X = Cn
Dim I As Integer
Dim z As Integer
z = Len(X)

Dim Temp1 As String
Dim Temp2 As String

Dim Usuario As String

Dim Pos1 As Integer
Dim Pos2 As Integer

For I = 1 To z
    If Mid(X, I, 7) = "User ID" Then
        Pos1 = I
        X = Mid(X, Pos1, z)
        Exit For
    End If
Next I

z = Len(X)
For I = 1 To z
    If Mid(X, I, 1) = "=" Then
        Pos1 = I
        Exit For
    End If
Next I

For I = 1 To z
    If Mid(X, I, 1) = ";" Then
        Pos2 = I
        Exit For
    End If
Next I
Usuario = Mid(X, Pos1 + 1, (Pos2) - (Pos1 + 1))
Usu = Usuario
End Sub

Private Sub ExtraeCadenaPss(ByVal Cn As String, ByRef Ps As String)
Dim X As String
X = Cn
Dim I As Integer
Dim z As Integer
z = Len(X)

Dim Temp1 As String
Dim Temp2 As String

Dim Clave As String

Dim Pos1 As Integer
Dim Pos2 As Integer

For I = 1 To z
    If UCase(Mid(X, I, 8)) = "PASSWORD" Then
        Pos1 = I
        X = Mid(X, Pos1, z)
        Exit For
    End If
Next I

z = Len(X)
For I = 1 To z
    If Mid(X, I, 1) = "=" Then
        Pos1 = I
        Exit For
    End If
Next I

For I = 1 To z
    If Mid(X, I, 1) = ";" Then
        Pos2 = I
        Exit For
    End If
Next I
Clave = Mid(X, Pos1 + 1, (Pos2) - (Pos1 + 1))
Ps = Clave
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("Col001", 1)
                        oColumn.Name = "Col001"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 32
                        oColumn.Size = 255
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Col001", 1)
                        oColumn.Name = "Col001"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 32
                        oColumn.Size = 255
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties
                
        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub
'------------- define Task_Sub1 for task Create Table [" & BaseConsolidada & "].[dbo].[rcc200401] Task (Create Table [" & BaseConsolidada & "].[dbo].[rcc200401] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask1.Description = "Create Table [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask1.SQLStatement = "CREATE TABLE [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Col001] varchar (255) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from rcc200401 to [" & BaseConsolidada & "].[dbo].[rcc200401] Task (Copy Data from rcc200401 to [" & BaseConsolidada & "].[dbo].[rcc200401] Task)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from rcc" & FechaTabla & " to [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask2.Description = "Copy Data from rcc" & FechaTabla & " to [" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceObjectName = Ruta
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[" & BaseConsolidada & "].[dbo].[rcc" & FechaTabla & "]"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0
        
Call oCustomTask2_Trans_Sub1(oCustomTask2)
                
goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Private Sub cmdActualiza1_Click()
'Dim loRcc As COMDCredito.DCOMColocEval
'Dim rs As ADODB.Recordset
'Dim opt As Integer
'Set loRcc = New COMDCredito.DCOMColocEval
'
''Verifica si se migro el cd del loRcc a la Fecha
'If loRcc.VerificaTabla(sServConsol, Me.mskFecha) = -1 Then
'    MsgBox "No Existe la tabla loRcc" & mskFecha, vbCritical, "AVISO"
'    Set loRcc = Nothing
'    Exit Sub
'End If
'
'If Format(LblRcc1, "YYYYMM") = Me.mskFecha Then
'    MsgBox "Ya se migro a esa Fecha", vbInformation, "Aviso"
'    Set loRcc = Nothing
'    Exit Sub
'End If
'
''Verificar las Fechas de las tablas a Actulizar
'Set rs = loRcc.FechasRcc(sServConsol)
'While Not rs.EOF
'    If Format(rs!Fecha, "YYYYMM") = Trim(mskFecha.Text) Then
'        MsgBox "Ya se realizo la migracion con esa Fecha", vbCritical, "Aviso"
'        Set loRcc = Nothing
'        Set rs = Nothing
'        Exit Sub
'    End If
'    rs.MoveNext
'Wend
'
'opt = MsgBox("Desea Actualizar el Rcc a " & Format(Left(Me.mskFecha, 4) & "/" & Right(Me.mskFecha, 2), "MMMM") & " del " & Left(Me.mskFecha, 4), vbInformation + vbYesNo, "AVISO")
'
'If opt = vbNo Then
'    Set rs = Nothing
'    Set loRcc = Nothing
'    Exit Sub
'End If
'
'PrgBar.Min = 0
'PrgBar.Max = 100
'
'PrgBar.value = 10
''Transfiere la data del RccTotal y RccTotalDet hacia una Tabla Historica
''Call loRcc.InsertaRCChistorico(sServConsol)
'Call loRcc.InsertaRccHistorico(fcConexConsol)
'DoEvents
'PrgBar.value = 45
'
'
''Blanquea la Tabla RccTotal  y RccTotalDet
'Call loRcc.BorraTablaRCC(sServConsol)
'
'DoEvents
'PrgBar.value = 60
'
''Inserta la Data Nueva
'Call loRcc.InsertaRccCab(sServConsol, Me.mskFecha)
'PrgBar.value = 65
'DoEvents
'Call loRcc.InsertRCCDet(sServConsol, Me.mskFecha)
'PrgBar.value = 85
'DoEvents
'
''LblRcc1 = loRcc.GetFecha(sServConsol, 1)
''LblRcc2 = loRcc.GetFecha(sServConsol, 3)
'
'PrgBar.value = 100
'DoEvents
'
'MsgBox "Proceso Terminado", vbInformation, "AVISO"
'Set loRcc = Nothing
'Set rs = Nothing
End Sub

Private Sub Form_Load()
Dim Rcc As COMDCredito.DCOMColocEval
Dim Co As COMConecta.DCOMConecta
Dim lsCadConexConsol As String

Set Rcc = New COMDCredito.DCOMColocEval
    sServConsol = Rcc.ServConsol(gConstSistServCentralRiesgos)
    BaseConsolidada = Trim(Rcc.NombreServerConsol)
    'LblRcc1 = Rcc.GetFecha(sServConsol, 1)
    'LblRcc2 = Rcc.GetFecha(sServConsol, 3)
Set Rcc = Nothing

Set Co = New COMConecta.DCOMConecta
    Co.AbreConexion
    ServidorInstitucion = Trim(Co.ServerName)
    Co.CierraConexion
Set Co = Nothing

'Conexion al servidor consol --- LAYG
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lsConsolServidor As String, lsConsolBaseDatos As String

Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lsConsolServidor = loConstSis.LeeConstSistema(145)
    lsConsolBaseDatos = loConstSis.LeeConstSistema(146)
    fsRCCServidor = loConstSis.LeeConstSistema(143)
    fsRCCBaseDatos = loConstSis.LeeConstSistema(144)
Set loConstSis = Nothing

lsCadConexConsol = fgCadenaConexConsol(lsConsolServidor, lsConsolBaseDatos)
Set Co = New COMConecta.DCOMConecta
    Co.AbreConexion lsCadConexConsol
    Set fcConexConsol = Co.ConexionActiva
    Co.CierraConexion
Set Co = Nothing

End Sub
