VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmColocCalCargaRcc_CD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga CD Rcc"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "FrmColocCalCargaRcc_CD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Actualizar MaestroRCC"
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtruta 
         Enabled         =   0   'False
         Height          =   330
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   5085
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   330
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   405
      End
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
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   1650
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
         Height          =   360
         Left            =   3120
         TabIndex        =   1
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   435
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmColocCalCargaRcc_CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2

Dim User As String
Dim Pss As String
Dim Ruta As String
Dim FechaTabla As String

Private Sub cmdActualiza_Click()

Dim Co As DConecta
Dim sCn As String
Set Co = New DConecta


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
        oConnection.ConnectionProperties("Initial Catalog") = "dbCmacicaConsol"
        oConnection.ConnectionProperties("Data Source") = "01SRVSICMAC01"
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 2"
        oConnection.id = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = "01SRVSICMAC01"
        oConnection.UserID = User
        oConnection.Password = Pss
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = "dbCmacicaConsol"
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

        oStep.Name = "Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.Description = "Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Task"
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

        oStep.Name = "Copy Data from rcc" & FechaTabla & " to [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.Description = "Copy Data from rcc" & FechaTabla & " to [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from rcc" & FechaTabla & " to [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Task"
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

Set oStep = goPackage.Steps("Copy Data from rcc" & FechaTabla & " to [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step")
        oPrecConstraint.StepName = "Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Create Table [dbCmacicaConsol].[dbo].[rcc200401] Task (Create Table [dbCmacicaConsol].[dbo].[rcc200401] Task)
Call Task_Sub1(goPackage)

'------------- call Task_Sub2 for task Copy Data from rcc200401 to [dbCmacicaConsol].[dbo].[rcc200401] Task (Copy Data from rcc200401 to [dbCmacicaConsol].[dbo].[rcc200401] Task)
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
Dim x As String
x = Cn
Dim i As Integer
Dim z As Integer
z = Len(x)

Dim Temp1 As String
Dim Temp2 As String

Dim Usuario As String

Dim pos1 As Integer
Dim pos2 As Integer

For i = 1 To z
    If Mid(x, i, 7) = "User ID" Then
        pos1 = i
        x = Mid(x, pos1, z)
        Exit For
    End If
Next i

z = Len(x)
For i = 1 To z
    If Mid(x, i, 1) = "=" Then
        pos1 = i
        Exit For
    End If
Next i

For i = 1 To z
    If Mid(x, i, 1) = ";" Then
        pos2 = i
        Exit For
    End If
Next i
Usuario = Mid(x, pos1 + 1, (pos2) - (pos1 + 1))
Usu = Usuario
End Sub
Private Sub ExtraeCadenaPss(ByVal Cn As String, ByRef Ps As String)
Dim x As String
x = Cn
Dim i As Integer
Dim z As Integer
z = Len(x)

Dim Temp1 As String
Dim Temp2 As String

Dim Clave As String

Dim pos1 As Integer
Dim pos2 As Integer

For i = 1 To z
    If UCase(Mid(x, i, 8)) = "PASSWORD" Then
        pos1 = i
        x = Mid(x, pos1, z)
        Exit For
    End If
Next i

z = Len(x)
For i = 1 To z
    If Mid(x, i, 1) = "=" Then
        pos1 = i
        Exit For
    End If
Next i

For i = 1 To z
    If Mid(x, i, 1) = ";" Then
        pos2 = i
        Exit For
    End If
Next i
Clave = Mid(x, pos1 + 1, (pos2) - (pos1 + 1))
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
'------------- define Task_Sub1 for task Create Table [dbCmacicaConsol].[dbo].[rcc200401] Task (Create Table [dbCmacicaConsol].[dbo].[rcc200401] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask1.Description = "Create Table [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask1.SQLStatement = "CREATE TABLE [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[Col001] varchar (255) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub



'------------- define Task_Sub2 for task Copy Data from rcc200401 to [dbCmacicaConsol].[dbo].[rcc200401] Task (Copy Data from rcc200401 to [dbCmacicaConsol].[dbo].[rcc200401] Task)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from rcc" & FechaTabla & " to [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask2.Description = "Copy Data from rcc" & FechaTabla & " to [dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "] Task"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceObjectName = Ruta
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[dbCmacicaConsol].[dbo].[rcc" & FechaTabla & "]"
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

