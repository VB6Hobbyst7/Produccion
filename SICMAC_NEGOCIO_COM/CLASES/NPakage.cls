VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPakage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents oPackage As DTS.Package
Attribute oPackage.VB_VarHelpID = -1
Public bCancel As Boolean
Public sCadena As String
Public nEstado As Integer
' Valores de nEstado
'1 Empieza
'2 En Progreso
'3 Cancelado
'4 Finalizado
'5 Error
Private Sub EjecutarDTS(ByVal oPkg As DTS.Package)
    Dim nStepCount As Integer
    For nStepCount = 1 To oPkg.Steps.Count
        oPkg.Steps(nStepCount).ExecuteInMainThread = True
    Next nStepCount
End Sub

Private Sub oPackage_OnError(ByVal EventSource As String, _
                                ByVal ErrorCode As Long, _
                                ByVal Source As String, _
                                ByVal Description As String, _
                                ByVal HelpFile As String, _
                                ByVal HelpContext As Long, _
                                ByVal IDofInterfaceWithError As String, _
                                ByRef pbCancel As Boolean)
     sCadena = sCadena & " " & ErrorCode & " " & Description & vbCrLf
     nEstado = 5
End Sub

Private Sub oPackage_OnFinish(ByVal EventSource As String)
    sCadena = sCadena & " " & EventSource & vbCrLf
    nEstado = 4
End Sub

Private Sub oPackage_onProgress(ByVal EventSource As String, _
                                  ByVal ProgressDescription As String, _
                                  ByVal PercentComplete As Long, _
                                  ByVal ProgressCountLow As Long, _
                                  ByVal ProgressCountHigh As Long)
    sCadena = sCadena & " " & ProgressDescription & " " & PercentComplete & vbCrLf
    nEstado = 2
End Sub
                                  

Private Sub oPackage_OnQueryCancel(ByVal EventSource As String, _
                                    ByRef pbCancel As Boolean)
' El paquete DTS activará este evento a determinados intervalos para comprobar
' si la ejecución del paquete debe finalizar. Establezca
' pbCancel en true para cancelar la ejecución del paquete.
'Debug.Print "oPackage_OnQueryCancel Fired"
'    If bCancel Then
'        pbCancel = True
'        Debug.Print "Canceling package execution."
'    Else
'        pbCancel = False
'    End If
If nEstado = 5 Then
      pbCancel = True
End If
nEstado = 3
End Sub

Private Sub oPackage_OnStart(ByVal EventSource As String)
    'Debug.Print "oPackage_OnStart Fired"
    nEstado = 1
End Sub

Public Sub ExecDTS(ByVal pstrNomDTS As String)
    Dim oTask As DTS.DataPumpTask

'    Dim strFileName As String
 '   Dim strFilePassword As String

    bCancel = False

    Set oPackage = New DTS.Package
    oPackage.LoadFromSQLServer "01srvsicmac01", , , DTSSQLStgFlag_UseTrustedConnection, , , , pstrNomDTS
    'strFileName = "C:\temp\test.dts"
    'strFilePassword = "VerySecurePassword" 'Usar si el archivo tiene la contraseña
    
    'oPackage.LoadFromStorageFile strFileName, strFilePassword
    'For x = 1 To oPackage.Tasks.Count
    ' Debug.Print oPackage.Tasks.Item(x)
    'Next x
    ' Sólo llame a lo siguiente cuando desarrolle una aplicación. Puede
    ' agregar comentarios a la llamada cuando cree la aplicación.
    
    EjecutarDTS oPackage
    
    Set oTask = oPackage.Tasks.item("DTSTask_DTSDataPumpTask_1").CustomTask
    
    oTask.ProgressRowCount = 1
    
    'Screen.MousePointer = vbHourglass
    
    oPackage.Execute
    
    'Screen.MousePointer = vbNormal
    
    Set oTask = Nothing
    Set oPackage = Nothing
End Sub

