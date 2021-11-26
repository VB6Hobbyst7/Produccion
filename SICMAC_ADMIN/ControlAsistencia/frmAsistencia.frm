VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAsistencia 
   Caption         =   "Control de Asistencia"
   ClientHeight    =   3180
   ClientLeft      =   2805
   ClientTop       =   4110
   ClientWidth     =   8775
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Icon            =   "frmAsistencia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   8775
   Begin VB.Timer timerControlDia 
      Interval        =   60000
      Left            =   555
      Top             =   1260
   End
   Begin VB.CommandButton cmdAyuda 
      Caption         =   "&Ayuda"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   105
      Top             =   1260
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdAvazados 
      Caption         =   "&Avanzado >>>"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "&Ver"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "&Iniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblmensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   8625
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblFecha 
      Caption         =   "01/01/2001 08:08:08 AM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "frmAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnFinalizar As Long

Dim lnMayor As Long
Dim lnMenor As Long
Dim lsTextoMayor As String
Dim lsTextoMenor As String
Dim lsPass As String
Dim ldFecha As Date

Private Sub cmdAvazados_Click()
'    If lsTextoMayor = cmdAvazados.Caption Then
'        If Not frmClave.GetClave Then
'            MsgBox "Clave Errada.", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        cmdAvazados.Caption = lsTextoMenor
'        Width = lnMayor
'        cmdVer.Enabled = True
'    Else
'        cmdAvazados.Caption = lsTextoMayor
'        Width = lnMenor
'        cmdVer.Enabled = False
'    End If
End Sub

Private Sub cmdAyuda_Click()
    MsgBox "Para Iniciar el Control de Asistencia debe ingresar el turno e indicar si es ingreso o salida. Tambien verifique la fecha del Programa, si no es la del dia modifiquela y presione ENTER, guardela si asi lo desea, luego presione el incono INICIAR", vbInformation, "Aviso"
End Sub

Private Sub cmdIniciar_Click()
    Dim lsIniAsistencia As String
    Dim lsCodTarjeta As String
    Dim lnDelay As Long
    Dim lsClave As String
    Dim lsNombre As String * 16
    Dim lsCodigo As String
    Dim lsCodPers As String
    Dim lsArea As String
    Dim lnVar As Integer
    'lnDelay = 15000000
    'lnDelay = 10000000
    lnDelay = 60000000
    
    While lnFinalizar = 0
        Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
        lsCodTarjeta = GetNumTarjeta
        If Len(lsCodTarjeta) = 16 Then
            If VerificaClave(lsCodTarjeta, lsCodigo, lsNombre, lsCodPers) Then
                lnVar = GrabaHoraRef(lsCodigo, lsCodPers, lsArea)
                If lnVar = 1 Then
                    'WriteToLcd lsNombre & Format(lblFecha, "hh:mm:ss AMPM")
                    lblmensaje.Caption = "INGRESO TURNO MAÑANA " + lsNombre + " " + Format(lblFecha, "hh:mm:ss AMPM")
                    lblmensaje.Refresh
                    Caption = lsNombre
                    'Demora lnDelay
                    'WriteToLcd "CONT.ASISTENCIA Pasar Tarjeta"
                    Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
                    'lblmensaje.Caption = "                      PASE  TARJETA               "
                    'lblmensaje.Refresh
                ElseIf lnVar = 2 Then
                    lblmensaje.Caption = "SALIDA TURNO MAÑANA " + lsNombre + " " + Format(lblFecha, "hh:mm:ss AMPM")
                    lblmensaje.Refresh
                    Caption = lsNombre
                    Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
                ElseIf lnVar = 3 Then
                    lblmensaje.Caption = "INGRESO TURNO TARDE " + lsNombre + " " + Format(lblFecha, "hh:mm:ss AMPM")
                    lblmensaje.Refresh
                    Caption = lsNombre
                    Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
                ElseIf lnVar = 4 Then
                    lblmensaje.Caption = "SALIDA TURNO TARDE " + lsNombre + " " + Format(lblFecha, "hh:mm:ss AMPM")
                    lblmensaje.Refresh
                    Caption = lsNombre
                    Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
                ElseIf lnVar = 0 Then
                    'WriteToLcd "Turno No Asignado a Empl."
                    lblmensaje.Caption = "NO REGISTRADO - FUERA DE RANGO DE HORARIO  "
                    lblmensaje.Refresh
                    'Demora lnDelay
                End If
            Else
                WriteToLcd "Tarjeta No Reconocida"
                Caption = "Tarjeta No Reconocida"
                Demora lnDelay
                'WriteToLcd "CONT.ASISTENCIA Pasar Tarjeta"
                Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
                lblmensaje.Caption = "TARJETA NO RECONOCIDA "
                lblmensaje.Refresh
            End If
        ElseIf Len(lsCodTarjeta) = 10 Then
            lsCodTarjeta = "112" & lsCodTarjeta
            If VerificaClaveFotoCheck(lsCodTarjeta, lsCodigo, lsNombre, lsCodPers) Then
                lnVar = GrabaHoraRef(lsCodigo, lsCodPers, lsArea)
                If lnVar = 0 Or lnVar = 1 Then
                    WriteToLcd lsNombre & Format(lblFecha, "hh:mm:ss AMPM")
                    Caption = lsNombre
                    Demora lnDelay
                    'WriteToLcd "CONT.ASISTENCIA Pasar Tarjeta"
                    Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
                Else
                    WriteToLcd "ERROR: Turno No Asignado a Empl."
                    Demora lnDelay
                End If
            Else
                WriteToLcd "Tarjeta No Reconocida"
                Caption = "Tarjeta No Reconocida"
                Demora lnDelay
                'WriteToLcd "CONT.ASISTENCIA Pasar Tarjeta"
                Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."
            End If

        End If
    Wend
End Sub

Private Sub cmdSalir_Click()
    FinLectora
    End
End Sub


Private Sub cmdVer_Click()
    'FinLectora
    frmVerAsistencia.Show 1
    'IniLectora
End Sub

Private Sub Form_Load()
    Dim sqlF As String
    Dim rsF As ADODB.Recordset
    Set oCon = New DConecta
    Dim lsCadena As String
    Me.Width = 8900
    Me.Height = 3690
    On Error GoTo Error
    
    Open App.Path & "\CtrAsist.Ini" For Input As #1
    Input #1, lsCadena
    Close #1
    gsPuerto = lsCadena
    
    oCon.AbreConexion
    
    gsPASS = "010127"
    
    'IniLectora
    IniciaPinPad Val(gsPuerto)
    lnFinalizar = 0
    
    lnMayor = 7800
    lnMenor = 5100
    
    lsTextoMayor = "&Avanzado >>>"
    lsTextoMenor = "&Regresar <<<"
    
    'MsgBox "1"
    ldFecha = oCon.GetFechaHoraServer()
    'MsgBox "2"
    mskFecha.Text = Format(ldFecha, "dd/mm/yyyy")
    Exit Sub
Error:
    MsgBox Err.Description, vbInformation, "Aviso"
    'cmdIniciar_Click
End Sub

Private Function VerificaClave(psCodTarj As String, psCodigo As String, psNombre As String, psCodPers As String) As Boolean
    Dim sqlE As String
    Dim lsCad As String
    Dim lnPos As Integer
    Dim lnFin As Integer
    
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    psCodTarj = Left(Right(Trim(psCodTarj), 7), 6)
    'sqlE = " Select E.cRHCod cEmpCod, PE.cPersNombre cNomPers, PE.cPersCod cCodPers From Persona PE" _
    '     & " Inner Join RHCuentas EC On PE.cPersCod = EC.cPersCod" _
    '     & " Inner join RRHH E on EC.cPersCod = E.cPersCod Where EC.cTarjCod = '" & psCodTarj & "'"
    
    sqlE = "Select E.cRHCod cEmpCod, PE.cPersNombre cNomPers, PE.cPersCod cCodPers From Persona PE " _
           & " Inner join RRHH E on PE.cPersCod = E.cPersCod " _
           & " Where E.cRHCod = '" & psCodTarj & "' "
    Set rsE = oCon.CargaRecordSet(sqlE)
    
    If rsE.EOF And rsE.BOF Then
        VerificaClave = False
    Else
        VerificaClave = True
        psCodigo = rsE!cEmpCod
        psCodPers = rsE!cCodPers
        
        lnPos = InStr(1, rsE!cNomPers, "/", vbTextCompare)
        lsCad = Left(rsE!cNomPers, lnPos)
        
        lnPos = InStr(1, rsE!cNomPers, ",", vbTextCompare)
        lnFin = InStr(lnPos, rsE!cNomPers, " ", vbTextCompare)
        
        If lnFin = 0 Then lnFin = 50
        
        lsCad = lsCad & Mid(rsE!cNomPers, lnPos + 1, lnFin - lnPos)
        psNombre = Trim(lsCad)
    End If
    
    rsE.Close
    Set rsE = Nothing
End Function

Private Function VerificaClaveFotoCheck(psCodTarj As String, psCodigo As String, psNombre As String, psCodPers As String) As Boolean
    Dim sqlE As String
    Dim lsCad As String
    Dim lnPos As Integer
    Dim lnFin As Integer
    
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    
    sqlE = " Select E.cRHCod cEmpCod, PE.cPersNombre cNomPers, PE.cPersCod cCodPers From Persona PE" _
         & " Inner Join RHCuentas EC On PE.cPersCod = EC.cPersCod" _
         & " Inner join RRHH E on EC.cPersCod = E.cPersCod Where E.cPersCod = '" & psCodTarj & "'"
    Set rsE = oCon.CargaRecordSet(sqlE)
    
    If rsE.EOF And rsE.BOF Then
        VerificaClaveFotoCheck = False
    Else
        VerificaClaveFotoCheck = True
        psCodigo = rsE!cEmpCod
        psCodPers = rsE!cCodPers
        
        lnPos = InStr(1, rsE!cNomPers, "/", vbTextCompare)
        lsCad = Left(rsE!cNomPers, lnPos)
        
        lnPos = InStr(1, rsE!cNomPers, ",", vbTextCompare)
        lnFin = InStr(lnPos, rsE!cNomPers, " ", vbTextCompare)
        
        If lnFin = 0 Then lnFin = 50
        
        lsCad = lsCad & Mid(rsE!cNomPers, lnPos + 1, lnFin - lnPos)
        psNombre = Trim(lsCad)
    End If
    
    rsE.Close
    Set rsE = Nothing
End Function

Private Function VerificaClaveCodPers(psCodigo As String, psNombre As String, psCodPers As String, psArea As String) As Boolean
'    Dim sqlE As String
'    Dim lsCad As String
'    Dim lnPos As Integer
'    Dim lnFin As Integer
'
'    Dim rsE As ADODB.Recordset
'    Set rsE = New ADODB.Recordset
'
'    'OpenDB
'    sqlE = " Select EC.cRHCod cEmpCod, PE.cPersNombre cNomPers, PE.cCodPers From Persona PE" _
'         & " Inner Join Empleado E on EC.cPersCod = E.cPersCod" _
'         & " Where PE.cPersCod = '" & psCodPers & "'"
'    rsE.Open sqlE, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If rsE.EOF And rsE.BOF Then
'        VerificaClave = False
'    Else
'        VerificaClave = True
'        psCodigo = rsE!cEmpCod
'        psCodPers = rsE!cCodPers
'        psArea = rsE!cAreCod
'
'        lnPos = InStr(1, rsE!cNomPers, "/", vbTextCompare)
'        lsCad = Left(rsE!cNomPers, lnPos)
'
'        lnPos = InStr(1, rsE!cNomPers, ",", vbTextCompare)
'        lnFin = InStr(lnPos, rsE!cNomPers, " ", vbTextCompare)
'
'        If lnFin = 0 Then lnFin = 50
'
'        lsCad = lsCad & Mid(rsE!cNomPers, lnPos + 1, lnFin - lnPos)
'
'        psNombre = Trim(lsCad)
'    End If
'
'    rsE.Close
'    Set rsE = Nothing
'    'CloseDB
End Function

'Private Function GetNumTarjeta() As String
'    Dim Result As Long
'    Dim lsTarjeta As String
'    Dim lsTarjetaAux1 As String
'    Dim lsTarjetaAux2 As String
'    Result = McrRead(lsTarjeta, 76, 0, lsTarjetaAux1, 10, 0, lsTarjetaAux2, 10, 0)
'    ' McrRead(szTrack1, Length1, wRetLength1, szTrack2, wLength2, wRetLength2, szTrack3, wLength3, wRetLength3)
'    If Result <= 0 Then
'        If IsNumeric(Left(lsTarjeta, 1)) Then
'            GetNumTarjeta = Mid$(lsTarjeta, 1, 10)
'        Else
'            GetNumTarjeta = Mid$(lsTarjeta, 2, 16)
'        End If
'    Else
'        MsgBox GetErrorPINPAD(Result) & " Consulte con Servicio Tecnico.", vbInformation, "Aviso"
'        GetNumTarjeta = ""
'    End If
'End Function

Private Function GrabaHoraRef(psEmpCod As String, psCodPers As String, psArea As String) As Integer
    Dim sqlE As String
    Dim lsHoy As String
    Dim lsNumMinutos As String
    Dim sqlT As String
    Dim rsT As ADODB.Recordset
    Dim lsTabla As String
    Set rsT = New ADODB.Recordset
    Dim sqlO As String
    lsTabla = "RHAsistenciaDet"
    'lsNumMinutos = 120
    lsNumMinutos = 90
    lsHoy = Format(CDate(lblFecha.Caption), "mm/dd/yyyy hh:mm:ss AMPM")
    
    oCon.AbreConexion
    sqlT = "SET DATEFIRST 1"
    oCon.Ejecutar sqlT
    sqlT = "Set Dateformat mdy"
    oCon.Ejecutar sqlT
    GrabaHoraRef = 0
        '1 INGRESO MAÑANA
        '2 SALIDA MAÑANA
        '3 INGRESO TARDE
        '4 SALIDA TARDE
    
    sqlT = " Select cRHHorarioTurno, dRHHorarioInicio, dRHHorarioFin , Tipo = Case " _
         & " When convert(varchar(10),dRHHorarioInicio,114) between convert(varchar(10),dateadd(minute,-" & lsNumMinutos & ",'" & lsHoy & "'),114) and convert(varchar(10),dateadd(minute," & lsNumMinutos & ",'" & lsHoy & "'),114) Then 'I'" _
         & " When convert(varchar(10),dRHHorarioFin,114) between convert(varchar(10),dateadd(minute,-" & lsNumMinutos & ",'" & lsHoy & "'),114) and convert(varchar(10),dateadd(minute," & lsNumMinutos & ",'" & lsHoy & "'),114) Then 'S'" _
         & " End  from RHHorarioDet RHD" _
         & "    Inner Join RHHorario RH" _
         & "            On RHD.cPersCod = RH.cPersCod And RHD.dRHHorarioFecha = RH.dRHHorarioFecha" _
         & "    Where   RHD.cPersCod = '" & psCodPers & "'" _
         & " And RHD.dRHHorarioFecha In" _
         & "    (   Select Max(RH1.dRHHorarioFecha) From RHHorario RH1" _
         & "            Where RH1.cPersCod = RH.cPersCod And RH1.dRHHorarioFecha = RH.dRHHorarioFecha And RH1.dRHHorarioFecha <=  '" & lsHoy & "')" _
         & " And (  Convert(varchar(10),dRHHorarioInicio,114) between Convert(varchar(10),dateadd(minute,-" & lsNumMinutos & ",'" & lsHoy & "'),114) And Convert(varchar(10),dateadd(minute," & lsNumMinutos & ",'" & lsHoy & "'),114)  or" _
         & "        Convert(varchar(10),dRHHorarioFin,114) between Convert(varchar(10),dateadd(minute,-" & lsNumMinutos & ",'" & lsHoy & "'),114) And Convert(varchar(10),dateadd(minute," & lsNumMinutos & ",'" & lsHoy & "'),114))" _
         & " And cRHHorarioDias = '" & Weekday(CDate(lblFecha.Caption), vbMonday) & "' Order by cRHHorarioTurno Desc"
    Set rsT = oCon.CargaRecordSet(sqlT)
    '********
     If Not RSVacio(rsT) Then
        'INGRESO
        If rsT!Tipo = "I" Then
            If Not ExisteTurno(Format(CDate(lblFecha.Caption), "dd/mm/yyyy hh:mm:ss AMPM"), rsT!cRHHorarioTurno, psCodPers, lsTabla) Then
                sqlE = " Insert " & lsTabla & " (cPersCod,dRHAsistenciaFechaRef,cRHTurno,dRHAsistenciaIngreso,dRHAsistenciaSalida,cUltimaActualizacion)" _
                     & " Values('" & psCodPers & "','" & lsHoy & "','" & rsT!cRHHorarioTurno & "','" & lsHoy & "',Null,'" & GetMovNro() & "')"
                    
            Else
                sqlE = " Update " & lsTabla _
                     & " Set dRHAsistenciaIngreso = '" & lsHoy & "'" _
                     & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = '" & rsT!cRHHorarioTurno & "'"
            End If
            
            If rsT!cRHHorarioTurno = "1" Then
                    GrabaHoraRef = 1
            ElseIf rsT!cRHHorarioTurno = "2" Then
                    GrabaHoraRef = 3
            End If
        'SALIDA
        Else
            If Not ExisteTurno(Format(CDate(lblFecha.Caption), "dd/mm/yyyy hh:mm:ss AMPM"), rsT!cRHHorarioTurno, psCodPers, lsTabla) Then
                sqlE = " Insert " & lsTabla & " (cPersCod,dRHAsistenciaFechaRef,cRHTurno,dRHAsistenciaIngreso,dRHAsistenciaSalida,cUltimaActualizacion)" _
                     & " Values('" & psCodPers & "','" & lsHoy & "','" & rsT!cRHHorarioTurno & "',Null,'" & lsHoy & "','" & GetMovNro() & "')"
                     
            Else
                sqlE = " Update " & lsTabla _
                     & " Set dRHAsistenciaSalida = '" & lsHoy & "'" _
                     & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = '" & rsT!cRHHorarioTurno & "'"
            End If
            
             If rsT!cRHHorarioTurno = "1" Then
                    GrabaHoraRef = 2
                 ElseIf rsT!cRHHorarioTurno = "2" Then
                    GrabaHoraRef = 4
             End If
            
            
        End If
        oCon.Ejecutar sqlE
        
        'INGRESOS DE LA TARDE ACTUALIZA LAS SALIDAS DEL TURNO 1
        'TURNO TARDE
        sqlO = ""
        If IIf(IsNull(rsT!Tipo), "X", rsT!Tipo) = "I" Then
               If rsT!cRHHorarioTurno = "2" Then
                    If ExisteTurnoEspecial(psCodPers, "T") = True And Weekday(lsHoy) < 7 Then
                        If Not ExisteTurno(Format(CDate(lblFecha.Caption), "dd/mm/yyyy hh:mm:ss AMPM"), "1", psCodPers, lsTabla) Then
                            sqlO = sqlO + " Insert " & lsTabla & " (cPersCod,dRHAsistenciaFechaRef,cRHTurno,dRHAsistenciaIngreso,dRHAsistenciaSalida,cUltimaActualizacion)" _
                            & " Values('" & psCodPers & "','" & Format(lsHoy, "mm/dd/yyyy 13:30:00.000") & "','" & "1" & "',Null,'" & Format(lsHoy, "mm/dd/yyyy 13:30:00.000") & "','" & GetMovNro() & "')"
                        Else
                            sqlO = sqlO + " Update " & lsTabla _
                            & " Set dRHAsistenciaSalida = '" & Format(lsHoy, "mm/dd/yyyy 13:30:00.000") & "'" _
                            & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = 1 "
                        End If
                    End If
              End If
        'SALIDAS DE LA MAÑANA ACTUALIZA LAS ENTRADAS DEL TURNO TARDE
        'TURNO MAÑANA
        ElseIf IIf(IsNull(rsT!Tipo), "X", rsT!Tipo) = "S" Then
              If rsT!cRHHorarioTurno = "1" Then
                    If ExisteTurnoEspecial(psCodPers, "M") = True And Weekday(lsHoy) < 7 Then
                       If Not ExisteTurno(Format(CDate(lblFecha.Caption), "dd/mm/yyyy hh:mm:ss AMPM"), "2", psCodPers, lsTabla) Then
                                '2004-01-01 15:45:00.000
                                sqlO = sqlO + " Insert " & lsTabla & " (cPersCod,dRHAsistenciaFechaRef,cRHTurno,dRHAsistenciaIngreso,dRHAsistenciaSalida,cUltimaActualizacion)" _
                                & " Values('" & psCodPers & "','" & Format(lsHoy, "mm/dd/yyyy 15:55:00.000") & "','" & "2" & "','" & Format(lsHoy, "mm/dd/yyyy 15:55:00.000") & "',Null,'" & GetMovNro() & "')"
                          Else
                                sqlO = sqlO + " Update " & lsTabla _
                                & " Set dRHAsistenciaIngreso = '" & Format(lsHoy, "mm/dd/yyyy 15:55:00.000") & "'" _
                                & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = 2 "
                        End If
                    End If
              End If
        End If
        
        If sqlO <> "" Then
            oCon.Ejecutar sqlO
            
        End If
        rsT.Close
        Set rsT = Nothing
    Else
        rsT.Close
        'SIN TURNO
        'TURNO MAÑANA
        sqlE = " "
        If DatePart("h", lsHoy) >= 7 And DatePart("h", lsHoy) < 12 Then
           If Not ExisteTurno(Format(CDate(lblFecha.Caption), "dd/mm/yyyy hh:mm:ss AMPM"), "1", psCodPers, lsTabla) Then
                    sqlE = sqlE + " Insert " & lsTabla & " (cPersCod,dRHAsistenciaFechaRef,cRHTurno,dRHAsistenciaIngreso,dRHAsistenciaSalida,cUltimaActualizacion)" _
                    & " Values('" & psCodPers & "','" & lsHoy & "','" & "1" & "','" & lsHoy & "',Null,'" & GetMovNro() & "')"
                    GrabaHoraRef = 1
            Else
                If DatePart("h", lsHoy) >= 10 And DatePart("h", lsHoy) <= 13 Then
                    sqlE = sqlE + " Update " & lsTabla _
                     & " set dRHAsistenciaSalida = '" & lsHoy & "'" _
                     & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = 1"
                     GrabaHoraRef = 2
                End If
                
            End If
        End If
        
        'TURNO TARDE
        If DatePart("h", lsHoy) >= 15 And DatePart("h", lsHoy) <= 19 And Weekday(lsHoy) < 7 Then
            
            If Not ExisteTurno(Format(CDate(lblFecha.Caption), "dd/mm/yyyy hh:mm:ss AMPM"), "2", psCodPers, lsTabla) Then
                  
                  If DatePart("h", lsHoy) >= 15 And DatePart("h", lsHoy) < 18 Then
                        sqlE = sqlE + " Insert " & lsTabla & " (cPersCod,dRHAsistenciaFechaRef,cRHTurno,dRHAsistenciaIngreso,dRHAsistenciaSalida,cUltimaActualizacion)" _
                        & " Values('" & psCodPers & "','" & lsHoy & "','" & "2" & "','" & lsHoy & "',Null,'" & GetMovNro() & "')"
                        GrabaHoraRef = 3
                  End If
                  
            Else
                  If DatePart("h", lsHoy) >= 17 And DatePart("h", lsHoy) <= 19 Then
                     sqlE = sqlE + " Update " & lsTabla _
                     & " set dRHAsistenciaSalida = '" & lsHoy & "'" _
                     & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = 2"
                     GrabaHoraRef = 4
                  End If
            End If
        
        End If
        If sqlE <> "" Then
            oCon.Ejecutar sqlE
        End If
        
        'HASTA LAS 12 DE LA NOCHE
        If DatePart("h", lsHoy) >= gnHoraFinal And Weekday(lsHoy) < 7 Then
            sqlE = " Update " & lsTabla _
            & " Set dRHAsistenciaSalida = '" & lsHoy & "'" _
            & " Where cPersCod = '" & psCodPers & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & lsHoy & "',101) And cRHTurno = 2"
            oCon.Ejecutar sqlE
            GrabaHoraRef = 4
        End If
        'rsT.Close
        'Set rsT = Nothing
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    FinLectora
    End
End Sub




Private Sub mskFecha_KeyPress(KeyAscii As Integer)
'    Dim sqlF As String
'    If KeyAscii = 13 Then
'        If IsDate(mskFecha.Text) Then
'            If MsgBox("Desea Actualizar la Fecha del Sistema ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'                mskFecha = Format(ldFecha, "dd/mm/yyyy")
'                Exit Sub
'            End If
'            sqlF = "Update FechaSistema Set dFecha = '" & Format(CDate(mskFecha), "dd/mm/yyyy") & "'"
'            dbCmact.Execute sqlF
'        End If
'    End If
End Sub

Private Sub Timer1_Timer()
    If IsDate(mskFecha.Text) Then
        lblFecha.Caption = Format(mskFecha.Text & " " & Format(oCon.GetFechaHoraServer, "hh:mm:ss"), "dd/mm/yyyy hh:mm:ss AMPM")
       ' WriteToLcd "CONT.ASISTENCIA " & Right(lblFecha.Caption, 11)
       
    End If
End Sub

Public Function Encripta(pnTexto As String, Valor As Boolean) As String
    Dim MiClase As cEncrypt
    Set MiClase = New cEncrypt
    Encripta = MiClase.ConvertirClave(pnTexto, , Valor)
End Function

Private Sub Demora(lnNumeroDemora As Long)
    Dim i As Long
    For i = 0 To lnNumeroDemora
    Next i
End Sub

Private Function ExisteTurno(pdFecha As Date, psTurno As String, psPersCod As String, psTabla) As Boolean
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    
    sqlE = "Select cPersCod from RHAsistenciaDet " & psTabla & "  Where cPersCod = '" & psPersCod & "' And convert(varchar(10),dRHAsistenciaFechaRef,101) = convert(varchar(10),'" & Format(pdFecha, "mm/dd/yyyy") & "',101) And cRHTurno = '" & psTurno & "'"
    Set rsE = oCon.CargaRecordSet(sqlE)
    
    If RSVacio(rsE) Then
        ExisteTurno = False
    Else
        ExisteTurno = True
    End If
    
    rsE.Close
    Set rsE = Nothing
End Function

Private Sub timerControlDia_Timer()
    Dim sqlP As String
    Dim lsFecha As String
    If Format(oCon.GetFechaHoraServer, "hh:mm:ss AM") > "06:00:00 AM" Then
        ldFecha = oCon.GetFechaHoraServer
        mskFecha.Text = Format(ldFecha, "dd/mm/yyyy")
        lsFecha = Format(ldFecha, "mm/dd/yyyy") & " 06:00:00 AM"
        
        'sqlP = " Insert tardanzassist (cEmpCod,dTarFec,cTurCod,cCodPers,dTarHorIng,dTarHorSal,nTarMin,nSalMin,bFalta)" _
             & " Select distinct EM.cEmpCod,'" & lsFecha & "' Fecha,TU.cTurCod,EM.cCodPers,null,null,0,0,1 from Turno TU" _
             & " Inner Join DiasSemana DS On TU.cJorCod = DS.cJorCod" _
             & " Inner Join EmpleadoTurno ET On ET.cTurCod = TU.cTurCod" _
             & " Inner Join Empleado EM On ET.cCodPers = EM.cCodPers" _
             & " Where EM.cEmpCod like 'L0000%' And cEmpEst <> '3' and cDiaSemCod = '" & Format(Weekday(ldFecha, vbMonday), "00") & "'" _
             & " And EM.cEmpCod+'" & lsFecha & "'+TU.cTurCod+EM.cCodPers not in " _
             & " (Select cEmpCod+'" & lsFecha & "'+cTurCod+cCodPers from tardanzassist" _
             & " Where dTarFec ='" & lsFecha & "')"
        'dbCmact.Execute sqlP
    End If
End Sub

Private Function GetMovNro()
    GetMovNro = Format(CDate(lblFecha.Caption), "yyyymmddhhmmss") & gsCodAge & "00" & gsCodUsu
End Function


Private Function ExisteTurnoEspecial(psPersCod As String, psTurno As String) As Boolean
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    
    sqlE = "SELECT  * FROM RHTURNOSPERSONAL where cPersCod ='" & psPersCod & "' and nRHTurno ='" & psTurno & "'"
    Set rsE = oCon.CargaRecordSet(sqlE)
    
    If RSVacio(rsE) Then
        ExisteTurnoEspecial = False
    Else
        ExisteTurnoEspecial = True
    End If
    
    rsE.Close
    Set rsE = Nothing
End Function



