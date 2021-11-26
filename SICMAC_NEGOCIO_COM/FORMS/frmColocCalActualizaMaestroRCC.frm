VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColocCalActualizaMaestroRCC 
   Caption         =   "Colocaciones - Calificacion Actualiza Maestros"
   ClientHeight    =   2445
   ClientLeft      =   5025
   ClientTop       =   1515
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Archivo a Cargar "
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdActualizarICS 
         Caption         =   "Actualizar ICS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2520
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtruta 
         Height          =   330
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   5085
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   330
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   405
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar RCC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
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
         Height          =   480
         Left            =   4860
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog cmdlOpen 
         Left            =   120
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ruta :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Actualizando :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblDato 
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   2865
      End
   End
End
Attribute VB_Name = "frmColocCalActualizaMaestroRCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* COLOCACIONES - CALIFICACION - ACTUALIZA MAESTRO RCC
'Archivo:  frmColocCalActualizaMaestroRCC.frm
'LAYG   :  01/11/2002.
'Resumen:  Actualiza el Maestro RCC, con informacion enviada por la SBS
Option Explicit
Dim fsServerConsol As String

Private Sub cmdOpen_Click()
   ' Establecer CancelError a True
    cmdlOpen.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    cmdlOpen.Flags = cdlOFNHideReadOnly
    cmdlOpen.InitDir = App.path
    ' Establecer los filtros
    cmdlOpen.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt|Archivos por lotes (*.bat)|*.bat|Archivos 112 (*.112)|*.112"
    ' Especificar el filtro predeterminado
    cmdlOpen.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
    cmdlOpen.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    txtruta = cmdlOpen.FileName
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

'*** Actualiza Maestro RCC
Private Sub cmdActualiza_Click()
If Len(Trim(Me.txtruta)) > 0 Then
    cmdActualiza.Enabled = False
    ActualizaMaestroRCC
    cmdActualiza.Enabled = True
Else
    MsgBox "Nombre de Tabla no válido", vbInformation, "Aviso"
End If
End Sub

Private Sub ActualizaMaestroRCC()
Dim Total As Long
Dim J As Long
Dim Nuevos As Long, Modif   As Long

Dim lsLinea As String, lsArchivo As String

Dim lsCodSbs  As String, lsDEUDOR As String, lsTPDOC_TRIB As String, lsCDDOC_TRIB  As String
Dim lsTPDOC_ID As String, lsCDDOC_ID As String, lsTipPers As String
Dim lnNROENTS As Integer
Dim lnCal0 As Single, lnCal1 As Single, lnCal2 As Single, lnCal3 As Single, lnCal4 As Single

Const lnPosCodSBS = 3
Const lnPosTPDOC_TRIB = 21
Const lnPosCDDOC_TRIB = 22
Const lnPosTPDOC_ID = 33
Const lnPosCDDOC_ID = 34
Const lnPosTipPers = 46
Const lnPosNroEnts = 48
Const lnPosCAL0 = 51
Const lnPosCAL1 = 56
Const lnPosCAL2 = 61
Const lnPosCAL3 = 66
Const lnPosCAL4 = 71
Const lnPosDEUDOR = 76

Dim loConec As DConecta
Dim lsSQL As String
'On Error GoTo ErrorActICd

If Len(Trim(txtruta)) = 0 Then
    MsgBox "Ruta no válida", vbInformation, "Aviso"
    Exit Sub
End If
lsArchivo = Trim(txtruta)
If Dir(lsArchivo) = "" Then Exit Sub

J = 0
Nuevos = 0
Modif = 0

Open lsArchivo For Input As #1   ' Abre el archivo.

Set loConec = New DConecta
loConec.AbreConexion

Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
    Input #1, lsLinea
    If Len(Trim(lsLinea)) > 0 Then ' Linea tiene datos
        If Mid(lsLinea, 1, 1) = "1" Then ' Datos de cliente
            ' Limpio las variables
            lsCodSbs = "": lsDEUDOR = "": lsTPDOC_TRIB = "": lsCDDOC_TRIB = ""
            lsTPDOC_ID = "": lsCDDOC_ID = "": lsTipPers = ""
            lnNROENTS = 0
            lnCal0 = 0: lnCal1 = 0: lnCal2 = 0: lnCal3 = 0: lnCal4 = 0
            
            lsCodSbs = Mid(lsLinea, lnPosCodSBS, 10)
            lsTPDOC_TRIB = Mid(lsLinea, lnPosTPDOC_TRIB, 1)
            lsCDDOC_TRIB = Mid(lsLinea, lnPosCDDOC_TRIB, 11)
            lsTPDOC_ID = Mid(lsLinea, lnPosTPDOC_ID, 1)
            lsCDDOC_ID = Mid(lsLinea, lnPosCDDOC_ID, 12)
            lsTipPers = Mid(lsLinea, lnPosTipPers, 1)
            lnNROENTS = Val(Mid(lsLinea, lnPosNroEnts, 3))
            lnCal0 = Val(Mid(lsLinea, lnPosCAL0, 3) + Val("0." & Mid(lsLinea, lnPosCAL0 + 3, 2)))
            lnCal1 = Val(Mid(lsLinea, lnPosCAL1, 3) + Val("0." & Mid(lsLinea, lnPosCAL1 + 3, 2)))
            lnCal2 = Val(Mid(lsLinea, lnPosCAL2, 3) + Val("0." & Mid(lsLinea, lnPosCAL2 + 3, 2)))
            lnCal3 = Val(Mid(lsLinea, lnPosCAL3, 3) + Val("0." & Mid(lsLinea, lnPosCAL3 + 3, 2)))
            lnCal4 = Val(Mid(lsLinea, lnPosCAL4, 3) + Val("0." & Mid(lsLinea, lnPosCAL4 + 3, 2)))
            lsDEUDOR = Mid(lsLinea, lnPosDEUDOR, 80)
            
            Nuevos = Nuevos + 1
            'FEREP  CDDEU  DSDEU  TPDOC_TRIB  CDDOC_TRIB  TPDOC_ID CDDOC_ID  NROENTS  TPEMP CALCRED_SIS TPREG CDENT DSENT
            lsSQL = "INSERT INTO " & fsServerConsol & "RCC ( cCodSBS, cTiDoTr, cNuDoTr, cTiDoCi, cNuDoCi, " _
               & " cTipPers, nEmpRep, nCal0, nCal1, nCal2, nCal3, nCal4, cDeudor ) " _
               & " VALUES ('" & lsCodSbs & "','" & lsTPDOC_TRIB & "','" & lsCDDOC_TRIB & "','" _
               & lsTPDOC_ID & "','" & lsCDDOC_ID & "','" & lsTipPers & "'," _
               & lnNROENTS & "," & lnCal0 & "," & lnCal1 & "," & lnCal2 & "," _
               & lnCal3 & "," & lnCal4 & ",'" _
               & ReemplazaApostrofe(lsDEUDOR) & "' )"
            
            loConec.Ejecutar lsSQL
            'Barra.Value = Int(j / Total * 100)
            Me.lblDato.Caption = Trim(lsCodSbs) & "-" & Mid(lsDEUDOR, 1, 13) & "  Nuevos " & Nuevos & " - Modificados " & Modif
        
        End If
    End If
    DoEvents
Loop
Close #1   ' Cierra el archivo.
Set loConec = Nothing

MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
Exit Sub

ErrorActICd:
    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"

End Sub

Private Sub cmdActualizarICS_Click()
If Len(Trim(Me.txtruta)) > 0 Then
    cmdActualizarICS.Enabled = False
    ActualizaMaestroICS
    cmdActualizarICS.Enabled = True
Else
    MsgBox "Nombre de Tabla no válido", vbInformation, "Aviso"
End If
End Sub

Private Sub ActualizaMaestroICS()
Dim Total As Long
Dim J As Long
Dim Nuevos As Long, Modif   As Long

Dim lsLinea As String, lsArchivo As String, lsCodSbs As String

Dim lsTipDoc_ID As String, lsNroDoc_ID As String
Dim lsApePat As String, lsApeMat As String, lsNombre As String
Dim lsTipPers As String
Dim lnCal0 As Single, lnCal1 As Single, lnCal2 As Single, lnCal3 As Single, lnCal4 As Single
Dim lnNROENTS As Integer

'Const lnPosCodSBS = 3
Const lnPosTipDoc_ID = 1
Const lnPosNroDoc_ID = 2
Const lnPosApePat = 14
Const lnPosApeMat = 74
Const lnPosNombre = 94
Const lnPosTipPers = 124
Const lnPosCAL0 = 644
Const lnPosCAL1 = 648
Const lnPosCAL2 = 652
Const lnPosCAL3 = 656
Const lnPosCAL4 = 660
Const lnPosNroEnts = 668

Dim loConec As DConecta
Dim lsSQL As String
On Error GoTo ErrorActICS

If Len(Trim(txtruta)) = 0 Then
    MsgBox "Ruta no válida", vbInformation, "Aviso"
    Exit Sub
End If
lsArchivo = Trim(txtruta)
If Dir(lsArchivo) = "" Then Exit Sub

J = 0
Nuevos = 0
Modif = 0

Open lsArchivo For Input As #1   ' Abre el archivo.

Set loConec = New DConecta
loConec.AbreConexion

Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
    Input #1, lsLinea
    If Len(Trim(lsLinea)) > 0 Then ' Linea tiene datos
        'If Mid(lsLinea, 1, 1) = "1" Then ' Datos de cliente
            ' Limpio las variables
            lsTipDoc_ID = "": lsNroDoc_ID = ""
            lsApePat = "": lsApeMat = "": lsNombre = ""
            lsTipPers = ""
            lnCal0 = 0: lnCal1 = 0: lnCal2 = 0: lnCal3 = 0: lnCal4 = 0
            lnNROENTS = 0
            
            lsTipDoc_ID = Mid(lsLinea, lnPosTipDoc_ID, 1)
            lsNroDoc_ID = Mid(lsLinea, lnPosNroDoc_ID, 12)
            lsApePat = Mid(lsLinea, lnPosApePat, 60)
            lsApeMat = Mid(lsLinea, lnPosApeMat, 20)
            lsNombre = Mid(lsLinea, lnPosNombre, 30)
            lsTipPers = Mid(lsLinea, lnPosTipPers, 1)
            lnCal0 = Val(Mid(lsLinea, lnPosCAL0, 3))
            lnCal1 = Val(Mid(lsLinea, lnPosCAL1, 3))
            lnCal2 = Val(Mid(lsLinea, lnPosCAL2, 3))
            lnCal3 = Val(Mid(lsLinea, lnPosCAL3, 3))
            lnCal4 = Val(Mid(lsLinea, lnPosCAL4, 3))
            lnNROENTS = Val(Mid(lsLinea, lnPosNroEnts, 3))

            Nuevos = Nuevos + 1
            
            lsSQL = "INSERT INTO " & fsServerConsol & "ICS (cTIDOID, cNUDOID, cTIPPERS, " _
              & "nCal0,nCal1,nCal2,nCal3,nCal4,NROENTS,cDEUDOR ) " _
              & " VALUES ('" & lsTipDoc_ID & "','" & lsNroDoc_ID & "','" & lsTipPers & "'," _
              & lnCal0 & "," & lnCal1 & "," & lnCal2 & "," & lnCal3 & "," & lnCal4 & "," _
              & lnNROENTS & ",'" _
              & ReemplazaApostrofe(Mid(Trim(lsApePat) & " " & Trim(lsApeMat) & " " & Trim(lsNombre), 1, 79)) & "' )"
            
            loConec.Ejecutar lsSQL
            Me.lblDato.Caption = "Actualizados " & Nuevos & " - " & lsNroDoc_ID & Trim(lsCodSbs) & "-" & ReemplazaApostrofe(Trim(lsApePat) & " " & Trim(lsApeMat))
        
        'End If
    End If
    DoEvents
Loop
Close #1   ' Cierra el archivo.
Set loConec = Nothing


MsgBox "Actualización Finalizada con Exito", vbInformation, "Aviso"
Exit Sub

ErrorActICS:
    MsgBox "Error Nº[" & Err.Number & " ] " & Err.Description, vbInformation, "Aviso"

End Sub


Private Function VerificaMaestroICS(ByVal psTiDoID As String, ByVal psNuDoId As String) As Boolean

End Function

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Dim loConstS As NConstSistemas
    Set loConstS = New NConstSistemas
        fsServerConsol = loConstS.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstS = Nothing

End Sub
