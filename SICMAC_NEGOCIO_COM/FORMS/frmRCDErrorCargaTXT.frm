VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRCDErrorCargaTXT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe RCD - Carga Archivo de Error TXT"
   ClientHeight    =   4920
   ClientLeft      =   795
   ClientTop       =   3435
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   10635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9060
      TabIndex        =   9
      Top             =   4500
      Width           =   1320
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdPasar 
      Caption         =   "Pasar a Tabla"
      Height          =   360
      Left            =   7080
      TabIndex        =   5
      Top             =   4500
      Width           =   1770
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "&Transformar Archivo"
      Height          =   360
      Left            =   5280
      TabIndex        =   4
      Top             =   4500
      Width           =   1725
   End
   Begin MSComDlg.CommonDialog cmdlOpen 
      Left            =   6600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   330
      Left            =   5700
      TabIndex        =   3
      Top             =   60
      Width           =   405
   End
   Begin VB.TextBox txtRuta 
      Height          =   330
      Left            =   780
      TabIndex        =   1
      Top             =   60
      Width           =   4905
   End
   Begin MSComctlLib.ListView lstCampos 
      Height          =   3975
      Left            =   165
      TabIndex        =   0
      Top             =   480
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Correlativo"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CodUnico"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre Cmact "
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NumDoc"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CiuuCMACT"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CodRRPP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CodSBS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "CodUnicoSBS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "NomSBS"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DocSBS"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "CIIUSBS"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "CodRRPPSBS"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSMask.MaskEdBox txtFechaArchivoTXT 
      Height          =   315
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Archivo TXT"
      Height          =   255
      Left            =   7380
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ruta :"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmRCDErrorCargaTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
' RCD - Carga Errores en TXT
'LAYG   :  10/01/2003.
'Resumen:  Nos permite cargar los errores enviados por la SBS

Option Explicit

Dim lsArchivo As String
Dim Con As ADODB.Connection

Dim lsCorrelativo As String
Dim lsNombreEmp As String
Dim lsCiiu As String
Dim lsSigla As String
Dim lsNumDoc As String
Dim lsCodUnico As String
Dim lsCodRRPP As String

Dim lsCodSbs As String
Dim lsNomSBS As String
Dim lsCiiuSBS As String
Dim lsNumDocSBS As String
Dim lsCodUnicSBS As String
Dim lsCodRRPPSBS As String

Dim fsServConsol As String

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
    txtRuta = cmdlOpen.FileName
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

Private Sub cmdPasar_Click()
Dim I As Integer
Dim lsSQL As String
Dim loBase As DConecta
Dim lnTotal As Long

lsSQL = "DELETE " & fsServConsol & "RCDError "

Set loBase = New DConecta
    loBase.AbreConexion
    
    loBase.Ejecutar (lsSQL)


    If Me.lstCampos.ListItems.Count > 0 Then
        For I = 1 To lstCampos.ListItems.Count
            lsCorrelativo = Trim(lstCampos.ListItems(I))
            lsCodUnico = Trim(lstCampos.ListItems(I).SubItems(1))
            lsNombreEmp = Trim(Replace(lstCampos.ListItems(I).SubItems(2), "'", "''"))
            lsNumDoc = Trim(lstCampos.ListItems(I).SubItems(3))
            lsCiiu = Trim(lstCampos.ListItems(I).SubItems(4))
            lsCodRRPP = Trim(lstCampos.ListItems(I).SubItems(5))
            lsCodSbs = Trim(lstCampos.ListItems(I).SubItems(6))
            lsCodUnicSBS = Trim(lstCampos.ListItems(I).SubItems(7))
            lsNomSBS = Trim(Replace(lstCampos.ListItems(I).SubItems(8), "'", "''"))
            lsNumDocSBS = Trim(lstCampos.ListItems(I).SubItems(9))
            lsCiiuSBS = Trim(lstCampos.ListItems(I).SubItems(10))
            lsCodRRPPSBS = Trim(lstCampos.ListItems(I).SubItems(11))
            
            If lsCorrelativo <> "" Then
            
                lsSQL = "Insert into " & fsServConsol & "RCDError " _
                    & " (cPersCod, cCodSBS, cCorr, cCodUnicoEmp, cNombreEmp, cNumDocEmp, cCiiuEmp, cCodRRPP, " _
                    & "  cCodUnicSBS, cNomSBS, cNumDocSBS, cCiiuSBS, cCodRRPPSBS) " _
                    & " VALUES('" & lsCodUnico & "','" & lsCodSbs & "','" & lsCorrelativo & "','" & lsCodUnico & "','" & lsNombreEmp & "','" _
                    & lsNumDoc & "','" & lsCiiu & "','" & lsCodRRPP & "','" & lsCodUnicSBS & "','" _
                    & lsNomSBS & "','" & lsNumDocSBS & "','" & lsCiiuSBS & "','" & lsCodRRPPSBS & "')"
            
                loBase.Ejecutar (lsSQL)
            End If
            
            barra.value = Int(I / lstCampos.ListItems.Count * 100)
        
        Next
         MsgBox "Actualiza Terminada con exito", vbInformation, "Aviso"
    End If

Set loBase = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTrans_Click()
Dim lsLinea As String
Dim lsAux As String
Dim lnPos As Long
Dim nItem As ListItem
Dim lnTotal As Long

Const lnPosCorr = 5
Const lnPosSBS = 12
Const lnPosDoc = 39
Const lnPosCodRRPP = 52
Const lnPosCIIU = 68

Const lnPosNombre = 5
'Const lnPosCodUnico = 19

Dim lbFlagSBS As Boolean
Dim rsPersRCD As ADODB.Recordset
Dim loBase As DConecta
Dim lsPersRCD As String
Dim ldFechaArchivoTXT As Date

If Not IsDate(Me.txtFechaArchivoTXT.Text) Then
    MsgBox "Fecha de Archivo no valida", vbInformation, "Aviso"
    Exit Sub
End If

ldFechaArchivoTXT = Format(Me.txtFechaArchivoTXT.Text, "dd/mm/yyyy")
If Len(Trim(txtRuta)) = 0 Then
    MsgBox "Ruta no válida", vbInformation, "Aviso"
    Exit Sub
End If
lsArchivo = Trim(txtRuta)
If Dir(lsArchivo) = "" Then Exit Sub

lstCampos.ListItems.Clear
lsCorrelativo = "": lsNombreEmp = "": lsNumDoc = "": lsCiiu = "": lsCodSbs = ""
lsNomSBS = "": lsNumDocSBS = "": lsCiiuSBS = ""

Open lsArchivo For Input As #1   ' Abre el archivo.

Set loBase = New DConecta
    loBase.AbreConexion
    Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
        Input #1, lsLinea
        If Len(Trim(lsLinea)) > 0 Then ' Linea tiene datos
            If Mid(lsLinea, 1, 4) = "RCD:" Then
                'Correlativo
                lsCorrelativo = Mid(lsLinea, lnPosCorr, 6)
                
                lsCodSbs = Mid(lsLinea, lnPosSBS, 10)
                '*******************
                lsPersRCD = "SELECT cPersCod From " & fsServConsol & "RCDvc" & Format(ldFechaArchivoTXT, "yyyymm") & "01" & _
                            " WHERE  cNumSec ='" & lsCorrelativo & "'"
                Set rsPersRCD = loBase.CargaRecordSet(lsPersRCD)
                
                If rsPersRCD.BOF And rsPersRCD.EOF Then
                    MsgBox "No encuentro codigo SBS en archivo"
                Else
                    lsCodUnico = rsPersRCD!cPersCod
                    lsCodUnicSBS = rsPersRCD!cPersCod
                End If
                rsPersRCD.Close
                '*******************
                lsNumDoc = Mid(lsLinea, lnPosDoc, 8)
                lsCodRRPP = Mid(lsLinea, lnPosCodRRPP, 15)
                lsCiiu = Mid(lsLinea, lnPosCIIU, 4)
            ElseIf Mid(lsLinea, 1, 4) <> "SBS:" And lsCorrelativo <> "" And lbFlagSBS = False Then
                lsNombreEmp = Replace(Mid(lsLinea, 1, Len(lsLinea)), "#", "Ñ")
            ElseIf Mid(lsLinea, 1, 4) = "SBS:" Then
                lbFlagSBS = True
                If Mid(lsLinea, 5, 1) = " " Then ' No tiene nombre SBS
                    lsNumDocSBS = Mid(lsLinea, lnPosDoc, 8)
                    lsCodRRPPSBS = Mid(lsLinea, lnPosCodRRPP, 15)
                    lsCiiuSBS = Mid(lsLinea, lnPosCIIU, 4)
                Else  ' Tiene nombre de SBS
                    lsNomSBS = Replace(Mid(lsLinea, lnPosNombre, Len(lsLinea)), "#", "Ñ")
                End If
            ElseIf Mid(lsLinea, 1, 4) <> "" And lbFlagSBS = True Then
                lsNomSBS = Replace(Mid(lsLinea, 1, Len(lsLinea)), "#", "Ñ")
            End If
        Else ' no tiene datos cambio de persona
            ' Pasa las variables al listbox
            Set nItem = Me.lstCampos.ListItems.Add(, , lsCorrelativo)
            nItem.SubItems(1) = lsCodUnico
            nItem.SubItems(2) = lsNombreEmp
            nItem.SubItems(3) = lsNumDoc
            nItem.SubItems(4) = lsCiiu
            nItem.SubItems(5) = lsCodRRPP
            nItem.SubItems(6) = lsCodSbs
            nItem.SubItems(7) = lsCodUnicSBS
            nItem.SubItems(8) = lsNomSBS
            nItem.SubItems(9) = lsNumDocSBS
            nItem.SubItems(10) = lsCiiuSBS
            nItem.SubItems(11) = lsCodRRPPSBS
            ' Limpio las variables
            lsCorrelativo = "": lsNombreEmp = "": lsNumDoc = "": lsCiiu = "": lsCodSbs = ""
            lsNomSBS = "": lsNumDocSBS = "": lsCiiuSBS = ""
            lsCodUnico = "": lsCodRRPP = "": lsCodUnicSBS = "": lsCodRRPPSBS = ""
            lbFlagSBS = False
        End If
        ''Tamano = LOF(1)  'Tamaño de la cadena
        DoEvents
    Loop

Set loBase = Nothing
Close #1   ' Cierra el archivo.
Set nItem = Me.lstCampos.ListItems.Add(, , lsCorrelativo)
nItem.SubItems(1) = lsCodUnico
nItem.SubItems(2) = lsNombreEmp
nItem.SubItems(3) = lsNumDoc
nItem.SubItems(4) = lsCiiu
nItem.SubItems(5) = lsCodRRPP
nItem.SubItems(6) = lsCodSbs
nItem.SubItems(7) = lsCodUnicSBS
nItem.SubItems(8) = lsNomSBS
nItem.SubItems(9) = lsNumDocSBS
nItem.SubItems(10) = lsCiiuSBS
nItem.SubItems(11) = lsCodRRPPSBS
End Sub


Private Sub Form_Load()
Dim loConstSistema As NConstSistemas
    
    Set loConstSistema = New NConstSistemas
        fsServConsol = loConstSistema.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSistema = Nothing
txtRuta.Text = App.path
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
