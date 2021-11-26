VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAnxEncajeBCR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos de Encaje BCR"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frmanxencajebcr.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraFecha 
      Height          =   630
      Left            =   615
      TabIndex        =   8
      Top             =   0
      Width           =   2715
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1320
         TabIndex        =   9
         Top             =   195
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   195
         TabIndex        =   10
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdAnx04ME 
      Caption         =   "Anx 4 ME"
      Height          =   450
      Left            =   2430
      TabIndex        =   7
      Top             =   2490
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx04MN 
      Caption         =   "Anx 4 MN"
      Height          =   450
      Left            =   435
      TabIndex        =   6
      Top             =   2490
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx03ME 
      Caption         =   "Anx 3 ME"
      Height          =   450
      Left            =   2415
      TabIndex        =   5
      Top             =   1920
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx03MN 
      Caption         =   "Anx 3 MN"
      Height          =   450
      Left            =   420
      TabIndex        =   4
      Top             =   1920
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx02ME 
      Caption         =   "Anx 2 ME"
      Height          =   450
      Left            =   2400
      TabIndex        =   3
      Top             =   1365
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx02MN 
      Caption         =   "Anx 2 MN"
      Height          =   450
      Left            =   405
      TabIndex        =   2
      Top             =   1365
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx01ME 
      Caption         =   "Anx 1 ME"
      Height          =   450
      Left            =   2400
      TabIndex        =   1
      Top             =   810
      Width           =   1250
   End
   Begin VB.CommandButton CmdAnx01MN 
      Caption         =   "Anx 1 MN"
      Height          =   450
      Left            =   405
      TabIndex        =   0
      Top             =   810
      Width           =   1250
   End
End
Attribute VB_Name = "frmAnxEncajeBCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim R As New ADODB.Recordset
Dim oCon As DConecta

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Public Sub GeneraAnx01MN(psFecha As String)
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "01MN.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX1") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo
'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sNumero
Print #ArcSal, "00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))

'INICIO
'*********************************************
Dim nSaltar As Integer
nSaltar = 0
For j = 1 To 35
nSaltar = 0
    Select Case j
                
            '*** PEAC 20100715
            Case 1:     sCol = "B": nSaltar = 1
                                    sCodOpe = "001000" 'OBLIGACIONES INMEDIATAS
            Case 4:     sCol = "E": nSaltar = 1
                                    sCodOpe = "004000" 'A PLAZO MAYOR DE 30 DIAS
            Case 6:     sCol = "G": nSaltar = 1
                                    sCodOpe = "009000" '
            Case 7:     sCol = "H": nSaltar = 1
                                    sCodOpe = "040000" '
            Case 8:     sCol = "I": nSaltar = 1
                                    sCodOpe = "050000" '
            Case 15:    sCol = "P": nSaltar = 1
                                    sCodOpe = "071000" '
            Case 30:    sCol = "AE": nSaltar = 1
                                    sCodOpe = "070000" '
            Case 31:    sCol = "AF": nSaltar = 1
                                    sCodOpe = "080000" '
            Case 32:    sCol = "AG": nSaltar = 1
                                    sCodOpe = "090000" '
            Case 33:    sCol = "AH": nSaltar = 1
                                    sCodOpe = "100000" '
            Case 35:    sCol = "AJ": nSaltar = 1
                                    sCodOpe = "085000" '
        
'            Case 16:    sCol = "Q": nSaltar = 1
'                                    sCodOpe = "071100" '
'            Case 27:    sCol = "AE": nSaltar = 1
'                                    sCodOpe = "070000" '
'            Case 28:    sCol = "AF": nSaltar = 1
'                                    sCodOpe = "080000" '
'            Case 29:    sCol = "AG": nSaltar = 1
'                                    sCodOpe = "090000" '
'            Case 30:    sCol = "AH": nSaltar = 1
'                                    sCodOpe = "100000" '
'            Case 32:    sCol = "AJ": nSaltar = 1
'                                    sCodOpe = "085000" '
        
        End Select
    If nSaltar = 1 Then
    i = 17
    For i = 17 To 17 + nDias - 1
        
        sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlhoja.Range("A" & i), "00") & "00"
        sNumero = Format(xlhoja.Range(sCol & i), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        
        If sNumero = "" Or sNumero = "00000000000000" Then

        Else
            sNumero = Format(sNumero, "00000000000000")
            Print #1, sCad; sNumero
        End If
            
    Next i
'        If J = 15 Then
'            I = 53
'            sCad = "071100" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "3100"
'            sNumero = Format(xlHoja.Range(sCol & I), "000000000000.00")
'            sNumero = Replace(sNumero, ".", "")
'            If sNumero = "" Or sNumero = "00000000000000" Then
'            Else
'                sNumero = Format(sNumero, "00000000000000")
'                Print #1, sCad; sNumero
'            End If
'        End If
    End If
Next j






'*********************************************
'FIN
'For i = 16 To 16 + nDias - 1
'    For J = 1 To 21
'        Select Case J
'            Case 1:     sCol = "B": sCodOpe = "001000"
'            Case 2:     sCol = "C": sCodOpe = "002000"
'            Case 3:     sCol = "D": sCodOpe = "003000"
'            Case 4:     sCol = "E": sCodOpe = "004000"
'            Case 5:     sCol = "F": sCodOpe = "005000"
'            Case 6:     sCol = "G": sCodOpe = "006000" '
'            Case 7:     sCol = "H": sCodOpe = "009000" '
'            Case 8:     sCol = "I": sCodOpe = "008000" '
'            Case 9:     sCol = "J": sCodOpe = "010000" '
'            Case 10:     sCol = "K": sCodOpe = "020000" '
'            Case 11:     sCol = "L": sCodOpe = "030000" '
'            Case 12:     sCol = "M": sCodOpe = "040000" '
'            Case 13:     sCol = "N": sCodOpe = "050000" '
'            Case 14:     sCol = "O": sCodOpe = "060000" '
'            Case 15:     sCol = "P": sCodOpe = "065000"
'            Case 16:     sCol = "Q": sCodOpe = "070000"
'            Case 17:     sCol = "R": sCodOpe = "080000"
'
'            Case 18:     sCol = "S": sCodOpe = "090000" '
'            Case 19:     sCol = "T": sCodOpe = "100000" '
'            Case 20:     sCol = "U": sCodOpe = "200000" '
'
'            Case 21:     sCol = "V": sCodOpe = "085000" '
'
'        End Select
'        sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlHoja.Range("A" & i), "00") & "00"
'        sNumero = Format(xlHoja.Range(sCol & i), "000000000000.00")
'        sNumero = Replace(sNumero, ".", "")
'
'        If sNumero = "" Or sNumero = "00000000000000" Then
'
'        Else
'            sNumero = Format(sNumero, "00000000000000")
'            Print #1, sCad; sNumero
'        End If
'
'    Next J
'Next i

Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub

ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub GeneraAnx02MN(psFecha As String)
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

'On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar

psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "02MN.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("Anx2") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo


'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sNumero
Print #ArcSal, "00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))


'INICIO
'*******************
Dim sCodOpe2 As String
sCodOpe2 = ""

For j = 1 To 21

      Select Case j
            '*** PEAC 20100614
            Case 1:     sCol = "B": sCodOpe = "100000": sCodOpe2 = "COFDPEPLXXX"
'            Case 2:    sCol = "C": sCodOpe = "100000": sCodOpe2 = "BSUDPEPLXXX"
'            Case 3:    sCol = "D": sCodOpe = "100000": sCodOpe2 = "BANCPEPLXXX"
            Case 2:    sCol = "C": sCodOpe = "101000": sCodOpe2 = ""
            Case 3:     sCol = "D": sCodOpe = "300000": sCodOpe2 = "0801PE73XXX"
            Case 4:     sCol = "E": sCodOpe = "300000": sCodOpe2 = "0802PE44XXX"
            Case 5:     sCol = "F": sCodOpe = "300000": sCodOpe2 = "0803PE54XXX"
            Case 6:     sCol = "G": sCodOpe = "300000": sCodOpe2 = "0805PE73XXX"
            Case 7:     sCol = "H": sCodOpe = "300000": sCodOpe2 = "0807PE43XXX"
            Case 8:     sCol = "I": sCodOpe = "300000": sCodOpe2 = "0806PE84XXX"
            Case 9:     sCol = "J": sCodOpe = "300000": sCodOpe2 = "0808PE64XXX"
            Case 10:     sCol = "K": sCodOpe = "300000": sCodOpe2 = "0809PE56XXX"
            Case 11:     sCol = "L": sCodOpe = "300000": sCodOpe2 = "0810PE73XXX"
            Case 12:     sCol = "M": sCodOpe = "300000": sCodOpe2 = "0812PE56XXX"
            Case 13:     sCol = "N": sCodOpe = "300000": sCodOpe2 = "0813PE52XXX"
            Case 14:     sCol = "O": sCodOpe = "300000": sCodOpe2 = "0842PE74XXX"
            Case 15:     sCol = "P": sCodOpe = "300000": sCodOpe2 = "0839PEPLXXX"
            Case 16:     sCol = "Q": sCodOpe = "300000": sCodOpe2 = "0844PE76XXX"
            Case 17:     sCol = "R": sCodOpe = "300000": sCodOpe2 = "0836PE43XXX"
            Case 18:     sCol = "S": sCodOpe = "300000": sCodOpe2 = "BSUDPEPLXXX"
            Case 19:     sCol = "T": sCodOpe = "300000": sCodOpe2 = "0092PE64BCR"
            Case 20:     sCol = "U": sCodOpe = "301000": sCodOpe2 = ""
        End Select

    'I = 19
    i = 20 '*** PEAC 20100715
        If sCol <> "" Then
            'For I = 19 To 19 + nDias - 1 '*** PEAC 20100715
            For i = 20 To 20 + nDias - 1
                sCad = ""
                sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00")
                
                'sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(I - 19), "00")
                sCad = sCad & sCodOpe2
                
                'sCad = sCad & CStr(xlHoja.Range(sCol & "15")) & Replace(Space(11 - Len(CStr(xlHoja.Range(sCol & "15")))), " ", "0")
               
                'ALPA 20100114******************
                'sCad = sCad & "00"
                'If sCol = "F" Or sCol = "W" Then
                If sCol = "C" Or sCol = "U" Then
                    sCad = sCad & Space(11)
                End If
                sCad = sCad & "00"
                sNumero = Format(CDbl(xlhoja.Range(sCol & i)), "000000000000.00")
                '*************************************************************
                sNumero = Replace(sNumero, ".", "")
                If CDbl(sNumero) = 0 Then
                    sNumero = ""
                Else
                    If sNumero = "00000000000000" Then
                        sNumero = ""
                    Else
                        sNumero = Format(sNumero, "00000000000000")
                    End If
                End If
                If sNumero <> "" Then
                   Print #1, sCad; sNumero
                End If
            Next i
        End If
    sCodOpe2 = ""
    sCol = ""
Next j

Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
End Sub
Public Sub GeneraAnx03MN(psFecha As String)
'''Dim psArchivoALeer As String
'''Dim psArchivoAGrabar As String
'''Dim xlAplicacion As Excel.Application
'''Dim xlLibro As Excel.Workbook
'''Dim xlHoja As Excel.Worksheet
'''Dim bExiste As Boolean
'''Dim bEncontrado As Boolean
'''Dim fs As New Scripting.FileSystemObject
'''
'''Dim I As Integer
'''Dim J As Integer
'''Dim sCad As String
'''
'''Dim Fecha As Date
'''
'''On Error GoTo ErrBegin
'''
''''Verifica el Archivo de Excel que se va a cargar
'''psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "03MN.xls"
'''bExiste = fs.FileExists(psArchivoALeer)
'''
'''If bExiste = False Then
'''    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
'''    Exit Sub
'''End If
'''    psArchivoAGrabar = App.path & "\SPOOLER\00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00.txt"
'''    Set xlAplicacion = New Excel.Application
'''    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
'''    bEncontrado = False
'''    For Each xlHoja In xlLibro.Worksheets
'''        If UCase(xlHoja.Name) = UCase("Anx3") Then
'''            bEncontrado = True
'''            xlHoja.Activate
'''            Exit For
'''        End If
'''    Next
'''    If bEncontrado = False Then
'''        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
'''        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
'''        Exit Sub
'''    End If
''''Creacion del Archivo
'''Open psArchivoAGrabar For Output As #1
'''Dim nDias As Integer
'''Dim sCol As String
'''Dim sCodOpe As String
'''Dim sNumero
'''Print #1, "00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
'''sCad = ""
'''nDias = CInt(Mid(psFecha, 1, 2))
'''
'''Dim sCodOpe2 As String
'''sCodOpe2 = ""
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "03MN.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX3") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo
'If Not psArchivoAGrabar Then
'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
'End If
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sCodOpe2 As String
sCodOpe2 = ""
Dim sNumero
Print #ArcSal, "00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
'INICIO
'***********************************

For j = 1 To 2
        Select Case j
            'Case 1:     sCol = "B": sCodOpe = "100000"
            'Case 2:     sCol = "C": sCodOpe = "101000"
            'Case 3:     sCol = "D": sCodOpe = "200000"
            'Case 4:     sCol = "E": sCodOpe = "201000"
            'Case 1:     sCol = "F": sCodOpe = "300000": sCodOpe2 = "BOIDC0B1XXX"
            'Case 2:     sCol = "G": sCodOpe = "301000"
            'Case 7:     sCol = "H": sCodOpe = "400000"
            'Case 8:     sCol = "I": sCodOpe = "400000"
            'Case 9:     sCol = "J": sCodOpe = "400000"
            'Case 10:     sCol = "K": sCodOpe = "401000"
            'Case 11:     sCol = "L": sCodOpe = "500000"
            'Case 12:     sCol = "M": sCodOpe = "500000"
            'Case 13:     sCol = "N": sCodOpe = "501000"
            Case 1:     sCol = "C": sCodOpe = "300100": sCodOpe2 = "BOIDC0B1XXX"
            Case 2:     sCol = "D": sCodOpe = "300110"
        End Select

    i = 28
    For i = 28 To 28 + nDias - 1
        
        sCad = ""
        sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00")
        sCad = sCad & sCodOpe2
        sCad = sCad & "00"
        'sCad = sCad & CStr(xlHoja.Range(sCol & "15")) & Replace(Space(11 - Len(CStr(xlHoja.Range(sCol & "15")))), " ", "0")
        sNumero = Format(CDbl(xlhoja.Range(sCol & i)), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        If sNumero = "" Then
            sNumero = ""
        Else
            If sNumero = "00000000000000" Then
                sNumero = ""
            Else
                sNumero = Format(sNumero, "00000000000000")
            End If
        End If
        If sNumero <> "" Then
           If sCol = "C" Then
            Print #1, sCad; sNumero; "K"; Format(Replace(xlhoja.Range(sCol & "27"), ".", ""), "00"); Format(xlhoja.Range(sCol & "25"), "YYYYMMDD"); Format(xlhoja.Range(sCol & "26"), "YYYYMMDD")
           Else
            Print #1, sCad; sNumero
           End If
        End If
    Next i
    sCodOpe2 = ""
Next j
'***ALPA***********
Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub


'FIN
'**********************************


'For I = 16 To 16 + nDias - 1
'    j = 2
'    While Not xlHoja.Range(ExcelColumnaString(j) & "10") = ""
'        sCad = ""
'        sCad = xlHoja.Range(ExcelColumnaString(j) & 10) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlHoja.Range("A" & I), "00") & IIf(xlHoja.Range(ExcelColumnaString(j) & 15) = "", Space(11), xlHoja.Range(ExcelColumnaString(j) & 15))
'        sCad = sCad & "03"
'        sNumero = Format(xlHoja.Range(ExcelColumnaString(j) & I), "000000000000.00")
'        sNumero = Replace(sNumero, ".", "")
'        If sNumero = "" Then
'            sNumero = "00000000000000"
'        Else
'            sNumero = Format(sNumero, "00000000000000")
'        End If
'        Print #1, sCad; sNumero
'    j = j + 1
'    Wend
'Next I
'Close #1
'
'ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
'MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
'
'Exit Sub

ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub GeneraAnx04MN(psFecha As String)
'''Dim psArchivoALeer As String
'''Dim psArchivoAGrabar As String
'''Dim xlAplicacion As Excel.Application
'''Dim xlLibro As Excel.Workbook
'''Dim xlHoja As Excel.Worksheet
'''Dim bExiste As Boolean
'''Dim bEncontrado As Boolean
'''Dim fs As New Scripting.FileSystemObject
'''
'''Dim I As Integer
'''Dim J As Integer
'''Dim sCad As String
'''
'''Dim Fecha As Date
'''
'''On Error GoTo ErrBegin
'''
''''Verifica el Archivo de Excel que se va a cargar
'''psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "04MN.xls"
'''bExiste = fs.FileExists(psArchivoALeer)
'''
'''If bExiste = False Then
'''    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
'''    Exit Sub
'''End If
'''    psArchivoAGrabar = App.path & "\SPOOLER\00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00.txt"
'''    Set xlAplicacion = New Excel.Application
'''    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
'''    bEncontrado = False
'''    For Each xlHoja In xlLibro.Worksheets
'''        If UCase(xlHoja.Name) = UCase("Anx4") Then
'''            bEncontrado = True
'''            xlHoja.Activate
'''            Exit For
'''        End If
'''    Next
'''    If bEncontrado = False Then
'''        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
'''        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
'''        Exit Sub
'''    End If
''''Creacion del Archivo
'''Open psArchivoAGrabar For Output As #1
'''Dim nDias As Integer
'''Dim sCol As String
'''Dim sCodOpe As String
'''Dim sNumero
'''Print #1, "00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
'''sCad = ""
'''nDias = CInt(Mid(psFecha, 1, 2))
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "04MN.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX4") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo
'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sNumero
Print #ArcSal, "00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "00"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
'j = 2
'For I = 16 To 16 + nDias - 1
'    j = 2
'    While Not xlHoja.Range(ExcelColumnaString(j) & "10") = ""
'        sCad = ""
'        sCad = xlHoja.Range(ExcelColumnaString(j) & 10) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlHoja.Range("A" & I), "00") & IIf(xlHoja.Range(ExcelColumnaString(j) & 15) = "", "   ", xlHoja.Range(ExcelColumnaString(j) & 15))
'
'        sCad = sCad & "01"
'        sNumero = Format(xlHoja.Range(ExcelColumnaString(j) & I), "000000000000.00")
'        sNumero = Replace(sNumero, ".", "")
'        If sNumero = "" Or sNumero = "00000000000000" Then
'            sNumero = "00000000000000"
'        Else
'            sNumero = Format(sNumero, "00000000000000")
'            Print #1, sCad; sNumero
'        End If
'        j = j + 1
'    Wend
'Next I
'sNumero = Format(xlHoja.Range("S50"), "000000000000.00")
'sNumero = Replace(sNumero, ".", "")
'sCad = "800000" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "3000000"
'Print #1, sCad; sNumero
'Close #1
'
'ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
'MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
'
'Exit Sub

'INICIO
'********************************
Dim sCodOpe2 As String
Dim nSaltar As Integer
Dim sCadenaTemp As String
sCodOpe2 = ""

For j = 1 To 7
        Select Case j
           
''''            Case 8:      sCol = "I": sCodOpe = "100000": sCodOpe2 = "001"
''''            Case 9:      sCol = "J": sCodOpe = "100000"
''''            Case 10:     sCol = "K": sCodOpe = "200000"
''''            Case 11:     sCol = "L": sCodOpe = "300000"
''''            Case 12:     sCol = "M": sCodOpe = "300000"
''''            Case 13:     sCol = "N": sCodOpe = "500000"
''''            Case 14:     sCol = "O": sCodOpe = "500000"
''''            Case 15:     sCol = "P": sCodOpe = "700000"
            Case 1:      sCol = "G": sCodOpe = "040000"
                                nSaltar = 1
            Case 2:      sCol = "H": sCodOpe = "041000"
                                nSaltar = 1
            'Case 2:      sCol = "I": sCodOpe = "040000"
            Case 3:      sCol = "O": sCodOpe = "100000"
                                nSaltar = 1
            Case 4:      sCol = "P": sCodOpe = "101000"
                                nSaltar = 1
            'Case 4:      sCol = "Q": sCodOpe = "100000"
            Case 5:      sCol = "S": sCodOpe = "300000"
                                nSaltar = 1
            Case 6:      sCol = "U": sCodOpe = "500000"
                                nSaltar = 1
            Case 7:      sCol = "W": sCodOpe = "700000"
                                nSaltar = 1
            'Case 6:      sCol = "U": sCodOpe = "900000"
            
        End Select
        
     If sCol <> "" And nSaltar = 1 Then
        i = 23
        For i = 23 To 23 + nDias - 1
            sCad = ""
            If (sCol = "O") Then
                sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00") & "001"
'            ElseIf (sCol = "G") Or (sCol = "H") Or (sCol = "P") Then
'                sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlHoja.Range("A" & I)), "00") & "000"
            Else
                sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00") & "000"
            End If
            sCad = sCad & sCodOpe2
        
'            If j = 2 Then
'                sCad = sCad & "001"
'            Else
'                sCad = sCad & "000"
'            End If

            sCad = sCad & "00"
            sNumero = Format(CDbl(xlhoja.Range(sCol & i)), "000000000000.00")
            sNumero = Replace(sNumero, ".", "")
            
            If (sCol = "G") Or (sCol = "O") Then
                If sCol = "G" Then
                    sCadenaTemp = Replace(Format(xlhoja.Range("G22"), "00.00"), ".", "") & Replace(Format(xlhoja.Range("G20"), "YYYYMMDD"), ".", "") & Replace(Format(xlhoja.Range("G21"), "YYYYMMDD"), ".", "")
                End If
                If sCol = "O" Then
                    sCadenaTemp = Replace(Format(xlhoja.Range("O22"), "00.00"), ".", "") & Replace(Format(xlhoja.Range("O20"), "YYYYMMDD"), ".", "") & Replace(Format(xlhoja.Range("O21"), "YYYYMMDD"), ".", "")
                End If
            End If
            
            If sNumero = "" Then
                sNumero = ""
            Else
                If sNumero = "00000000000000" Then
                    sNumero = ""
                Else
                    sNumero = Format(sNumero, "00000000000000")
                End If
            End If
            If sNumero <> "" Then
               Print #1, sCad; sNumero; sCadenaTemp
            End If
            sCadenaTemp = ""
         Next i
         If j = 7 Then
            sCad = "800000" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "3100000"
            sNumero = Format(xlhoja.Range("X53"), "000000000000.00")
            sNumero = Replace(sNumero, ".", "")
            If sNumero = "" Or sNumero = "00000000000000" Then
            Else
                sNumero = Format(sNumero, "00000000000000")
                Print #1, sCad; sNumero
            End If
        End If
     End If
     sCodOpe2 = ""
     nSaltar = 0
Next j

sNumero = Format(CDbl(xlhoja.Range("P51")), "000000000000.00")
sNumero = Replace(sNumero, ".", "")

'sCad = "800000" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Mid(psFecha, 1, 2) & "00"
'Print #1, sCad; sNumero
'ALPA
Close #ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub
'FIN
'********************************************

ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub


Public Sub GeneraAnx04ME(psFecha As String)
''Dim psArchivoALeer As String
''Dim psArchivoAGrabar As String
''Dim xlAplicacion As Excel.Application
''Dim xlLibro As Excel.Workbook
''Dim xlHoja As Excel.Worksheet
''Dim bExiste As Boolean
''Dim bEncontrado As Boolean
''Dim fs As New Scripting.FileSystemObject
''
''Dim I As Integer
''Dim J As Integer
''Dim sCad As String
''
''Dim Fecha As Date
''
''On Error GoTo ErrBegin
''
'''Verifica el Archivo de Excel que se va a cargar
''psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "04ME.xls"
''bExiste = fs.FileExists(psArchivoALeer)
''
''If bExiste = False Then
''    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
''    Exit Sub
''End If
''    psArchivoAGrabar = App.path & "\SPOOLER\00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
''    Set xlAplicacion = New Excel.Application
''    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
''    bEncontrado = False
''    For Each xlHoja In xlLibro.Worksheets
''        If UCase(xlHoja.Name) = UCase("Anx4") Then
''            bEncontrado = True
''            xlHoja.Activate
''            Exit For
''        End If
''    Next
''    If bEncontrado = False Then
''        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
''        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
''        Exit Sub
''    End If
'''Creacion del Archivo
''Open psArchivoAGrabar For Output As #1
''Dim nDias As Integer
''Dim sCol As String
''Dim sCodOpe As String
''Dim sNumero
''Print #1, "00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
''sCad = ""
''nDias = CInt(Mid(psFecha, 1, 2))
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "04ME.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX4") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo
'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sNumero
Print #ArcSal, "00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
'For I = 16 To 16 + nDias - 1
'    j = 2
'    While Not xlHoja.Range(ExcelColumnaString(j) & "10") = ""
'        sCad = ""
'        sCad = xlHoja.Range(ExcelColumnaString(j) & 10) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlHoja.Range("A" & I), "00") & IIf(xlHoja.Range(ExcelColumnaString(j) & 15) = "", "   ", xlHoja.Range(ExcelColumnaString(j) & 15))
'
'        sCad = sCad & "03"
'        sNumero = Format(xlHoja.Range(ExcelColumnaString(j) & I), "000000000000.00")
'        sNumero = Replace(sNumero, ".", "")
'        If sNumero = "" Or sNumero = "00000000000000" Then
'            sNumero = "00000000000000"
'        Else
'            sNumero = Format(sNumero, "00000000000000")
'            Print #1, sCad; sNumero
'        End If
'        j = j + 1
'    Wend
'Next I
'
'Close #1
'
'ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
'MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
'
'Exit Sub

'INICIO
'********************************//
Dim nSaltar As Integer
Dim sCadenaTemp As String
For j = 1 To 3
        Select Case j
           
            Case 1:      sCol = "G": sCodOpe = "040000"
                            nSaltar = 1
            Case 2:      sCol = "H": sCodOpe = "041000"
                            nSaltar = 1
            'Case 2:     sCol = "M": sCodOpe = "040000"
            '                nSaltar = 1
            Case 3:      sCol = "V": sCodOpe = "700000"
                            nSaltar = 1
            'Case 4:     sCol = "S": sCodOpe = "900000"
            
        End Select
        
     If sCol <> "" And nSaltar = 1 Then
        i = 23
        For i = 23 To 23 + nDias - 1
            sCad = ""
             If (sCol = "G") Then
                sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00") & "000"
             Else
                sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00") & "000"
            End If
                
        
'            If j = 2 Then
'                sCad = sCad & "001"
'            Else
'                sCad = sCad & "000"
'            End If
          
                If sCol = "G" Then
                    sCadenaTemp = Replace(Format(xlhoja.Range("G22"), "00.00"), ".", "") & Replace(Format(xlhoja.Range("G20"), "YYYYMMDD"), ".", "") & Replace(Format(xlhoja.Range("G21"), "YYYYMMDD"), ".", "")
                End If
          
            sCad = sCad & "03"
            sNumero = Format(CDbl(xlhoja.Range(sCol & i)), "000000000000.00")
            sNumero = Replace(sNumero, ".", "")
            If sNumero = "" Then

                sNumero = ""
            Else
                If sNumero = "00000000000000" Then
                    sNumero = ""
                Else
                    sNumero = Format(sNumero, "00000000000000")
                End If
            End If
            If sNumero <> "" Then
               Print #1, sCad; sNumero; sCadenaTemp
            End If
            sCadenaTemp = ""
         Next i
         nSaltar = 0
         sCol = ""
     End If
Next j

'sNumero = Format(CDbl(xlHoja.Range("R52")), "000000000000.00")
'sNumero = Replace(sNumero, ".", "")
'sCad = "800000" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "3000000"
'Print #1, sCad; sNumero


'sNumero = Format(CDbl(xlHoja.Range("P51")), "000000000000.00")
'sNumero = Replace(sNumero, ".", "")
'
'sCad = "800000" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Mid(psFecha, 1, 2) & "03"
'Print #1, sCad; sNumero

'ALPA
Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub
'FIN
'********************************************


ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub


Public Sub GeneraAnx03ME(psFecha As String)
''Dim psArchivoALeer As String
''Dim psArchivoAGrabar As String
''Dim xlAplicacion As Excel.Application
''Dim xlLibro As Excel.Workbook
''Dim xlHoja As Excel.Worksheet
''Dim bExiste As Boolean
''Dim bEncontrado As Boolean
''Dim fs As New Scripting.FileSystemObject
''
''Dim I As Integer
''Dim J As Integer
''Dim sCad As String
''
''Dim Fecha As Date
''
''On Error GoTo ErrBegin
''
'''Verifica el Archivo de Excel que se va a cargar
''psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "03ME.xls"
''bExiste = fs.FileExists(psArchivoALeer)
''
''If bExiste = False Then
''    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
''    Exit Sub
''End If
''    psArchivoAGrabar = App.path & "\SPOOLER\00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
''    Set xlAplicacion = New Excel.Application
''    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
''    bEncontrado = False
''    For Each xlHoja In xlLibro.Worksheets
''        If UCase(xlHoja.Name) = UCase("Anx3") Then
''            bEncontrado = True
''            xlHoja.Activate
''            Exit For
''        End If
''    Next
''    If bEncontrado = False Then
''        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
''        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
''        Exit Sub
''    End If
'''Creacion del Archivo
''Open psArchivoAGrabar For Output As #1
''Dim nDias As Integer
''Dim sCol As String
''Dim sCodOpe As String
''Dim sNumero
''Print #1, "00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
''sCad = ""
''nDias = CInt(Mid(psFecha, 1, 2))
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "03ME.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX3") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo
'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sCodOpe2 As String
sCodOpe2 = ""
Dim sNumero
Dim nSalto As Integer
Print #ArcSal, "00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
'For I = 16 To 16 + nDias - 1
'    j = 2
'    While Not xlHoja.Range(ExcelColumnaString(j) & "10") = ""
'        sCad = ""
'        sCad = xlHoja.Range(ExcelColumnaString(j) & 10) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlHoja.Range("A" & I), "00") & IIf(xlHoja.Range(ExcelColumnaString(j) & 15) = "", Space(11), xlHoja.Range(ExcelColumnaString(j) & 15))
'        sCad = sCad & "03"
'        sNumero = Format(xlHoja.Range(ExcelColumnaString(j) & I), "000000000000.00")
'        sNumero = Replace(sNumero, ".", "")
'        If sNumero = "" Then
'            sNumero = "00000000000000"
'        Else
'            sNumero = Format(sNumero, "00000000000000")
'        End If
'        Print #1, sCad; sNumero
'    j = j + 1
'    Wend
'Next I
'
'Close #1
'
'ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
'MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
'
'Exit Sub

'INICIO
'***********************************
'Dim sCodOpe2 As String
'Dim nSalto As Integer
For j = 1 To 2
nSalto = 0
        Select Case j
            'Case 1:     sCol = "B": sCodOpe = "100000"
            'Case 2:     sCol = "C": sCodOpe = "101000"
            'Case 3:     sCol = "D": sCodOpe = "200000"
            'Case 4:     sCol = "E": sCodOpe = "201000"
            '*********************
            'Case 1:     sCol = "H": sCodOpe = "300000": sCodOpe2 = "ICROESM1XXX"
            Case 1:     sCol = "B": nSalto = 1
                                    sCodOpe = "400100": sCodOpe2 = "ICROESM1XXX"
            '*********************
            '*********************
            Case 2:     sCol = "D": nSalto = 1
                                    sCodOpe = "400110"
            '*********************
            'Case 7:     sCol = "H": sCodOpe = "400000"
            'Case 8:     sCol = "I": sCodOpe = "400000"
            'Case 9:     sCol = "J": sCodOpe = "400000"
            'Case 10:     sCol = "K": sCodOpe = "401000"
            'Case 11:     sCol = "L": sCodOpe = "500000"
            'Case 12:     sCol = "M": sCodOpe = "500000"
            'Case 13:     sCol = "N": sCodOpe = "501000"
        End Select
If nSalto = 1 Then
    i = 28
    For i = 28 To 28 + nDias - 1
        
        sCad = ""
        sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00")
        sCad = sCad & sCodOpe2
        sCad = sCad & "03"
        'sCad = sCad & CStr(xlHoja.Range(sCol & "15")) & Replace(Space(11 - Len(CStr(xlHoja.Range(sCol & "15")))), " ", "0")
        If sCol = "D" Then
             sCad = sCad & Space(11)
        End If
             sCad = sCad & "00"
        sNumero = Format(CDbl(xlhoja.Range(sCol & i)), "000000000000.00")
'        sNumero = Format(CDbl(xlHoja.Range(sCol & I)), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        If sNumero = "" Then
            sNumero = ""
        Else
            If sNumero = "00000000000000" Then
                sNumero = "" & IIf(sCol = "D", "T", "")
            Else
                sNumero = Format(sNumero, "00000000000000") & IIf(sCol = "D", "T", "")
            End If
        End If
        If sNumero <> "" Then
           'Print #1, sCad; sNumero
           If sCol = "B" Then
            Print #1, sCad; sNumero; "K"; Format(Replace(xlhoja.Range("B27"), ".", ""), "0000"); Format(xlhoja.Range("B25"), "YYYYMMDD"); Format(xlhoja.Range("B26"), "YYYYMMDD")
           Else
            Print #1, sCad; sNumero
           End If
        End If
    Next i
End If
    sCodOpe2 = ""
Next j
'ALPA
Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub


'FIN
'**********************************


ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub



Public Sub GeneraAnx02ME(psFecha As String)
'''Dim psArchivoALeer As String
'''Dim psArchivoAGrabar As String
'''Dim xlAplicacion As Excel.Application
'''Dim xlLibro As Excel.Workbook
'''Dim xlHoja As Excel.Worksheet
'''Dim bExiste As Boolean
'''Dim bEncontrado As Boolean
'''Dim fs As New Scripting.FileSystemObject
'''
'''Dim I As Integer
'''Dim J As Integer
'''Dim sCad As String
'''Dim nSalta As Integer
'''Dim Fecha As Date
'''
'''On Error GoTo ErrBegin
'''
''''Verifica el Archivo de Excel que se va a cargar
'''psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "02ME.xls"
'''bExiste = fs.FileExists(psArchivoALeer)
'''
'''If bExiste = False Then
'''    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
'''    Exit Sub
'''End If
'''    psArchivoAGrabar = App.path & "\SPOOLER\p00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
'''    Set xlAplicacion = New Excel.Application
'''    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
'''    bEncontrado = False
'''    For Each xlHoja In xlLibro.Worksheets
'''        If UCase(xlHoja.Name) = UCase("ANX2") Then
'''            bEncontrado = True
'''            xlHoja.Activate
'''            Exit For
'''        End If
'''    Next
'''    If bEncontrado = False Then
'''        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
'''        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
'''        Exit Sub
'''    End If
''''Creacion del Archivo
'''Open psArchivoAGrabar For Output As #1
'''Dim nDias As Integer
'''Dim sCol As String
'''Dim sCodOpe As String
'''Dim sNumero
'''Print #1, "00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
'''sCad = ""
'''nDias = CInt(Mid(psFecha, 1, 2))
Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim nSalta As Integer
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "02ME.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX2") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo

'Open psArchivoAGrabar For Output As #1
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sNumero
Print #ArcSal, "00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
'INICIO
'*******************ALPA
Dim sCodOpe2 As String
sCodOpe2 = ""

For j = 1 To 28

     Select Case j
''''            Case 1:     sCol = "B": sCodOpe = "100000": sCodOpe2 = "COFDPEPLXXX"
''''            Case 2:     sCol = "C": sCodOpe = "100000": sCodOpe2 = "BSUDPEPLXXX"
''''            Case 3:     sCol = "D": sCodOpe = "100000": sCodOpe2 = "BCPLPEPLXXX"
''''            Case 4:     sCol = "E": sCodOpe = "200000"
''''            Case 5:     sCol = "F": sCodOpe = "100000": sCodOpe2 = "0051PE01XXX"
''''            Case 6:     sCol = "G": sCodOpe = "101000"
''''
''''            Case 7:     sCol = "H": sCodOpe = "300000": sCodOpe2 = "0801PE73XXX"
''''            Case 8:     sCol = "I": sCodOpe = "300000": sCodOpe2 = "0802PE44XXX"
''''            Case 9:     sCol = "J": sCodOpe = "300000": sCodOpe2 = "0803PE54XXX"
''''
''''            Case 10:     sCol = "K": sCodOpe = "300000": sCodOpe2 = "0805PE73XXX"
''''            Case 11:     sCol = "L": sCodOpe = "300000": sCodOpe2 = "0807PE43XXX"
''''            Case 12:     sCol = "M": sCodOpe = "300000": sCodOpe2 = "0806PE84XXX"
''''            Case 13:     sCol = "N": sCodOpe = "300000": sCodOpe2 = "0808PE64XXX"
''''
''''            Case 14:     sCol = "O": sCodOpe = "300000": sCodOpe2 = "0809PE56XXX"
''''            Case 15:     sCol = "P": sCodOpe = "300000": sCodOpe2 = "0810PE73XXX"
''''            Case 16:     sCol = "Q": sCodOpe = "300000": sCodOpe2 = "0812PE56XXX"
''''
''''            Case 17:     sCol = "R": sCodOpe = "300000": sCodOpe2 = "0813PE52XXX"
''''            Case 18:     sCol = "S": sCodOpe = "300000": sCodOpe2 = "0814PE56XXX"
''''
''''
''''            Case 19:     sCol = "T": sCodOpe = "301000"
''''
''''            Case 20:     sCol = "U": sCodOpe = "300000"
''''            Case 21:     sCol = "V": sCodOpe = "300000"
''''
''''            Case 22:     sCol = "W": sCodOpe = "300000": sCodOpe2 = "BSUDPEPLXXX"
''''            Case 23:     sCol = "X": sCodOpe = "300000": sCodOpe2 = "0051PE01XXX"
''''            Case 24:     sCol = "Y": sCodOpe = "300000": sCodOpe2 = "0832PE56XXX"
''''            Case 25:     sCol = "Z": sCodOpe = "300000"
''''            Case 26:     sCol = "AA": sCodOpe = "301000"
''''
''''            Case 27:     sCol = "AB": sCodOpe = "401000"

            Case 1:     sCol = "B": nSalta = 0: sCodOpe = "100000": sCodOpe2 = ""
            Case 2:     sCol = "C": nSalta = 0: sCodOpe = "100000": sCodOpe2 = ""
            Case 3:     sCol = "D": nSalta = 0: sCodOpe = "101000": sCodOpe2 = ""
            Case 4:     sCol = "E": nSalta = 0: sCodOpe = "200000": sCodOpe2 = ""
            Case 5:     sCol = "F": nSalta = 0: sCodOpe = "200000": sCodOpe2 = ""
            Case 6:     sCol = "G": nSalta = 0: sCodOpe = "201000": sCodOpe2 = ""
            'CMAC
            Case 7:     sCol = "H": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0801PE73XXX" '
            Case 8:     sCol = "I": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0802PE44XXX" '
            Case 9:     sCol = "J": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0803PE54XXX" '
            Case 10:     sCol = "K": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0805PE73XXX"
            Case 11:     sCol = "L": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0807PE43XXX"
            Case 12:     sCol = "M": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0806PE84XXX"
            Case 13:     sCol = "N": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0808PE64XXX"
            Case 14:     sCol = "O": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0809PE56XXX"
            Case 15:     sCol = "P": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0810PE73XXX"
            Case 16:     sCol = "Q": nSalta = 0: sCodOpe = "300000": sCodOpe2 = "0812PE56XXX"
            Case 17:     sCol = "R": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0813PE52XXX"
            Case 18:     sCol = "U": nSalta = 1: sCodOpe = "301000"
            Case 19:     sCol = "V": nSalta = 0: sCodOpe = "400000"
            Case 20:     sCol = "W": nSalta = 0: sCodOpe = "401000"
            
            'Case 18:     sCol = "S": nSalta = 0: sCodOpe = "300000": sCodOpe2 = "0814PE56XXX"
            'Case 19:     sCol = "T": nSalta = 1: sCodOpe = "300000": sCodOpe2 = "0832PE56XXX"
'            Case 20:     sCol = "U": nSalta = 1: sCodOpe = "301000"
'            Case 21:     sCol = "V": nSalta = 0: sCodOpe = "400000"
'            Case 22:     sCol = "W": nSalta = 0: sCodOpe = "401000"
            
        End Select

If nSalta = 1 Then
    i = 21
    For i = 21 To 21 + nDias - 1
        sCad = ""
        sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(CInt(xlhoja.Range("A" & i)), "00")
        sCad = sCad & sCodOpe2
        
        'sCad = sCad & CStr(xlHoja.Range(ssCol & "15")) & Replace(Space(11 - Len(CStr(xlHoja.Range(sCol & "15")))), " ", "0")
       
        If sCol = "D" Or sCol = "U" Then
            sCad = sCad & Space(11)
            sCad = sCad & "03"
        Else
            sCad = sCad & "03"
        End If
        sNumero = Format(CDbl(xlhoja.Range(sCol & i)), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        If sNumero = "" Then
            sNumero = ""
        Else
            If sNumero = "00000000000000" Then
                sNumero = ""
            Else
                sNumero = Format(sNumero, "00000000000000")
            End If
        End If
        If sNumero <> "" Then
           Print #1, sCad; sNumero
        End If
    Next i
    sCodOpe2 = ""
End If
Next j
Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub


'******************
'FIN




'For I = 16 To 16 + nDias - 1
'        j = 2
'        While Not xlHoja.Range(ExcelColumnaString(j) & "10") = ""
'            sCad = ""
'            sCad = xlHoja.Range(ExcelColumnaString(j) & 10) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlHoja.Range("A" & I), "00") & IIf(xlHoja.Range(ExcelColumnaString(j) & 15) = "", Space(11), xlHoja.Range(ExcelColumnaString(j) & 15))
'            sCad = sCad & "03"
'            sNumero = Format(xlHoja.Range(ExcelColumnaString(j) & I), "000000000000.00")
'            sNumero = Replace(sNumero, ".", "")
'            If sNumero = "" Then
'                sNumero = "00000000000000"
'            Else
'                sNumero = Format(sNumero, "00000000000000")
'            End If
'            Print #1, sCad; sNumero
'    j = j + 1
'    Wend
'Next I
'Close #1


ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub

ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub


Public Sub GeneraAnx01ME(psFecha As String)

Dim psArchivoALeer As String
Dim psArchivoAGrabar As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlhoja As Excel.Worksheet
Dim bExiste As Boolean
Dim bEncontrado As Boolean
Dim fs As New Scripting.FileSystemObject

Dim i As Integer
Dim j As Integer
Dim sCad As String

Dim Fecha As Date

On Error GoTo ErrBegin

'Verifica el Archivo de Excel que se va a cargar
psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "01ME.xls"
bExiste = fs.FileExists(psArchivoALeer)

If bExiste = False Then
    MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
    Exit Sub
End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03.txt"
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("Anx1") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
'Creacion del Archivo
'Open psArchivoAGrabar For Output As #
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal

Dim nDias As Integer
Dim sCol As String
Dim sCodOpe As String
Dim sNumero
Print #ArcSal, "00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & "03"
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))

Dim nSalto As Integer
nSalto = 0

For j = 1 To 25
nSalto = 0
      Select Case j
            '*** PEAC 20100614
            Case 1:     sCol = "B": nSalto = 1: sCodOpe = "001000"
            Case 2:     sCol = "C": nSalto = 0: sCodOpe = "002000"
            Case 3:     sCol = "D": nSalto = 0: sCodOpe = "003000"
            Case 4:     sCol = "E": nSalto = 1: sCodOpe = "004000"
            Case 5:     sCol = "F": nSalto = 0: sCodOpe = "007000"
            Case 6:     sCol = "G": nSalto = 1: sCodOpe = "009000"
            Case 7:     sCol = "H": nSalto = 1: sCodOpe = "040000"
            Case 8:     sCol = "I": nSalto = 1: sCodOpe = "050000"
            Case 9:     sCol = "J": nSalto = 0: sCodOpe = "060000"
            Case 10:     sCol = "K": nSalto = 0: sCodOpe = "061000"
            Case 11:     sCol = "L": nSalto = 0: sCodOpe = "062000"
            Case 12:     sCol = "M": nSalto = 0: sCodOpe = "066000"
            Case 13:     sCol = "N": nSalto = 0: sCodOpe = "005600"
            Case 14:     sCol = "O": nSalto = 0: sCodOpe = "067000"
            Case 15:     sCol = "P": nSalto = 1: sCodOpe = "071000"
            Case 16:     sCol = "Q": nSalto = 1: sCodOpe = "065100"
            Case 17:     sCol = "R": nSalto = 1: sCodOpe = "065200"
            Case 18:     sCol = "S": nSalto = 0: sCodOpe = "065000"
            Case 19:     sCol = "T": nSalto = 0: sCodOpe = "068000"
            Case 20:     sCol = "U": nSalto = 0: sCodOpe = "066600"
            Case 20:     sCol = "V": nSalto = 0: sCodOpe = "073000"
            Case 21:     sCol = "W": nSalto = 1: sCodOpe = "070000"
            Case 22:     sCol = "X": nSalto = 0: sCodOpe = "200000"
            Case 23:     sCol = "Y": nSalto = 1: sCodOpe = "085000"
            Case 24:     sCol = "Z": nSalto = 1: sCodOpe = "090000"
            Case 25:     sCol = "AA": nSalto = 1: sCodOpe = "100000"
            
            Case Else
                GoTo Siguiente
        End Select
        
        If nSalto = 1 Then
        i = 19
        For i = 19 To 19 + nDias - 1
        
            sCad = sCodOpe & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlhoja.Range("A" & i), "00") & "03"
            sNumero = Format(xlhoja.Range(sCol & i), "000000000000.00")
            sNumero = Replace(sNumero, ".", "")
            If sNumero = "" Or sNumero = "00000000000000" Then
                sNumero = "00000000000000"
            Else
                sNumero = Format(sNumero, "00000000000000")
                If sNumero <> "" Then
                   Print #1, sCad; sNumero
                End If
            End If
Siguiente:
           
    Next i
    End If
Next j

Close ArcSal

ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub

ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"

End Sub



Private Sub CmdAnx01ME_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx01ME
End Sub

Private Sub CmdAnx01MN_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx01MN
End Sub


Private Sub CmdAnx02ME_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx02ME
End Sub

Private Sub CmdAnx02MN_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx02MN
End Sub


Private Sub CmdAnx03ME_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx03ME
End Sub

Private Sub CmdAnx03MN_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx03MN
End Sub


Private Sub CmdAnx04ME_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx04ME
End Sub

Private Sub CmdAnx04MN_Click()
If ValFecha(txtFecha) = False Then
    Me.txtFecha.SetFocus
    Exit Sub
End If
'GeneraAnx04MN
End Sub

'**** Pasi 20140305 TI-ERS102-2013
Public Sub GeneraTxt01(pnMoneda As Integer, psFecha As String)
    Dim psArchivoALeer As String
    Dim psArchivoAGrabar As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlhoja As Excel.Worksheet
    Dim bExiste As Boolean
    Dim bEncontrado As Boolean
    Dim fs As New Scripting.FileSystemObject
    
    Dim i As Integer
    Dim j As Integer
    Dim sCad As String
    
    Dim Fecha As Date
    
    On Error GoTo ErrBegin
    psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "01" & IIf(pnMoneda = 1, "MN", "ME") & ".xlsx"
    bExiste = fs.FileExists(psArchivoALeer)
    
    If bExiste = False Then
        MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03") & ".txt"
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX1") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If

Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sNumero
Dim oValoresCab As ADODB.Recordset
Set oValoresCab = New ADODB.Recordset
Dim oRep As DRepFormula
Set oRep = New DRepFormula
Dim ntotcol As Integer
Dim m, N As Integer
Dim lnfil As Integer
Print #ArcSal, "00350100811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03")
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))

Set oValoresCab = oRep.ObtenerCabeceraRep1BaseFormulaBCR(pnMoneda)
ntotcol = oValoresCab.RecordCount
RSClose oValoresCab
m = 1
N = 1
lnfil = 15
Do While N <= ntotcol
    Do While m <= nDias
        sCad = xlhoja.Cells(lnfil, N + 1) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlhoja.Range("A" & CStr(lnfil + m)), "00") & IIf(pnMoneda = 1, "00", "03")
        sNumero = Format(xlhoja.Cells(lnfil + m, N + 1), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        If sNumero <> "" Or sNumero <> "00000000000000" Then
            sNumero = Format(sNumero, "00000000000000")
            Print #1, sCad; sNumero
        End If
        m = m + 1
    Loop
    m = 1
    N = N + 1
Loop
Close ArcSal
ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub

    
ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub GeneraTxt02(pnMoneda As Integer, psFecha As String)
    Dim psArchivoALeer As String
    Dim psArchivoAGrabar As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlhoja As Excel.Worksheet
    Dim bExiste As Boolean
    Dim bEncontrado As Boolean
    Dim fs As New Scripting.FileSystemObject
    
    Dim i As Integer
    Dim j As Integer
    Dim sCad As String
    
    Dim Fecha As Date
    
    On Error GoTo ErrBegin
    
    psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "02" & IIf(pnMoneda = 1, "MN", "ME") & ".xlsx"
    bExiste = fs.FileExists(psArchivoALeer)
        
    If bExiste = False Then
        MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03") & ".txt"
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX2") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If

Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sNumero
Dim oValoresCab As ADODB.Recordset
Set oValoresCab = New ADODB.Recordset
Dim oRep As DRepFormula
Set oRep = New DRepFormula
Dim ntotcol As Integer
Dim m, N As Integer
Dim lnfil As Integer

Print #ArcSal, "00350200811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03")
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
Set oValoresCab = oRep.ObtenerCabeceraRep2FormulaBCR(pnMoneda)
ntotcol = oValoresCab.RecordCount
RSClose oValoresCab
m = 1
N = 1
lnfil = 11

Do While N <= ntotcol
    Do While m <= nDias
        sCad = xlhoja.Cells(lnfil, N + 1) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlhoja.Range("A" & CStr(lnfil + m + 2)), "00") & IIf(xlhoja.Cells(lnfil - 1, N + 1) = "", Space(11), xlhoja.Cells(lnfil - 1, N + 1)) & IIf(pnMoneda = 1, "00", "03")
        sNumero = Format(xlhoja.Cells(lnfil + m + 2, N + 1), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        If sNumero <> "" Or sNumero <> "00000000000000" Then
            sNumero = Format(sNumero, "00000000000000")
            Print #1, sCad; sNumero
        End If
        m = m + 1
    Loop
    m = 1
    N = N + 1
Loop
Close ArcSal
ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub
ErrBegin:
  ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Public Sub GeneraTxt03(pnMoneda As Integer, psFecha)
    Dim psArchivoALeer As String
    Dim psArchivoAGrabar As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlhoja As Excel.Worksheet
    Dim bExiste As Boolean
    Dim bEncontrado As Boolean
    Dim fs As New Scripting.FileSystemObject
    
    Dim sCad As String

    
    On Error GoTo ErrBegin
    psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "03" & IIf(pnMoneda = 1, "MN", "ME") & ".xlsx"
    bExiste = fs.FileExists(psArchivoALeer)
    If bExiste = False Then
        MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03") & ".txt"
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX3") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    
Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sNumero
Dim oValoresCab As ADODB.Recordset
Set oValoresCab = New ADODB.Recordset
Dim oValoresCab2 As ADODB.Recordset
Set oValoresCab2 = New ADODB.Recordset
Dim oRep As DRepFormula
Set oRep = New DRepFormula
Dim ntotcol As Integer
Dim m, N As Integer
Dim lnfil As Integer
Dim sDest As String
Dim sPlazProm
Dim sFecha As String

Print #ArcSal, "00350300811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03")
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
Set oValoresCab = oRep.ObtenerSubColRep3FormulaBCR(2, pnMoneda)
Set oValoresCab2 = oRep.ObtenerSubColRep3FormulaBCR(1, pnMoneda)
If Not oValoresCab.EOF And Not oValoresCab.BOF Then
    ntotcol = CInt(oValoresCab.RecordCount)
End If
If Not oValoresCab2.EOF And Not oValoresCab2.BOF Then
    ntotcol = ntotcol + CInt(oValoresCab2.RecordCount)
End If
RSClose oValoresCab
RSClose oValoresCab2
m = 1
N = 1
lnfil = 14

Do While N <= ntotcol
    Do While m <= nDias
        sCad = xlhoja.Cells(lnfil, N + 1) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlhoja.Range("A" & CStr(lnfil + m + 6)), "00") & IIf(xlhoja.Cells(lnfil + 1, N + 1) = "", Space(11), xlhoja.Cells(lnfil + 1, N + 1)) & IIf(pnMoneda = 1, "00", "03")
        sDest = IIf(xlhoja.Cells(lnfil + 2, N + 1) = "", Space(1), Mid(xlhoja.Cells(lnfil + 2, N + 1), 1, 1))
        sPlazProm = xlhoja.Cells(lnfil + 6, N + 1)
        If sPlazProm <> "" Then
            sPlazProm = Format(sPlazProm, "00.00")
            sPlazProm = Replace(sPlazProm, ".", "")
        Else
            sPlazProm = Space(4)
        End If
        sNumero = Format(xlhoja.Cells(lnfil + m + 6, N + 1), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        
        If sNumero <> "" Or sNumero <> "00000000000000" Then
            sNumero = Format(sNumero, "00000000000000")
            Print #1, sCad; sNumero; sDest; sPlazProm; IIf(xlhoja.Cells(lnfil + 4, N + 1) = "", Space(8), Format(xlhoja.Cells(lnfil + 4, N + 1), "yyyyMMdd")); IIf(xlhoja.Cells(lnfil + 5, N + 1) = "", Space(8), Format(xlhoja.Cells(lnfil + 5, N + 1), "yyyyMMdd"))
        End If
        m = m + 1
    Loop
    m = 1
    N = N + 1
Loop

Close ArcSal
ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
Exit Sub
ErrBegin:
    ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
Public Sub GeneraTxt04(pnMoneda As Integer, psFecha)
    Dim psArchivoALeer As String
    Dim psArchivoAGrabar As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlhoja As Excel.Worksheet
    Dim bExiste As Boolean
    Dim bEncontrado As Boolean
    Dim fs As New Scripting.FileSystemObject
    
    'Dim i, j As Integer
    Dim sCad As String
    Dim Fecha As Date
    
    On Error GoTo ErrBegin
    
    psArchivoALeer = App.path & "\Spooler\C1" & Mid(psFecha, 4, 2) & "04" & IIf(pnMoneda = 1, "MN", "ME") & ".xlsx"
    bExiste = fs.FileExists(psArchivoALeer)
    
    If bExiste = False Then
        MsgBox "Ud debe generar previamente el reporte " & Chr(13) & psArchivoALeer, vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    psArchivoAGrabar = App.path & "\SPOOLER\00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03") & ".txt"
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    
    For Each xlhoja In xlLibro.Worksheets
        If UCase(xlhoja.Name) = UCase("ANX4") Then
            bEncontrado = True
            xlhoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True
        MsgBox "No existen datos con la fecha especificada", vbExclamation, "Aviso!!!"
        Exit Sub
    End If

Dim ArcSal As Integer
ArcSal = FreeFile
Open psArchivoAGrabar For Output As ArcSal
Dim nDias As Integer
Dim sNumero
Dim oValoresCab As ADODB.Recordset
Set oValoresCab = New ADODB.Recordset
Dim oRep As DRepFormula
Set oRep = New DRepFormula
Dim ntotcol As Integer
Dim m, N As Integer
Dim lnfil As Integer
Dim sPlazProm
Dim sFecha As String

Print #ArcSal, "00350400811" & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & IIf(pnMoneda = 1, "00", "03")
sCad = ""
nDias = CInt(Mid(psFecha, 1, 2))
Set oValoresCab = oRep.ObtenerCabeceraRep4FormulaBCR(pnMoneda)
ntotcol = oValoresCab.RecordCount
RSClose oValoresCab
m = 1
N = 1
lnfil = 14


Do While N <= ntotcol
    Do While m <= nDias
        sCad = xlhoja.Cells(lnfil, N + 1) & Mid(psFecha, 7, 4) & Mid(psFecha, 4, 2) & Format(xlhoja.Range("A" & CStr(lnfil + m + 6)), "00") & Format(xlhoja.Cells(lnfil + 95, N + 1), "000") & IIf(pnMoneda = 1, "00", "03")
        sPlazProm = xlhoja.Cells(lnfil + 6, N + 1)
        If sPlazProm <> "" Then
            sPlazProm = Format(sPlazProm, "00.00")
            sPlazProm = Replace(sPlazProm, ".", "")
        Else
            sPlazProm = Space(4)
        End If
        sNumero = Format(xlhoja.Cells(lnfil + m + 6, N + 1), "000000000000.00")
        sNumero = Replace(sNumero, ".", "")
        
        If sNumero <> "" Or sNumero <> "00000000000000" Then
            If Len(sNumero) = 15 Then
                sNumero = Format(sNumero, "0000000000000")
            Else
                sNumero = Format(sNumero, "00000000000000")
            End If
            Print #1, sCad; sNumero; sPlazProm; IIf(xlhoja.Cells(lnfil + 4, N + 1) = "", Space(8), Format(xlhoja.Cells(lnfil + 4, N + 1), "yyyyMMdd")); IIf(xlhoja.Cells(lnfil + 5, N + 1) = "", Space(8), Format(xlhoja.Cells(lnfil + 5, N + 1), "yyyyMMdd"))
        End If
        m = m + 1
    Loop
    m = 1
    N = N + 1
Loop
Close ArcSal
ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, False
MsgBox "Archivo generado satisfactoriamente" & Chr(13) & Chr(13) & " en " & App.path & "\SPOOLER\", vbInformation, "Aviso!!!"
 
Exit Sub
ErrBegin:
    ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlhoja, True

    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
'*** end Pasi

Private Sub Form_Load()

End Sub
