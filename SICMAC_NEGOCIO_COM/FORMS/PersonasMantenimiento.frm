VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Personas: Mantenimiento"
   ClientHeight    =   11175
   ClientLeft      =   1290
   ClientTop       =   2445
   ClientWidth     =   12420
   Icon            =   "PersonasMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboAutoriazaUsoDatos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "PersonasMantenimiento.frx":030A
      Left            =   10560
      List            =   "PersonasMantenimiento.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   226
      Top             =   7800
      Width           =   930
   End
   Begin VB.ComboBox cboOfCumplimiento 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "PersonasMantenimiento.frx":034E
      Left            =   10560
      List            =   "PersonasMantenimiento.frx":0358
      Style           =   2  'Dropdown List
      TabIndex        =   207
      Top             =   7440
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.ComboBox cboSujetoObligado 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "PersonasMantenimiento.frx":0392
      Left            =   10560
      List            =   "PersonasMantenimiento.frx":039C
      Style           =   2  'Dropdown List
      TabIndex        =   205
      Top             =   7080
      Width           =   930
   End
   Begin VB.ComboBox cboPEPS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "PersonasMantenimiento.frx":03D6
      Left            =   10560
      List            =   "PersonasMantenimiento.frx":03E0
      Style           =   2  'Dropdown List
      TabIndex        =   203
      Top             =   6720
      Width           =   930
   End
   Begin TabDlg.SSTab SSTabs 
      Height          =   2655
      Left            =   120
      TabIndex        =   81
      Top             =   480
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "Persona &Natural"
      TabPicture(0)   =   "PersonasMantenimiento.frx":041A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPersNatEstCiv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPersNatHijos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPeso"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblTpoSangre"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPersNatSexo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblTalla"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblNacionalidad"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPersNombreAP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPersNombreAM"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblApCasada"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPersNombreN"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblPersNatEmpleados"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label20"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblResidente"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label25"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblTopera"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbPersNatEstCiv"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtPersNatHijos"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtPeso"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmbPersNatSexo"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtTalla"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "CboTipoSangre"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmbNacionalidad"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtPersNombreAP"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtPersNombreAM"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtApellidoCasada"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtPersNombreN"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtPersNatNumEmp"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmbPersNatMagnitud"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cboResidente"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cboPaisReside"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdPerNatDatAdc"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Persona &Juridica"
      TabPicture(1)   =   "PersonasMantenimiento.frx":0436
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPersNombre"
      Tab(1).Control(1)=   "lblPersJurSiglas"
      Tab(1).Control(2)=   "lblPersJurTpo"
      Tab(1).Control(3)=   "lblPersJurMagnitud"
      Tab(1).Control(4)=   "lblPersJurEmpleados"
      Tab(1).Control(5)=   "lblMagnitudEmpresarial"
      Tab(1).Control(6)=   "lblPersJurObjSocial"
      Tab(1).Control(7)=   "txtPersNombreRS"
      Tab(1).Control(8)=   "TxtSiglas"
      Tab(1).Control(9)=   "cmbPersJurTpo"
      Tab(1).Control(10)=   "cmbPersJurMagnitud"
      Tab(1).Control(11)=   "txtPersJurEmpleados"
      Tab(1).Control(12)=   "txtPersJurObjSocial"
      Tab(1).Control(13)=   "cmdAdicJuridica"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "&Relac. c/Pers."
      TabPicture(2)   =   "PersonasMantenimiento.frx":0452
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPersRelacAceptar"
      Tab(2).Control(1)=   "cmdPersRelacCancelar"
      Tab(2).Control(2)=   "cmdPersRelacEditar"
      Tab(2).Control(3)=   "FERelPers"
      Tab(2).Control(4)=   "cmdPersRelacDel"
      Tab(2).Control(5)=   "cmdPersRelacNew"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "&Fte. Ingreso"
      TabPicture(3)   =   "PersonasMantenimiento.frx":046E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdPersFteConsultar"
      Tab(3).Control(1)=   "CmdFteIngNuevo"
      Tab(3).Control(2)=   "CmdFteIngEliminar"
      Tab(3).Control(3)=   "CmdFteIngEditar"
      Tab(3).Control(4)=   "FEFteIng"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Fir&ma"
      TabPicture(4)   =   "PersonasMantenimiento.frx":048A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "IDBFirma"
      Tab(4).Control(1)=   "CmdActFirma"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Ref Com/Per"
      TabPicture(5)   =   "PersonasMantenimiento.frx":04A6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdRefComNuevo"
      Tab(5).Control(1)=   "cmdRefComElimina"
      Tab(5).Control(2)=   "cmdRefComEdita"
      Tab(5).Control(3)=   "cmdRefComCancela"
      Tab(5).Control(4)=   "cmdRefComAcepta"
      Tab(5).Control(5)=   "feRefComercial"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Ref Banc."
      TabPicture(6)   =   "PersonasMantenimiento.frx":04C2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdRefBanAcepta"
      Tab(6).Control(1)=   "cmdRefBanCancela"
      Tab(6).Control(2)=   "cmdRefBanEdita"
      Tab(6).Control(3)=   "cmdRefBanElimina"
      Tab(6).Control(4)=   "cmdRefBanNuevo"
      Tab(6).Control(5)=   "feRefBancaria"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Patrimonio"
      TabPicture(7)   =   "PersonasMantenimiento.frx":04DE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cboMonePatri"
      Tab(7).Control(1)=   "cboTipoPatri"
      Tab(7).Control(2)=   "cmdPatVehAcepta"
      Tab(7).Control(3)=   "cmdPatVehCancela"
      Tab(7).Control(4)=   "cmdPatVehEdita"
      Tab(7).Control(5)=   "cmdPatVehElimina"
      Tab(7).Control(6)=   "cmdPatVehNuevo"
      Tab(7).Control(7)=   "fePatInmuebles"
      Tab(7).Control(8)=   "fePatVehicular"
      Tab(7).Control(9)=   "fePatOtros"
      Tab(7).Control(10)=   "llblMonePatri"
      Tab(7).Control(11)=   "lblTipoPatri"
      Tab(7).ControlCount=   12
      TabCaption(8)   =   "V&isitas"
      TabPicture(8)   =   "PersonasMantenimiento.frx":04FA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "cmdVisitasNuevo"
      Tab(8).Control(1)=   "cmdVisitasEliminar"
      Tab(8).Control(2)=   "cmdVisitasEditar"
      Tab(8).Control(3)=   "cmdVisitasCancelar"
      Tab(8).Control(4)=   "cmdVisitasAceptar"
      Tab(8).Control(5)=   "FEVisitas"
      Tab(8).ControlCount=   6
      TabCaption(9)   =   "Vent Anual"
      TabPicture(9)   =   "PersonasMantenimiento.frx":0516
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "cmdVentasAceptar"
      Tab(9).Control(1)=   "cmdVentasCancelar"
      Tab(9).Control(2)=   "cmdVentasEditar"
      Tab(9).Control(3)=   "cmdVentasEliminar"
      Tab(9).Control(4)=   "cmdVentasNuevo"
      Tab(9).Control(5)=   "FEVentas"
      Tab(9).ControlCount=   6
      Begin VB.CommandButton cmdPerNatDatAdc 
         Caption         =   "Datos &Accionariales"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   230
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdicJuridica 
         Caption         =   "Datos &Accionariales"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67200
         TabIndex        =   229
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtPersJurObjSocial 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70560
         MaxLength       =   200
         TabIndex        =   201
         Top             =   960
         Width           =   5190
      End
      Begin VB.ComboBox cboPaisReside 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0532
         Left            =   8730
         List            =   "PersonasMantenimiento.frx":0534
         Style           =   2  'Dropdown List
         TabIndex        =   199
         Top             =   840
         Width           =   1560
      End
      Begin VB.ComboBox cboResidente 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0536
         Left            =   8760
         List            =   "PersonasMantenimiento.frx":0540
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   930
      End
      Begin VB.ComboBox cmbPersNatMagnitud 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":057A
         Left            =   1560
         List            =   "PersonasMantenimiento.frx":057C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   3000
      End
      Begin VB.CommandButton cmdVentasAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -66135
         TabIndex        =   182
         Top             =   1680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdVentasCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -66135
         TabIndex        =   181
         Top             =   2010
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdVentasEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66135
         TabIndex        =   180
         Top             =   855
         Width           =   1140
      End
      Begin VB.CommandButton cmdVentasEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66135
         TabIndex        =   179
         Top             =   1170
         Width           =   1140
      End
      Begin VB.CommandButton cmdVentasNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66135
         TabIndex        =   178
         Top             =   540
         Width           =   1140
      End
      Begin VB.CommandButton cmdVisitasNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66135
         TabIndex        =   177
         Top             =   540
         Width           =   1140
      End
      Begin VB.CommandButton cmdVisitasEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66135
         TabIndex        =   176
         Top             =   1170
         Width           =   1140
      End
      Begin VB.CommandButton cmdVisitasEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66135
         TabIndex        =   174
         Top             =   855
         Width           =   1140
      End
      Begin VB.CommandButton cmdVisitasCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -66135
         TabIndex        =   173
         Top             =   2010
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdVisitasAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -66135
         TabIndex        =   172
         Top             =   1680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ComboBox cboMonePatri 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":057E
         Left            =   -69480
         List            =   "PersonasMantenimiento.frx":0580
         Style           =   2  'Dropdown List
         TabIndex        =   163
         Top             =   480
         Width           =   1305
      End
      Begin VB.ComboBox cboTipoPatri 
         Height          =   315
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   160
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtPersNatNumEmp 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   8760
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "0"
         Top             =   2280
         Width           =   420
      End
      Begin SICMACT.ImageDB IDBFirma 
         Height          =   1935
         Left            =   -73320
         TabIndex        =   147
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3413
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtPersNombreN 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtApellidoCasada 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtPersNombreAM 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtPersNombreAP 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton cmdPatVehAcepta 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -65805
         TabIndex        =   137
         Top             =   1605
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehCancela 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -65805
         TabIndex        =   136
         Top             =   1905
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65820
         TabIndex        =   135
         Top             =   810
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehElimina 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -65820
         TabIndex        =   134
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdPatVehNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65820
         TabIndex        =   133
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdRefBanAcepta 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -66000
         TabIndex        =   131
         Top             =   1575
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanCancela 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -66000
         TabIndex        =   130
         Top             =   1920
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66000
         TabIndex        =   129
         Top             =   735
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanElimina 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66000
         TabIndex        =   128
         Top             =   1050
         Width           =   990
      End
      Begin VB.CommandButton cmdRefBanNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66000
         TabIndex        =   127
         Top             =   420
         Width           =   990
      End
      Begin VB.CommandButton cmdRefComNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65985
         TabIndex        =   126
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComElimina 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65985
         TabIndex        =   125
         Top             =   1065
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComEdita 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -65985
         TabIndex        =   123
         Top             =   750
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComCancela 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -65970
         TabIndex        =   122
         Top             =   1890
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRefComAcepta 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -65970
         TabIndex        =   121
         Top             =   1590
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbNacionalidad 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0582
         Left            =   5880
         List            =   "PersonasMantenimiento.frx":0584
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1553
         Width           =   1560
      End
      Begin VB.ComboBox CboTipoSangre 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0586
         Left            =   8760
         List            =   "PersonasMantenimiento.frx":0588
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   1125
      End
      Begin VB.CommandButton CmdPersFteConsultar 
         Caption         =   "Consultar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   114
         Top             =   1935
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   -66150
         TabIndex        =   112
         Top             =   1575
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   -66150
         TabIndex        =   111
         Top             =   1905
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   110
         Top             =   540
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngEliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   109
         Top             =   1170
         Width           =   1140
      End
      Begin VB.CommandButton CmdFteIngEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66375
         TabIndex        =   108
         Top             =   855
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66150
         TabIndex        =   107
         Top             =   750
         Width           =   1140
      End
      Begin VB.TextBox TxtTalla 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8760
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1200
         Width           =   870
      End
      Begin VB.ComboBox cmbPersNatSexo 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":058A
         Left            =   5850
         List            =   "PersonasMantenimiento.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   473
         Width           =   1545
      End
      Begin VB.TextBox TxtPeso 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8760
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1920
         Width           =   870
      End
      Begin SICMACT.FlexEdit FEFteIng 
         Height          =   1785
         Left            =   -74775
         TabIndex        =   74
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3149
         Cols0           =   6
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-Tipo-Razon Social-Moneda-CodPersFI-Indice"
         EncabezadosAnchos=   "400-1200-5000-1300-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-L-L-L"
         FormatosEdit    =   "3-1-1-1-0-0"
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit FERelPers 
         Height          =   1785
         Left            =   -74775
         TabIndex        =   71
         Top             =   435
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   3149
         Cols0           =   7
         FixedCols       =   0
         HighLight       =   2
         EncabezadosNombres=   "Item-Codigo-Nombres-Relacion-Beneficiario-Porcentaje-Asist.Med.Priv."
         EncabezadosAnchos=   "400-1300-4000-1400-2000-1200-2500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-5-6"
         ListaControles  =   "0-1-0-3-3-0-3"
         EncabezadosAlineacion=   "C-L-L-L-L-R-L"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         CantEntero      =   3
         TextArray0      =   "Item"
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.CommandButton CmdActFirma 
         Caption         =   "&Actualizar Firma"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -67740
         TabIndex        =   75
         Top             =   600
         Width           =   1380
      End
      Begin VB.CommandButton cmdPersRelacDel 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66150
         TabIndex        =   73
         Top             =   1065
         Width           =   1140
      End
      Begin VB.CommandButton cmdPersRelacNew 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66150
         TabIndex        =   72
         Top             =   435
         Width           =   1140
      End
      Begin VB.TextBox txtPersJurEmpleados 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   69
         Top             =   1800
         Width           =   1380
      End
      Begin VB.ComboBox cmbPersJurMagnitud 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":058E
         Left            =   -68880
         List            =   "PersonasMantenimiento.frx":0590
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1320
         Width           =   3545
      End
      Begin VB.ComboBox cmbPersJurTpo 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0592
         Left            =   -73680
         List            =   "PersonasMantenimiento.frx":0594
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1320
         Width           =   3000
      End
      Begin VB.TextBox TxtSiglas 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         MaxLength       =   15
         TabIndex        =   79
         Top             =   960
         Width           =   1500
      End
      Begin VB.TextBox txtPersNombreRS 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73680
         MaxLength       =   150
         TabIndex        =   67
         Top             =   600
         Width           =   8295
      End
      Begin VB.TextBox txtPersNatHijos 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5865
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "0"
         Top             =   1200
         Width           =   300
      End
      Begin VB.ComboBox cmbPersNatEstCiv 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0596
         Left            =   5850
         List            =   "PersonasMantenimiento.frx":0598
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   833
         Width           =   1560
      End
      Begin SICMACT.FlexEdit feRefComercial 
         Height          =   1785
         Left            =   -74880
         TabIndex        =   124
         Top             =   480
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   3149
         Cols0           =   7
         HighLight       =   2
         EncabezadosNombres=   "#-Nom/Raz.Soc/Entrev/Cargo-Referencia-Comentario-Telefono-Direccion-C"
         EncabezadosAnchos=   "350-4900-1500-4000-1500-1500-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-X"
         ListaControles  =   "0-0-3-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         CantEntero      =   3
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   3
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit feRefBancaria 
         Height          =   1785
         Left            =   -74925
         TabIndex        =   132
         Top             =   480
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   3149
         Cols0           =   7
         HighLight       =   2
         EncabezadosNombres=   "#-Codigo-Referencia Bancaria-Nro Cuenta-Nro Tarjeta-Linea Cred U$-Item"
         EncabezadosAnchos=   "350-1500-3500-1800-1800-1500-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-5-X"
         ListaControles  =   "0-1-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-L-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit fePatInmuebles 
         Height          =   1740
         Left            =   -74880
         TabIndex        =   158
         Top             =   960
         Visible         =   0   'False
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   3069
         Cols0           =   8
         HighLight       =   1
         EncabezadosNombres=   "#-Ubicación-Area terreno-Area construida-Tipo d/uso-R.P.-Valor-C"
         EncabezadosAnchos=   "400-3300-1200-1700-2200-1200-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-6-X"
         ListaControles  =   "0-0-0-0-0-3-0-0"
         EncabezadosAlineacion=   "C-L-R-R-L-L-R-C"
         FormatosEdit    =   "0-0-2-2-0-1-2-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit fePatVehicular 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   138
         Top             =   900
         Visible         =   0   'False
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   3493
         Cols0           =   7
         HighLight       =   1
         EncabezadosNombres=   "#-Marca-Año Fabrica-Valor Comercial-Modelo-Placa-C"
         EncabezadosAnchos=   "400-3300-1200-1700-2200-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-L-L-C"
         FormatosEdit    =   "0-0-3-2-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit fePatOtros 
         Height          =   1980
         Left            =   -74880
         TabIndex        =   159
         Top             =   1140
         Visible         =   0   'False
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   3493
         Cols0           =   4
         HighLight       =   1
         EncabezadosNombres=   "#-Descripción-Valor-C"
         EncabezadosAnchos=   "400-3300-1700-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-X"
         ListaControles  =   "0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-C"
         FormatosEdit    =   "0-0-2-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit FEVisitas 
         Height          =   1785
         Left            =   -74760
         TabIndex        =   175
         Top             =   480
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   3149
         Cols0           =   7
         FixedCols       =   0
         HighLight       =   2
         EncabezadosNombres=   "Item-Codigo-Nombres-Direccion-Fecha-Condicion-Observacion"
         EncabezadosAnchos=   "400-1300-4000-3000-1200-1000-3500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-5-6"
         ListaControles  =   "0-1-0-0-2-3-0"
         EncabezadosAlineacion=   "C-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         CantEntero      =   3
         TextArray0      =   "Item"
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit FEVentas 
         Height          =   1785
         Left            =   -74880
         TabIndex        =   183
         Top             =   480
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   3149
         Cols0           =   7
         HighLight       =   2
         EncabezadosNombres=   "#-Codigo-Auditor EEFF-Monto Vtas Anuales-Fec. Aud-Periodo-A"
         EncabezadosAnchos=   "350-1500-2000-1800-1500-1000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-3-4-5-X"
         ListaControles  =   "0-1-0-0-2-0-0"
         EncabezadosAlineacion=   "C-L-L-R-L-C-C"
         FormatosEdit    =   "0-0-0-2-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         TipoBusqueda    =   3
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label lblTopera 
         Height          =   255
         Left            =   5640
         TabIndex        =   231
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label lblPersJurObjSocial 
         AutoSize        =   -1  'True
         Caption         =   "Objeto Social"
         Height          =   195
         Left            =   -71760
         TabIndex        =   202
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Pais Reside"
         Height          =   195
         Left            =   7680
         TabIndex        =   200
         Top             =   915
         Width           =   840
      End
      Begin VB.Label lblResidente 
         AutoSize        =   -1  'True
         Caption         =   "Residente"
         Height          =   195
         Left            =   7710
         TabIndex        =   197
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Magnitud Persona"
         Height          =   195
         Left            =   120
         TabIndex        =   185
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label llblMonePatri 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   -70200
         TabIndex        =   164
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblTipoPatri 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   161
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblPersNatEmpleados 
         Caption         =   "Nº Empleados"
         Height          =   255
         Left            =   7710
         TabIndex        =   149
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label lblMagnitudEmpresarial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -68880
         TabIndex        =   144
         Top             =   1320
         Width           =   2985
      End
      Begin VB.Label lblPersNombreN 
         Caption         =   "Nombres"
         Height          =   255
         Left            =   240
         TabIndex        =   143
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblApCasada 
         Caption         =   "Apellido Casada"
         Height          =   255
         Left            =   240
         TabIndex        =   142
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblPersNombreAM 
         Caption         =   "Apellido Materno"
         Height          =   255
         Left            =   240
         TabIndex        =   141
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblPersNombreAP 
         Caption         =   "Apellido Paterno"
         Height          =   255
         Left            =   240
         TabIndex        =   140
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblNacionalidad 
         Caption         =   "Nacionalidad"
         Height          =   240
         Left            =   4800
         TabIndex        =   120
         Top             =   1635
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "m."
         Height          =   195
         Left            =   9720
         TabIndex        =   116
         Top             =   1260
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kg."
         Height          =   195
         Left            =   9720
         TabIndex        =   115
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label LblTalla 
         AutoSize        =   -1  'True
         Caption         =   "Talla"
         Height          =   195
         Left            =   7710
         TabIndex        =   106
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label lblPersNatSexo 
         AutoSize        =   -1  'True
         Caption         =   "Sexo"
         Height          =   195
         Left            =   4815
         TabIndex        =   105
         Top             =   533
         Width           =   375
      End
      Begin VB.Label LblTpoSangre 
         AutoSize        =   -1  'True
         Caption         =   "T. Sangre"
         Height          =   195
         Left            =   7710
         TabIndex        =   104
         Top             =   1620
         Width           =   705
      End
      Begin VB.Label lblPeso 
         AutoSize        =   -1  'True
         Caption         =   "Peso "
         Height          =   195
         Left            =   7710
         TabIndex        =   103
         Top             =   1980
         Width           =   405
      End
      Begin VB.Label lblPersJurEmpleados 
         AutoSize        =   -1  'True
         Caption         =   "N° Empleados"
         Height          =   195
         Left            =   -74880
         TabIndex        =   94
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblPersJurMagnitud 
         AutoSize        =   -1  'True
         Caption         =   "Magnitud Empresarial"
         Height          =   195
         Left            =   -70560
         TabIndex        =   92
         Top             =   1360
         Width           =   1515
      End
      Begin VB.Label lblPersJurTpo 
         AutoSize        =   -1  'True
         Caption         =   "Tpo"
         Height          =   195
         Left            =   -74880
         TabIndex        =   91
         Top             =   1360
         Width           =   285
      End
      Begin VB.Label lblPersJurSiglas 
         AutoSize        =   -1  'True
         Caption         =   "Siglas"
         Height          =   195
         Left            =   -74880
         TabIndex        =   90
         Top             =   1000
         Width           =   420
      End
      Begin VB.Label lblPersNombre 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social"
         Height          =   195
         Left            =   -74880
         TabIndex        =   89
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblPersNatHijos 
         AutoSize        =   -1  'True
         Caption         =   "N°Dependientes"
         Height          =   195
         Left            =   4560
         TabIndex        =   85
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label lblPersNatEstCiv 
         AutoSize        =   -1  'True
         Caption         =   "Est. Civil"
         Height          =   195
         Left            =   4815
         TabIndex        =   84
         Top             =   893
         Width           =   615
      End
   End
   Begin VB.CheckBox chkResidente 
      Alignment       =   1  'Right Justify
      Caption         =   "Residente"
      Enabled         =   0   'False
      Height          =   270
      Left            =   8760
      TabIndex        =   198
      Top             =   960
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CommandButton cmdVerFirma 
      Caption         =   "&Ver Firma"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8250
      TabIndex        =   184
      Top             =   75
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ComboBox cboMotivoActu 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "PersonasMantenimiento.frx":059A
      Left            =   10560
      List            =   "PersonasMantenimiento.frx":059C
      Style           =   2  'Dropdown List
      TabIndex        =   167
      Top             =   2520
      Width           =   1560
   End
   Begin VB.TextBox txtFecUltAct 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10560
      TabIndex        =   165
      Top             =   1920
      Width           =   1530
   End
   Begin VB.CommandButton cmdActualizarFirma 
      Caption         =   "&Actualizar Firma"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8250
      TabIndex        =   148
      Top             =   75
      Width           =   1380
   End
   Begin VB.CommandButton CmdPersCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11025
      TabIndex        =   102
      ToolTipText     =   "Cancelar Todos los cambios Realizados"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton CmdPersAceptar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9480
      TabIndex        =   101
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   300
      Left            =   10560
      TabIndex        =   76
      Top             =   480
      Width           =   1000
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   300
      Left            =   10560
      TabIndex        =   77
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   300
      Left            =   10560
      TabIndex        =   78
      Top             =   1200
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTDatosGen 
      Height          =   6975
      Left            =   120
      TabIndex        =   86
      Top             =   3240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos &Generales"
      TabPicture(0)   =   "PersonasMantenimiento.frx":059E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPersNac"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPersCIIU"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPersEstado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblRela"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTipoComp"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblNumPtosVta"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTipoSisInform"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblTipoCadena"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblActiviComple"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblFecIncRuc"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblFecIniActi"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label16"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label18"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label24"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label32"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label33"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtPersFallec"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPersFecIniActi"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtPersFecInscRuc"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbPersEstado"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "CboPersCiiu"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtPersNacCreac"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "CmbRela"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtSbs"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtIngresoProm"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtCodCIIU"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtActiGiro"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cboTipoComp"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtNumPtosVta"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cboTipoSistInfor"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cboCadenaProd"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtActComple"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtNumDependi"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "chkcred"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "chkser"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "chkaho"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "chkotro"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cboocupa"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Frame2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Frame3"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Frame4"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cboCargos"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "cboRemInfoEmail"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "&Domicilio"
      TabPicture(1)   =   "PersonasMantenimiento.frx":05BA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox cboRemInfoEmail 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":05D6
         Left            =   7680
         List            =   "PersonasMantenimiento.frx":05E0
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1580
         Width           =   810
      End
      Begin VB.Frame Frame5 
         Caption         =   " Información Negocio / Centro Laboral "
         Height          =   3255
         Left            =   -74760
         TabIndex        =   209
         Top             =   3480
         Width           =   7725
         Begin VB.TextBox txtNombreCentroLaboral 
            Height          =   405
            Left            =   1620
            TabIndex        =   227
            Top             =   2400
            Width           =   5595
         End
         Begin VB.TextBox txtNegDireccion 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1620
            MaxLength       =   200
            TabIndex        =   59
            Top             =   1560
            Width           =   5595
         End
         Begin VB.TextBox txtRefNegocio 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1620
            MaxLength       =   200
            TabIndex        =   60
            Top             =   1920
            Width           =   5595
         End
         Begin VB.ComboBox cmbNegUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "PersonasMantenimiento.frx":0624
            Left            =   2370
            List            =   "PersonasMantenimiento.frx":0626
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1140
            Width           =   2190
         End
         Begin VB.ComboBox cmbNegUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "PersonasMantenimiento.frx":0628
            Left            =   4800
            List            =   "PersonasMantenimiento.frx":062A
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbNegUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "PersonasMantenimiento.frx":062C
            Left            =   4800
            List            =   "PersonasMantenimiento.frx":062E
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbNegUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "PersonasMantenimiento.frx":0630
            Left            =   2400
            List            =   "PersonasMantenimiento.frx":0632
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbNegUbiGeo 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "PersonasMantenimiento.frx":0634
            Left            =   330
            List            =   "PersonasMantenimiento.frx":0636
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   525
            Width           =   1815
         End
         Begin VB.Label lblNombreDel 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Centro Laboral"
            Height          =   435
            Left            =   240
            TabIndex        =   228
            Top             =   2400
            Width           =   1170
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Dirrección"
            Height          =   195
            Left            =   240
            TabIndex        =   224
            Top             =   1590
            Width           =   720
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Referencia"
            Height          =   195
            Left            =   240
            TabIndex        =   215
            Top             =   1950
            Width           =   780
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2385
            TabIndex        =   214
            Top             =   900
            Width           =   600
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4815
            TabIndex        =   213
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   330
            TabIndex        =   212
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2385
            TabIndex        =   211
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4815
            TabIndex        =   210
            Top             =   285
            Width           =   750
         End
      End
      Begin VB.ComboBox cboCargos 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0638
         Left            =   5160
         List            =   "PersonasMantenimiento.frx":063A
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3360
         Width           =   3340
      End
      Begin VB.Frame Frame4 
         Caption         =   "Correos Electrónicos"
         Height          =   615
         Left            =   200
         TabIndex        =   193
         Top             =   1950
         Width           =   8415
         Begin VB.TextBox TxtEmail2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4920
            MaxLength       =   45
            TabIndex        =   26
            Top             =   200
            Width           =   3345
         End
         Begin VB.TextBox TxtEmail 
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            MaxLength       =   45
            TabIndex        =   25
            Top             =   200
            Width           =   3345
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Email 2:"
            Height          =   195
            Left            =   4200
            TabIndex        =   195
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Email 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   194
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Teléfonos Celulares"
         Height          =   615
         Left            =   200
         TabIndex        =   189
         Top             =   1320
         Width           =   6390
         Begin VB.TextBox txtCel3 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5040
            MaxLength       =   15
            TabIndex        =   23
            Top             =   200
            Width           =   1275
         End
         Begin VB.TextBox txtCel2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2880
            MaxLength       =   15
            TabIndex        =   22
            Top             =   200
            Width           =   1275
         End
         Begin VB.TextBox txtCel1 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   840
            MaxLength       =   15
            TabIndex        =   21
            Top             =   200
            Width           =   1275
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Celular 3:"
            Height          =   195
            Left            =   4320
            TabIndex        =   192
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Celular 2:"
            Height          =   195
            Left            =   2160
            TabIndex        =   191
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Celular 1:"
            Height          =   195
            Left            =   120
            TabIndex        =   190
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Teléfonos Fijos"
         Height          =   600
         Left            =   200
         TabIndex        =   186
         Top             =   720
         Width           =   4815
         Begin VB.TextBox txtPersTelefono2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3360
            MaxLength       =   12
            TabIndex        =   18
            Top             =   200
            Width           =   1275
         End
         Begin VB.TextBox txtPersTelefono 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   17
            Top             =   200
            Width           =   1275
         End
         Begin VB.Label lblPersTelefono2 
            AutoSize        =   -1  'True
            Caption         =   "Trabajo:"
            Height          =   195
            Left            =   2640
            TabIndex        =   188
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblPersTelefono 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio:"
            Height          =   195
            Left            =   120
            TabIndex        =   187
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.ComboBox cboocupa 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":063C
         Left            =   1320
         List            =   "PersonasMantenimiento.frx":063E
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3000
         Width           =   7185
      End
      Begin VB.CheckBox chkotro 
         Caption         =   "Otros"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7680
         TabIndex        =   35
         Top             =   3840
         Width           =   855
      End
      Begin VB.CheckBox chkaho 
         Caption         =   "Ahorr"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6000
         TabIndex        =   33
         Top             =   3840
         Width           =   855
      End
      Begin VB.CheckBox chkser 
         Caption         =   "Servc"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6840
         TabIndex        =   34
         Top             =   3840
         Width           =   855
      End
      Begin VB.CheckBox chkcred 
         Caption         =   "Cred"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5280
         TabIndex        =   32
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtNumDependi 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   7800
         TabIndex        =   44
         Top             =   5610
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtActComple 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   40
         Top             =   5160
         Width           =   4935
      End
      Begin VB.ComboBox cboCadenaProd 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0640
         Left            =   2040
         List            =   "PersonasMantenimiento.frx":0642
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   4680
         Width           =   1785
      End
      Begin VB.ComboBox cboTipoSistInfor 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0644
         Left            =   6000
         List            =   "PersonasMantenimiento.frx":0646
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   4680
         Width           =   2625
      End
      Begin VB.TextBox txtNumPtosVta 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   8040
         TabIndex        =   41
         Top             =   5160
         Width           =   495
      End
      Begin VB.ComboBox cboTipoComp 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0648
         Left            =   2640
         List            =   "PersonasMantenimiento.frx":064A
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   4200
         Width           =   1305
      End
      Begin VB.TextBox txtActiGiro 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   29
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox TxtCodCIIU 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   7080
         MaxLength       =   9
         TabIndex        =   20
         Text            =   "0"
         Top             =   1160
         Width           =   1395
      End
      Begin VB.TextBox txtIngresoProm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   7080
         MaxLength       =   9
         TabIndex        =   19
         Text            =   "0"
         Top             =   780
         Width           =   1395
      End
      Begin VB.TextBox TxtSbs 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   7065
         MaxLength       =   10
         TabIndex        =   16
         Top             =   390
         Width           =   1395
      End
      Begin VB.ComboBox CmbRela 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":064C
         Left            =   5160
         List            =   "PersonasMantenimiento.frx":064E
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   4200
         Width           =   3465
      End
      Begin MSMask.MaskEdBox txtPersNacCreac 
         Height          =   300
         Left            =   1320
         TabIndex        =   14
         Top             =   405
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CboPersCiiu 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0650
         Left            =   1320
         List            =   "PersonasMantenimiento.frx":0652
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2640
         Width           =   7185
      End
      Begin VB.Frame Frame1 
         Caption         =   " Información Domicilio "
         Height          =   2895
         Left            =   -74760
         TabIndex        =   95
         Top             =   480
         Width           =   7725
         Begin VB.TextBox txtPersDireccDomicilio 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   50
            Top             =   1560
            Width           =   6075
         End
         Begin VB.ComboBox cmbPersDireccCondicion 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtRefDomicilio 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   53
            Top             =   2370
            Width           =   6075
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            ItemData        =   "PersonasMantenimiento.frx":0654
            Left            =   330
            List            =   "PersonasMantenimiento.frx":0656
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   525
            Width           =   1815
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            ItemData        =   "PersonasMantenimiento.frx":0658
            Left            =   2370
            List            =   "PersonasMantenimiento.frx":065A
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   540
            Width           =   2175
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "PersonasMantenimiento.frx":065C
            Left            =   4800
            List            =   "PersonasMantenimiento.frx":065E
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1155
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "PersonasMantenimiento.frx":0660
            Left            =   4800
            List            =   "PersonasMantenimiento.frx":0662
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   540
            Width           =   2430
         End
         Begin VB.ComboBox cmbPersUbiGeo 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "PersonasMantenimiento.frx":0664
            Left            =   2355
            List            =   "PersonasMantenimiento.frx":0666
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   1140
            Width           =   2190
         End
         Begin SICMACT.EditMoney txtValComercial 
            Height          =   285
            Left            =   5655
            TabIndex        =   52
            Top             =   1950
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin VB.Label lblPersDireccDomicilio 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            Height          =   195
            Left            =   240
            TabIndex        =   219
            Top             =   1590
            Width           =   630
         End
         Begin VB.Label lblPersDireccCondicion 
            AutoSize        =   -1  'True
            Caption         =   "Condicion"
            Height          =   195
            Left            =   240
            TabIndex        =   218
            Top             =   1995
            Width           =   705
         End
         Begin VB.Label Label13 
            Caption         =   "Valor Comercial U$"
            Height          =   240
            Left            =   4140
            TabIndex        =   217
            Top             =   1995
            Width           =   1440
         End
         Begin VB.Label lblRefDomicilio 
            AutoSize        =   -1  'True
            Caption         =   "Referencia"
            Height          =   195
            Left            =   240
            TabIndex        =   216
            Top             =   2370
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Provincia :"
            Height          =   195
            Left            =   4815
            TabIndex        =   100
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento :"
            Height          =   195
            Left            =   2385
            TabIndex        =   99
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pais : "
            Height          =   195
            Left            =   330
            TabIndex        =   98
            Top             =   285
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Zona : "
            Height          =   195
            Left            =   4800
            TabIndex        =   97
            Top             =   915
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Distrito :"
            Height          =   195
            Left            =   2355
            TabIndex        =   96
            Top             =   900
            Width           =   600
         End
      End
      Begin VB.ComboBox cmbPersEstado 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "PersonasMantenimiento.frx":0668
         Left            =   1320
         List            =   "PersonasMantenimiento.frx":066A
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3720
         Width           =   2625
      End
      Begin MSMask.MaskEdBox txtPersFecInscRuc 
         Height          =   300
         Left            =   2040
         TabIndex        =   42
         Top             =   5610
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPersFecIniActi 
         Height          =   300
         Left            =   5070
         TabIndex        =   43
         Top             =   5610
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPersFallec 
         Height          =   300
         Left            =   3480
         TabIndex        =   15
         Top             =   405
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Info x Email"
         Height          =   195
         Left            =   6720
         TabIndex        =   223
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Remisión de"
         Height          =   195
         Left            =   6720
         TabIndex        =   222
         Top             =   1540
         Width           =   870
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Cargo :"
         Height          =   195
         Left            =   4440
         TabIndex        =   196
         Top             =   3480
         Width           =   510
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Ocupación : "
         Height          =   195
         Left            =   240
         TabIndex        =   171
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Relac CMAC"
         Height          =   195
         Left            =   4200
         TabIndex        =   170
         Top             =   3840
         Width           =   915
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "F. Fallec."
         Height          =   195
         Left            =   2760
         TabIndex        =   169
         Top             =   480
         Width           =   645
      End
      Begin VB.Label lblFecIniActi 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio Actividad:"
         Height          =   195
         Left            =   3360
         TabIndex        =   162
         Top             =   5640
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Nº Dependientes:"
         Height          =   195
         Left            =   6480
         TabIndex        =   157
         Top             =   5640
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lblFecIncRuc 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inscrip.RUC:"
         Height          =   195
         Left            =   240
         TabIndex        =   156
         Top             =   5640
         Width           =   1620
      End
      Begin VB.Label lblActiviComple 
         AutoSize        =   -1  'True
         Caption         =   "Actv. complementarias:"
         Height          =   195
         Left            =   240
         TabIndex        =   155
         Top             =   5160
         Width           =   1650
      End
      Begin VB.Label lblTipoCadena 
         AutoSize        =   -1  'True
         Caption         =   "Tipo cadena productiva:"
         Height          =   195
         Left            =   240
         TabIndex        =   154
         Top             =   4695
         Width           =   1740
      End
      Begin VB.Label lblTipoSisInform 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Sistema Información:"
         Height          =   195
         Left            =   4080
         TabIndex        =   153
         Top             =   4695
         Width           =   1830
      End
      Begin VB.Label lblNumPtosVta 
         AutoSize        =   -1  'True
         Caption         =   "Nº Ptos Vta:"
         Height          =   195
         Left            =   7080
         TabIndex        =   152
         Top             =   5160
         Width           =   870
      End
      Begin VB.Label lblTipoComp 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Competencia que enfrenta:"
         Height          =   195
         Left            =   240
         TabIndex        =   151
         Top             =   4200
         Width           =   2280
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Act/Giro Prnc :"
         Height          =   195
         Left            =   240
         TabIndex        =   150
         Top             =   3480
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "COD CIIU"
         Height          =   195
         Left            =   6240
         TabIndex        =   145
         Top             =   1150
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso Promedio (S/.) "
         Height          =   195
         Left            =   5400
         TabIndex        =   139
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Codigo SBS "
         Height          =   195
         Left            =   5940
         TabIndex        =   119
         Top             =   420
         Width           =   900
      End
      Begin VB.Label lblRela 
         AutoSize        =   -1  'True
         Caption         =   "Relac. Inst :"
         Height          =   195
         Left            =   4200
         TabIndex        =   118
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblPersEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Left            =   240
         TabIndex        =   93
         Top             =   3840
         Width           =   540
      End
      Begin VB.Label lblPersCIIU 
         AutoSize        =   -1  'True
         Caption         =   "CIIU :"
         Height          =   195
         Left            =   240
         TabIndex        =   88
         Top             =   2715
         Width           =   405
      End
      Begin VB.Label lblPersNac 
         AutoSize        =   -1  'True
         Caption         =   "F. Nac/Creac"
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   405
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbPersPersoneria 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3615
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   105
      Width           =   4335
   End
   Begin TabDlg.SSTab SSTIdent 
      Height          =   2340
      Left            =   9045
      TabIndex        =   83
      Top             =   3840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4128
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Identificación"
      TabPicture(0)   =   "PersonasMantenimiento.frx":066C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdPersIDnew"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdPersIDCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPersIDAceptar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPersIDedit"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdPersIDDel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FEDocs"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin SICMACT.FlexEdit FEDocs 
         Height          =   1440
         Left            =   120
         TabIndex        =   62
         Top             =   435
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   2540
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Tipo-Numero"
         EncabezadosAnchos=   "350-1200-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2"
         ListaControles  =   "0-3-0"
         BackColor       =   12648447
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L"
         FormatosEdit    =   "0-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   12648447
      End
      Begin VB.CommandButton cmdPersIDDel 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   65
         Top             =   1920
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDedit 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   64
         Top             =   1920
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDAceptar 
         Caption         =   "Aceptar"
         Height          =   300
         Left            =   240
         TabIndex        =   63
         Top             =   1920
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   1200
         TabIndex        =   113
         Top             =   1920
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPersIDnew 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   300
         Left            =   240
         TabIndex        =   61
         Top             =   1920
         Width           =   885
      End
   End
   Begin VB.PictureBox CdlgImg 
      Height          =   480
      Left            =   11640
      ScaleHeight     =   420
      ScaleWidth      =   435
      TabIndex        =   117
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin SICMACT.TxtBuscar TxtBCodPers 
      Height          =   285
      Left            =   720
      TabIndex        =   146
      Top             =   120
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   503
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   3
      sTitulo         =   ""
   End
   Begin SICMACT.TxtBuscar txtBUsuario 
      Height          =   315
      Left            =   10560
      TabIndex        =   220
      Top             =   3120
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   503
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
   End
   Begin VB.Label lblAutorizarUsoDatos 
      Caption         =   "Autorizar Uso Datos:"
      Height          =   375
      Left            =   9240
      TabIndex        =   225
      Top             =   7770
      Width           =   1215
   End
   Begin VB.Label lblBUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuario"
      Height          =   195
      Left            =   10560
      TabIndex        =   221
      Top             =   2880
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblOfCumplimiento 
      AutoSize        =   -1  'True
      Caption         =   "Of. Cumplimiento:"
      Height          =   195
      Left            =   9240
      TabIndex        =   208
      Top             =   7500
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblSujetoObligado 
      AutoSize        =   -1  'True
      Caption         =   "Sujeto Obligado:"
      Height          =   195
      Left            =   9240
      TabIndex        =   206
      Top             =   7140
      Width           =   1170
   End
   Begin VB.Label lblPEPS 
      AutoSize        =   -1  'True
      Caption         =   "PEPS:"
      Height          =   195
      Left            =   9240
      TabIndex        =   204
      Top             =   6780
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Motivo actualización"
      Height          =   195
      Left            =   10560
      TabIndex        =   168
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblFecUltAct 
      AutoSize        =   -1  'True
      Caption         =   "Ultima actualización:"
      Height          =   195
      Left            =   10560
      TabIndex        =   166
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblPersPersoneria 
      AutoSize        =   -1  'True
      Caption         =   "Personería:"
      Height          =   195
      Left            =   2745
      TabIndex        =   82
      Top             =   165
      Width           =   825
   End
   Begin VB.Label lblPersCod 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   60
      TabIndex        =   80
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "frmPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variables de Excel
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo As String
    Dim lsHoja         As String
    Dim lbLibroOpen As Boolean
'Fin Variables excel
Public rsHojEval As ADODB.Recordset
Dim MatrixHojaEval() As String
Dim nNumeroDoc As Integer 'ALPA 20080922************
Dim MatrixTipoDoc() As String
'***************************************************
Dim nPos As Integer
Dim nDat As Integer
''PTI120170530 según ERS014-2017 agrego
Dim nTipoInicioFteIngreso As Integer
'Fin 'PTI1

Private Enum TTipoCombo
    ComboDpto = 1
    ComboProv = 2
    ComboDist = 3
    ComboZona = 4
End Enum

Enum TPersonaTipoInicio
    PersonaConsulta = 1
    PersonaActualiza = 2
End Enum

Dim NomMoverSSTabs As Integer

'Dim oBuscaPersona As UPersona
'Dim oPersona As DPersona
Dim oBuscaPersona As COMDPersona.UCOMPersona
Dim oPersona As UPersona_Cli   ' COMDPersona.DCOMPersona
Dim olimpiar As New UPersona_Cli 'CTI3
Dim Nivel1() As String
Dim ContNiv1 As Long
Dim Nivel2() As String
Dim ContNiv2 As Long
Dim Nivel3() As String
Dim ContNiv3 As Long
Dim Nivel4() As String
Dim ContNiv4 As Long
Dim Nivel5() As String
Dim ContNiv5 As Long
Dim bEstadoCargando As Boolean

'Para flexEdit de Relacion con Personas
Dim cmdPersRelaEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FERelPersNoMoverdeFila As Integer

'madm 20100326 - Para flexEdit de Visitas
Dim cmdPersVisitasEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FEVisitasPersNoMoverdeFila As Integer

'MAVM 20100607 - Para flexEdit de Ventas
Dim cmdPersVentasEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FEVentasPersNoMoverdeFila As Integer

'Para flexEdit de Documentos
Dim cmdPersDocEjecutado As Integer '1: Nuevo, 2:Editar, 3: Eliminar
Dim FEDocsPersNoMoverdeFila As Integer

'Para Fuentes de Ingreso
Dim cmdPersFteIngresoEjecutado As Integer
Dim FEFtePersNoMoverdeFila As Integer

'Para Ref Comercial
Dim cmdPersRefComercialEjecutado As Integer
Dim FERefComPersNoMoverdeFila As Integer
Dim lnNumRefCom As Integer

'Para Ref Bancaria
Dim cmdPersRefBancariaEjecutado As Integer
Dim FERefBanPersNoMoverdeFila As Integer

'Para PAt Vehicular
Dim cmdPersPatVehicularEjecutado As Integer
Dim FEPatVehPersNoMoverdeFila As Integer
Dim lnNumPatVeh As Integer
'*** PEAC 20080412
Dim lnNumPatOtros As Integer
Dim lnNumPatInmuebles As Integer

Dim BotonEditar As Boolean
Dim BotonNuevo As Boolean

'Para Persona Nueva desde BuscaPersona
Dim sPersCodNombre As String
Dim bBuscaNuevo As Boolean
Public bNuevaPersona As Boolean
Dim bPersonaAct As Boolean
'Dim nPos As Integer
'05-06-2006
Dim bPersonaAGrabar As Boolean

Dim bCIIU As Boolean
Dim RsCIIUTemp As New ADODB.Recordset
Dim lcTexto As String

'*** PEAC 20080412
Dim RsTIPOCOMPTemp As New ADODB.Recordset
Dim RsTIPOSISTINFORTemp As New ADODB.Recordset
Dim RsCADENAPRODTemp As New ADODB.Recordset

'*** FIN PEAC

Dim bConsultaVerFirma As Boolean

Dim nCodPersAuto As Integer 'MADM 20101221
Dim bPermisoEditarTodo As Boolean 'EJVG20111219
Dim bRealizaMantenimiento As Boolean 'EJVG20120323
Dim bUsuarioRealizoMantenimiento As Boolean 'EJVG20120323
Dim fbPermisoEditarSujetoObligadoDJ As Boolean 'EJVG20120815
Dim lnCondicionBNeg As Integer 'WIOR 20121122
Dim lsDireccionActualizada '***Agreado por ELRO el 20130219, según INC1302150010
Dim MatPersona(1 To 2) As TActAutDatos 'FRHU 20151130 ERS077-2015
'WIOR 20130827 *************************************
Private fbPermisoCargo As Boolean
Private fsNombreActual As String
Private rsDocPersActual As ADODB.Recordset
Private rsDocPersUlt As ADODB.Recordset
'WIOR FIN ******************************************
Dim bValidaCampDatos As Boolean 'JUEZ 20131024
Dim codigoOpcionMenu As String 'PTI1 23-06-2017
Dim oHabilitar As COMDConstSistema.DCOMConstSistema 'APRI20170630 TI-ERS025
Dim bHabilitarBoton As Boolean 'APRI20170630 TI-ERS025

'Variables para Obtener valores de las Propiedades de la Persona

'********CTI3
Dim oDetaJuri, oValidaGrilla As UPersona_Cli
Dim nValor As Integer
Dim sGrillaLleno As Boolean
Dim nIngPromedio As String 'ADD PT1 ERS070-2018 14/12/2018
Dim bSensible As Boolean 'ADD PT1 ERS070-2018 14/12/2018
Dim bPemisoAD As Boolean 'ADD PT1 ERS070-2018 14/12/2018
Dim nTipoForm As Integer 'ADD PT1 ERS070-2018 14/12/2018

'Dim odetaJuri As TDatosJuridicos
'********

'->***** LUCV20181220, Anexo01 de Acta 199-2018
Dim objPista As COMManejador.Pista
Dim lsMovNro As String
Dim nTipoAccion As Integer '1: Registro, 2: Mantenimiento, 3: Consulta
'<-***** Fin LUCV20181220

Sub CargaCIIU()
    If Not (RsCIIUTemp.EOF And RsCIIUTemp.BOF) Then
        RsCIIUTemp.MoveFirst
        CboPersCiiu.Clear
        Do While Not RsCIIUTemp.EOF
            CboPersCiiu.AddItem Trim(RsCIIUTemp!cCIIUdescripcion) & Space(100) & Trim(RsCIIUTemp!cCIIUcod)
            RsCIIUTemp.MoveNext
        Loop
                       
    End If
End Sub

'***PEAC 20080412
Sub CargaTIPOCOMP()
    If Not (RsTIPOCOMPTemp.EOF And RsTIPOCOMPTemp.BOF) Then
        RsTIPOCOMPTemp.MoveFirst
        cboTipoComp.Clear
        Do While Not RsTIPOCOMPTemp.EOF
            cboTipoComp.AddItem Trim(RsTIPOCOMPTemp!cCIIUdescripcion) & Space(100) & Trim(RsTIPOCOMPTemp!cCIIUcod)
            RsTIPOCOMPTemp.MoveNext
        Loop
    End If
End Sub

'***PEAC 20080412
Sub CargaTIPOSISTINFOR()
    If Not (RsTIPOSISTINFORTemp.EOF And RsTIPOSISTINFORTemp.BOF) Then
        RsTIPOSISTINFORTemp.MoveFirst
        cboTipoSistInfor.Clear
        Do While Not RsTIPOSISTINFORTemp.EOF
            cboTipoSistInfor.AddItem Trim(RsTIPOSISTINFORTemp!cCIIUdescripcion) & Space(100) & Trim(RsTIPOSISTINFORTemp!cCIIUcod)
            RsTIPOSISTINFORTemp.MoveNext
        Loop
    End If
End Sub
'***PEAC 20080412
Sub CargaCADENAPROD()
    If Not (RsCADENAPRODTemp.EOF And RsCADENAPRODTemp.BOF) Then
        RsCADENAPRODTemp.MoveFirst
        cboCadenaProd.Clear
        Do While Not RsCADENAPRODTemp.EOF
            cboCadenaProd.AddItem Trim(RsCADENAPRODTemp!cCIIUdescripcion) & Space(100) & Trim(RsCADENAPRODTemp!cCIIUcod)
            RsCADENAPRODTemp.MoveNext
        Loop
    End If
End Sub



'Dim fnTipoPersona As Integer
'Dim fsUbicGeografica As String
'Dim fsDomicilio  As String
'Dim fsCondicionDomic As String
'Dim fnValComDomicilio   As Currency
'Dim fsTipoSangre As String
'Dim fsApePat As String
'Dim fsApeMat As String
'Dim fsNombres  As String
'Dim fsNombreCompleto As String
'Dim fnTalla  As Double
'Dim fnPeso   As Double
'Dim fsEmail  As String
'Dim fsTelefonos2 As String
'Dim fsSexo   As String
'Dim fsApeCas As String
'Dim fsEstadoCivil As String
'Dim fnHijos  As Integer
'Dim fcNacionalidad  As String
'Dim fnResidencia As Integer
'Dim fdFechaNac As Date
'Dim fsTelefonos As String
'Dim fsCiiu   As String
'Dim fsEstado As String
'Dim fsSiglas As String
'Dim fsPersCodSbs As String
'Dim fsTipoPersJur As String
'Dim fnPersRelInst  As Integer
'Dim fsMagnitudEmp  As String
'Dim fnNumEmplead As Integer
'Dim fnMaxRefCom  As Integer
'Dim fnMaxPatVeh  As Integer
'Dim fRfirma As ADODB.Recordset
'Dim fsPersCod  As String
'Dim fnNumFtes  As Integer
'Dim fnNumPatVeh As Integer
'Dim fnNumRefCom As Integer
'Dim fsCodAge As String
'Dim fsActualiza  As String
'Dim fdFechaHoy  As Date
'Dim fnTipoAct   As Integer
'Dim fnNumDocs As Integer


Public Sub Registrar()
    Me.Caption = "Personas:Registro"
    nTipoAccion = -1 'LUCV20181220, Anexo01 de Acta 199-2018
    cmdEditar.Enabled = False
    cmdnuevo.Enabled = True
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Caption = "Persona : Registro"
    'EJVG20120813 ***
    'Me.chkpeps.Visible = True 'MADM 20100524
    Me.lblPEPS.Visible = True
    Me.cboPEPS.Visible = True
    TxtBCodPers.Enabled = False
    'END EJVG
    'FRHU 20151130 ERS077-2015
    
    nTipoForm = 1 'add pti1 ers070-2018
    Call ValidaAutorizardatos 'ADD PTI1 ERS070-2018
    
    'lblAutorizarUsoDatos.Visible = False 'comentado por pti1 ers070-2018
    'cboAutoriazaUsoDatos.Visible = False 'comentado por pti1 ers070-2018
    
    'FIN FRHU
    bPermisoEditarTodo = True
    nTipoInicioFteIngreso = -1 ''PTI120170530 según ERS014-2017
    nTipoAccion = 1 'LUCV20181220, Anexo01 de Acta 199-2018
    gsOpeCod = gPersonaRegistro 'LUCV20181220, Anexo01 de Acta 199-2018
    Me.Show 1
End Sub

Public Sub Mantenimeinto()
    Me.Caption = "Personas:Mantenimiento"
    nTipoAccion = -1 'LUCV20181220, Anexo01 de Acta 199-2018
    cmdEditar.Enabled = True
    cmdnuevo.Enabled = False
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    Me.Caption = "Persona : Mantenimiento"
    'EJVG20120813 ***
    'Me.chkpeps.Visible = False 'MADM 20100524
    Me.lblPEPS.Visible = False
    Me.cboPEPS.Visible = False
    'END EJVG *******
    
    'ADD PTI1 ERS070-2018 26/12/2018
    lblAutorizarUsoDatos.Visible = True
    CboAutoriazaUsoDatos.Visible = True
    nTipoForm = 2
    'FIN add pti1 ers070-2018
    
    nTipoInicioFteIngreso = -1 ''PTI120170530 según ERS014-2017
    nTipoAccion = 2 'LUCV20181220, Anexo01 de Acta 199-2018
    gsOpeCod = gPersonaMantenimiento 'LUCV20181220, Anexo01 de Acta 199-2018
    Me.Show 1
End Sub

Public Sub Consultar()
    Me.Caption = "Personas:Consulta"
    nTipoAccion = -1 'LUCV20181220, Anexo01 de Acta 199-2018
    cmdEditar.Enabled = False
    cmdnuevo.Enabled = False
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    bConsultaVerFirma = True
    
    'ADD PTI1 ERS070-2018 16/01/2019
    lblAutorizarUsoDatos.Visible = True
    CboAutoriazaUsoDatos.Visible = True
    nTipoForm = 3
    'FIN add pti1 ers070-2018 16/01/2019
    
    nTipoAccion = 3 'LUCV20181220, Anexo01 de Acta 199-2018
    gsOpeCod = gPersonaConsulta 'LUCV20181220, Anexo01 de Acta 199-2018
    Me.Show 1
    nTipoInicioFteIngreso = -1 'PTI120170530 según ERS014-2017
End Sub
'*****-> PTI120170530 según ERS014-2017
Public Sub ConsultarFteIngreso(Optional ByVal codigoOpcion As String)
    codigoOpcionMenu = codigoOpcion 'PTI120170530 según ERS014-2017
    nTipoInicioFteIngreso = 1
    Me.Caption = "Personas:Consulta Fuente de Ingreso"
    cmdEditar.Enabled = False
    cmdnuevo.Enabled = False
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    bConsultaVerFirma = True
    nTipoForm = 4 'ADD pti1 ers070-2018 26/12/2018
    gsOpeCod = gCredConsultarFteIngreso 'LUCV20181220, Anexo01 de Acta 199-2018
    Me.Show 1
    nTipoInicioFteIngreso = -1
End Sub
'<-*****  Fin PTI120170530

Public Function PersonaNueva() As String
    On Error GoTo LblErrorNew
    'ADD PTI1 16/01/2019 LPDP ACTA 199-2018
    Me.Caption = "Personas: Registro"
    nTipoForm = 1
    Call ValidaAutorizardatos
    nIngPromedio = ""
    'END PTI1 LPDP
    
    
    sPersCodNombre = ""
    bBuscaNuevo = True
    
    'ARCV 08-06-2006
    TxtBCodPers.Enabled = False
    'GITU 18-06-2008 Se comento la funcion HabilitaControles_BotonNuevo porque
    ' el codigo ciiu devolvia en blanco y se agrego el evento clic del cmdnuevo
    ' poruqe ahi si llena el codigo ciiu
    ' dentro de ese evento tambien esta la funcion HabilitaControles_BotonNuevo
    'Call HabilitaControles_BotonNuevo
    nTipoAccion = -1 'LUCV20181220, Anexo01 de Acta 199-2018
    Call cmdNuevo_Click
    '-------------
    nTipoAccion = 1 'LUCV20181220, Anexo01 de Acta 199-2018
    gsOpeCod = gPersonaRegistro 'LUCV20181220, Anexo01 de Acta 199-2018
    Me.Show 1
    'Call cmdNuevo_Click
    PersonaNueva = sPersCodNombre
    Exit Function
LblErrorNew:
    Select Case Err.Number
        Case 400
            MsgBox "No se Puede Crear una Persona desde el Mantenimiento de Persona", vbInformation, "AVISO"
        Case Else
            MsgBox Err.Description, vbInformation, "AVISO"
    End Select
    
End Function

Public Sub Inicio(ByVal cPersCod As String, ByVal pTipoInicio As TPersonaTipoInicio)
    
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona
    TxtBCodPers.Text = cPersCod
    
    bValidaCampDatos = False 'JUEZ 20131024
'    Call oPersona.RecuperaPersona(Trim(TxtBCodPers.Text))
'    If oPersona.PersCodigo = "" Then
'        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
'        Exit Sub
'    End If
'    Call CargaDatos
    If Cargar_Datos_Persona(Trim(TxtBCodPers.Text)) = False Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "AVISO"
        Exit Sub
    End If
    
    TxtBCodPers.Enabled = False
    cmdnuevo.Enabled = False
    
    If pTipoInicio = PersonaActualiza Then
        cmdEditar.Enabled = True
    Else
        cmdEditar.Enabled = False
    End If
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    
    Me.Show 1
End Sub

'EJVG20120323
Public Function realizarMantenimiento(ByVal psPersCod As String, Optional ByRef psDireccionActualizada As String = "") As Boolean
    '***Parametro psDireccion agreado por ELRO el 20130219, según INC1302150010
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli
    TxtBCodPers.Text = psPersCod
    bValidaCampDatos = False 'JUEZ 20131024
    
    If Cargar_Datos_Persona(Trim(TxtBCodPers.Text)) = False Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "AVISO"
        realizarMantenimiento = False
        Exit Function
    End If
    
    TxtBCodPers.Enabled = False
    cmdnuevo.Enabled = False
    cmdEditar.Enabled = True
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
        
    Call cmdEditar_Click
    oPersona.TipoActualizacion = PersFilaModificada
    bRealizaMantenimiento = True
    bUsuarioRealizoMantenimiento = False
    Me.Show 1
    '***Agreado por ELRO el 20130219, según INC1302150010
    psDireccionActualizada = lsDireccionActualizada
    '***Fin Agreado por ELRO el 20130219*****************
    realizarMantenimiento = bUsuarioRealizoMantenimiento
End Function

Private Function ValidaDatosDocumentos() As Boolean
Dim i, J As Integer
Dim bEnc As Boolean

'MADM 20110120 - VISTO
'***JGPA20191210 Comentó según ACTA Nº 106-2019
'Dim loVistoElectronico As frmVistoElectronico
'Set loVistoElectronico = New frmVistoElectronico
'***End JGPA20191210
'END

    ValidaDatosDocumentos = True
    
        
    '**DAOR 20071011, Según Memorandum JTI-062-2007-CMACM******
    If CInt(Right(cmbPersPersoneria.Text, 2)) <> gPersonaNat Then
        If CInt(Right(FEDocs.TextMatrix(FEDocs.row, 1), 2)) <> gPersIdRUC Then 'MADM 20100826
    '    If CInt(Right(FEDocs.TextMatrix(FEDocs.Row, 1), 2)) = gPersIdDNI Then
            MsgBox "No es posible ingresar Documentos diferentes al RUC para tipo de persona Jurídica", vbInformation, "AVISO"
            If FEDocs.Enabled And SSTIdent.Enabled Then FEDocs.SetFocus
            ValidaDatosDocumentos = False
            Exit Function
        End If
    End If
    '**********************************************************
    
    'Verifica el Tipo de Documento
    If Len(Trim(FEDocs.TextMatrix(FEDocs.row, 1))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Documento", vbInformation, "AVISO"
        If FEDocs.Enabled And SSTIdent.Enabled Then FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    'Verifica el Numero de Documento
    If Len(Trim(FEDocs.TextMatrix(FEDocs.row, 2))) = 0 Then
        MsgBox "Falta Ingresar el Numero de Documento", vbInformation, "AVISO"
        FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    'Verifica Duplicidad de Documento
    bEnc = False
    For i = 1 To FEDocs.rows - 2
        If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
            For J = i + 1 To FEDocs.rows - 1
                If Trim(Right(FEDocs.TextMatrix(i, 1), 20)) = Trim(Right(FEDocs.TextMatrix(J, 1), 20)) Then
                    bEnc = True
                    Exit For
                End If
            Next J
        End If
        If bEnc Then
            Exit For
        End If
    Next i
    If bEnc Then
        MsgBox "Existe un Documento Duplicado", vbInformation, "Aviso"
        FEDocs.SetFocus
        ValidaDatosDocumentos = False
        Exit Function
    End If
    
    '******************************************************************
    '*** PEAC 20090720 - BUSCA SI NUMERO DE DOC ESTA EN LISTA NEGATIVA
    If BuscaNumDocEnListaNegativa(CInt(Right(FEDocs.TextMatrix(FEDocs.row, 1), 2)), Trim(FEDocs.TextMatrix(FEDocs.row, 2)), lnCondicionBNeg) Then 'WIOR 20121122 AGREGO lnCondicionBNeg
       'marg 12-05-2016
       Dim denominacion As String
       Select Case lnCondicionBNeg
            Case 1
                denominacion = "NEGATIVO"
            Case 2
                 denominacion = "FRAUDULENTOS"
            Case 3
                denominacion = "PEPS"
            Case 4
               denominacion = "VINCULADOS-NT"
            Case 5
               denominacion = "LISTA OFAC"
            Case 6
                denominacion = "LISTA ONU"
             Case 7
                denominacion = "PEPS - NEGATIVO"
        End Select
       
       
        MsgBox "El número de Documento de Identidad que acaba de ingresar está registrado como " & denominacion, 48, "Atención"
        'MADM 20110115-------------------------------------------------------------------------------------------
        '***JGPA20191210 Comentó según ACTA Nº 106-2019
        'ValidaDatosDocumentos = loVistoElectronico.Inicio(3)
        'If Not ValidaDatosDocumentos Then
           'Exit Function
        'End If
        '***End JGPA20191210
        'END MADM -------------------------------------------------------------------------------------------
    End If
    
    '******************************************************************
    'Verfica Tipo de Valores de Documento
    For i = 1 To FEDocs.rows - 1
        If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
            If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdDNI Then
                For J = 1 To Len(Trim(FEDocs.TextMatrix(i, 2)))
                    If (Mid(FEDocs.TextMatrix(i, 2), J, 1) < "0" Or Mid(FEDocs.TextMatrix(i, 2), J, 1) > "9") Then
                       MsgBox "Uno de los Digitos del DNI no es un Numero", vbInformation, "AVISO"
                       FEDocs.SetFocus
                       ValidaDatosDocumentos = False
                       Exit Function
                    End If
                Next J
            End If
            If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC Then
                 For J = 1 To Len(Trim(FEDocs.TextMatrix(i, 2)))
                    If (Mid(FEDocs.TextMatrix(i, 2), J, 1) < "0" Or Mid(FEDocs.TextMatrix(i, 2), J, 1) > "9") Then
                       MsgBox "Uno de los Digitos del RUC no es un Numero", vbInformation, "AVISO"
                       FEDocs.SetFocus
                       ValidaDatosDocumentos = False
                       Exit Function
                    End If
                Next J
            End If
        End If
    Next i
    
    'Verfica Longitud de Documento
    '******************************************************************
    For i = 1 To FEDocs.rows - 1
        If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
            If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdDNI Then
                If Len(Trim(FEDocs.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                    MsgBox "DNI No es de " & gnNroDigitosDNI & " digitos", vbInformation, "AVISO"
                    FEDocs.SetFocus
                    ValidaDatosDocumentos = False
                    Exit Function
                End If
            End If
            If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC Then
                If Len(Trim(FEDocs.TextMatrix(i, 2))) <> gnNroDigitosRUC Then
                    MsgBox "RUC No es de " & gnNroDigitosRUC & " digitos", vbInformation, "AVISO"
                    FEDocs.SetFocus
                    ValidaDatosDocumentos = False
                    Exit Function
                End If
            End If
        End If
    Next i
    
End Function

Private Function ValidaDatosPersRelacion() As Boolean
Dim i As Integer

    ValidaDatosPersRelacion = True
    
    'Valida Titular No Este como Relacion
    For i = 1 To Me.FERelPers.rows - 1
        If FERelPers.TextMatrix(i, 1) = TxtBCodPers.Text Then
            MsgBox "No se puede Agregar al Titular en la Relacion de Personas", vbInformation, "AVISO"
            ValidaDatosPersRelacion = False
            FERelPers.SetFocus
            Exit Function
        End If
    Next i
    
    'Falta Persona a Relacionar
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 1))) = 0 Then
        MsgBox "Falta Ingresar la Persona con la que se va a Relacionar", vbInformation, "AVISO"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta Ingresar el tipo de Relacion
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 3))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Relacion con la Persona", vbInformation, "AVISO"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta el Tipo de Beneficio
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 4))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Beneficio con la Persona", vbInformation, "AVISO"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
    'Falta ingresar el tipo de Asistenacia Medica Privada AMP
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 6))) = 0 Then
        MsgBox "Falta Ingresar el Tipo Asist. Med. Privada", vbInformation, "AVISO"
        ValidaDatosPersRelacion = False
        Exit Function
    End If
    
End Function
'madm 20100327
Private Function ValidaDatosPersVisita() As Boolean
Dim i As Integer

    ValidaDatosPersVisita = True
    
    'Valida Titular No Este como Relacion
    For i = 1 To Me.FEVisitas.rows - 1
        If FEVisitas.TextMatrix(i, 1) = TxtBCodPers.Text Then
            MsgBox "No se puede Agregar al Titular en la Visita", vbInformation, "AVISO"
            ValidaDatosPersVisita = False
            FEVisitas.SetFocus
            Exit Function
        End If
    Next i
    
    If Len(Trim(FEVisitas.TextMatrix(FEVisitas.row, 1))) = 0 Then
        MsgBox "Falta Ingresar la Persona que realizó la visita", vbInformation, "AVISO"
        ValidaDatosPersVisita = False
        Exit Function
    End If
    
    If Len(Trim(FEVisitas.TextMatrix(FEVisitas.row, 3))) = 0 Then
        MsgBox "Falta Ingresar la Fecha de Visita", vbInformation, "AVISO"
        ValidaDatosPersVisita = False
        Exit Function
    End If
    
    If Len(Trim(FEVisitas.TextMatrix(FEVisitas.row, 4))) = 0 Then
        MsgBox "Falta Ingresar el Tipo de Visita", vbInformation, "Aviso"
        ValidaDatosPersVisita = False
        Exit Function
    End If
    
    If Len(Trim(FEVisitas.TextMatrix(FEVisitas.row, 5))) = 0 Then
        MsgBox "Falta Ingresar el Comentario generado", vbInformation, "Aviso"
        ValidaDatosPersVisita = False
        Exit Function
    End If
    
End Function
'end madm

'MAVM 20100607 BAS II
Private Function ValidaDatosPersVentas() As Boolean
Dim i As Integer

    ValidaDatosPersVentas = True
    
    'Valida Titular No Este como Relacion
    For i = 1 To Me.FEVentas.rows - 1
        If FEVentas.TextMatrix(i, 1) = TxtBCodPers.Text Then
            MsgBox "No se puede Agregar al Titular en la Lista", vbInformation, "Aviso"
            ValidaDatosPersVentas = False
            FEVentas.SetFocus
            Exit Function
        End If
    Next i
    
    If Len(Trim(FEVentas.TextMatrix(FEVentas.row, 1))) = 0 Then
        MsgBox "Falta Ingresar la Persona que realizó la Auditoria", vbInformation, "Aviso"
        ValidaDatosPersVentas = False
        Exit Function
    End If
    
    If Len(Trim(FEVentas.TextMatrix(FEVentas.row, 5))) = 0 Then
        MsgBox "Falta Ingresar el periodo", vbInformation, "Aviso"
        ValidaDatosPersVentas = False
        Exit Function
    End If
    
    If Len(Trim(FEVentas.TextMatrix(FEVentas.row, 3))) = 0 Then
        MsgBox "Falta Ingresar el Monto de VA", vbInformation, "Aviso"
        ValidaDatosPersVentas = False
        Exit Function
    End If
    
    If Len(Trim(FEVentas.TextMatrix(FEVentas.row, 4))) = 0 Then
        MsgBox "Falta Ingresar la Fec. de Auditoria", vbInformation, "Aviso"
        ValidaDatosPersVentas = False
        Exit Function
    End If
    
End Function
' ***

Private Function ValidaDatosRefComercial() As Boolean
    
    ValidaDatosRefComercial = True

    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 1))) = 0 Then 'Nombre/Razon Social
        MsgBox "Falta Ingresar Nombre o Razón Social de la Referencia ", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 2))) = 0 Then 'Tipo de Referencia
        MsgBox "Falta Ingresar el Tipo de Referencia ", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 3))) = 0 Then 'Número Telefónico
        MsgBox "Falta Ingresar Número Telefónico de Referencia ", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 4))) = 0 Then 'Comentario
        MsgBox "Falta Ingresar Comentario de Referencia ", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
    If Len(Trim(feRefComercial.TextMatrix(feRefComercial.row, 5))) = 0 Then 'Direccion
        MsgBox "Falta Ingresar Direccion de Referencia ", vbInformation, "Aviso"
        ValidaDatosRefComercial = False
        Exit Function
    End If
    
End Function

Private Function ValidaDatosRefBancaria() As Boolean
    
    ValidaDatosRefBancaria = True

    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 1))) = 0 Then '
        MsgBox "Falta Seleccionar la Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 3))) = 0 Then 'Número de Cuenta
        MsgBox "Falta Ingresar el Número de Cuenta de la Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 4))) = 0 Then 'Número de Tarjeta
        MsgBox "Falta Ingresar Número de Tarjeta de Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(Trim(feRefBancaria.TextMatrix(feRefBancaria.row, 5))) = 0 Then 'Línea de Crédito
        MsgBox "Falta Ingresar Línea de Crédito de Referencia Bancaria", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(feRefBancaria.TextMatrix(feRefBancaria.row, 3)) > 20 Then
        MsgBox "Longitud de Cuenta no válida", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
    If Len(feRefBancaria.TextMatrix(feRefBancaria.row, 4)) > 20 Then
        MsgBox "Longitud de Numero de Tarjeta no válida", vbInformation, "Aviso"
        ValidaDatosRefBancaria = False
        Exit Function
    End If
    
End Function

'*** PEAC 20080412
Private Function ValidaDatosPatOtros() As Boolean
    
    ValidaDatosPatOtros = True

    If Len(Trim(fePatOtros.TextMatrix(fePatOtros.row, 1))) = 0 Then 'Descripcion
        MsgBox "Falta Ingresar la Descripcion", vbInformation, "Aviso"
        ValidaDatosPatOtros = False
        Exit Function
    End If
       
    If Len(Trim(fePatOtros.TextMatrix(fePatOtros.row, 2))) = 0 Then 'Valor Comercial
        MsgBox "Falta Ingresar el Valor Comercial.", vbInformation, "Aviso"
        ValidaDatosPatOtros = False
        Exit Function
    End If
    
End Function

'*** PEAC 20080412
Private Function ValidaDatosPatInmuebles() As Boolean

    ValidaDatosPatInmuebles = True

    If Len(Trim(fePatInmuebles.TextMatrix(fePatInmuebles.row, 1))) = 0 Then 'Ubicacion
        MsgBox "Falta Ingresar la Ubicacion", vbInformation, "Aviso"
        ValidaDatosPatInmuebles = False
        Exit Function
    End If
       
    If Len(Trim(fePatInmuebles.TextMatrix(fePatInmuebles.row, 2))) = 0 Then 'area terreno
        MsgBox "Falta Ingresar el Area de terrno.", vbInformation, "Aviso"
        ValidaDatosPatInmuebles = False
        Exit Function
    End If
       
    If Len(Trim(fePatInmuebles.TextMatrix(fePatInmuebles.row, 3))) = 0 Then 'area construida
        MsgBox "Falta Ingresar el Area Construida.", vbInformation, "Aviso"
        ValidaDatosPatInmuebles = False
        Exit Function
    End If
       
    If Len(Trim(fePatInmuebles.TextMatrix(fePatInmuebles.row, 4))) = 0 Then 'tipo uso
        MsgBox "Falta Ingresar el tipo de uso (industrial, comercial, vivienda, otros)", vbInformation, "Aviso"
        ValidaDatosPatInmuebles = False
        Exit Function
    End If
       
    If Len(Trim(fePatInmuebles.TextMatrix(fePatInmuebles.row, 5))) = 0 Then 'rrpp
        MsgBox "Falta Ingresar si el bien esta inscrito en los registros publicos", vbInformation, "Aviso"
        ValidaDatosPatInmuebles = False
        Exit Function
    End If
       
    If Len(Trim(fePatInmuebles.TextMatrix(fePatInmuebles.row, 6))) = 0 Then 'Valor Comercial
        MsgBox "Falta Ingresar el Valor Comercial.", vbInformation, "Aviso"
        ValidaDatosPatInmuebles = False
        Exit Function
    End If
    
End Function



Private Function ValidaDatosPatVehicular() As Boolean
    
    ValidaDatosPatVehicular = True

    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 1))) = 0 Then 'Marca
        MsgBox "Falta Ingresar la Marca del Vehículo", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 2))) = 0 Then 'Fecha de Fabricacion
        MsgBox "Falta Ingresar la Fecha de Fabricación", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    ElseIf CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)) > Year(gdFecSis) Then
        MsgBox "Año de Fabricación no puede ser mayor que Año actual", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 3))) = 0 Then 'Valor Comercial
        MsgBox "Falta Ingresar el Valor Comercial del Vehículo", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 4))) = 0 Then 'modelo
        MsgBox "Falta Ingresar el Modelo del vehiculo", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.row, 5))) = 0 Then 'placa
        MsgBox "Falta Ingresar la placa del vehiculo", vbInformation, "Aviso"
        ValidaDatosPatVehicular = False
        Exit Function
    End If
    
    
'    If Len(Trim(fePatVehicular.TextMatrix(fePatVehicular.Row, 4))) = 0 Then 'Condición del Patrimomio
'        MsgBox "Falta Seleccionar la Condición del Patrimonio Vehicular", vbInformation, "Aviso"
'        ValidaDatosPatVehicular = False
'        Exit Function
'    End If
    
End Function
'MADM 20101220
Private Function ValidaRestriccion(ByRef CadTmp As String) As Boolean
Dim bEnc As Boolean
Dim i As Integer
Dim ind As Integer
Dim bPersNatMagnitud As Boolean
bPersNatMagnitud = False
bEnc = False
ValidaRestriccion = False
ind = 0
CadTmp = ""
'If CadTmp <> "" Then
 If Not oPersona.Personeria = gPersonaNat Then
            For i = 1 To FEDocs.rows - 1
              
              If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
                   If (i = 1) Then
                      If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC) And oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                            bEnc = True
                      ElseIf (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC) And oPersona.ObtenerDocTipoAct(i - 1) = PersFilaEliminda Then 'JACA 20111003
                            bEnc = False
                            ind = 5
                            Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                            Exit For
                      ElseIf oPersona.ObtenerDocTipoAct(i - 1) = PersFilaEliminda Then 'JACA 20111003
                         bEnc = True
                      End If
                   Else
                        If oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                            bEnc = False
                            ind = 5
                            Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                            Exit For
                        End If
                   End If
              End If
            Next i
  Else
           
                If Trim(Right(cmbNacionalidad.Text, 12)) = "" Then
                   CadTmp = "Complete la Nacionalidad de la Persona"
                   Exit Function
                End If
                 'Persona Naturales con Nacionalidad Peruana
                If Trim(Right(cmbNacionalidad.Text, 12)) = "04028" Then
                            If oPersona.Personeria = gPersonaNat Then
                                  For i = 1 To FEDocs.rows - 1
                                    If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
                                            'DNI Y SOLO 1 REGISTRO
                                            If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdDNI) And ((FEDocs.rows - 1) = 1) Then
                                                bEnc = True
                                                Exit For
                                            End If
                                          
                                            'VALIDA 1 REGISTRO
                                            If (i = 1) Then
                                                If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdDNI) And oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                                                    bEnc = True
                                                ElseIf (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdDNI) And oPersona.ObtenerDocTipoAct(i - 1) = PersFilaEliminda Then 'JACA 20111003
                                                    bEnc = False
                                                    ind = 1
                                                     Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                                    Exit For
                                                Else
                                                    If oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                                                        bEnc = False
                                                        ind = 1
                                                        Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                          
                                            If (i <> 1) Then
                                            'JACA 20110525*************************************************************************************
                                                If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdPartNaC) Or (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdREPEV) Or (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC) And (bEnc = True) Then 'WIOR 20140703 AGREGO (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdREPEV)
                                                     If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC) And oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda And Me.cmbPersNatMagnitud.ListIndex = 1 Then
                                                            bPersNatMagnitud = True
                                                     End If
                                                     bEnc = True
                                                        
                                                 Else
                                                     If oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                                                        bEnc = False
                                                        ind = 4
                                                        Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                                        Exit For
                                                     End If
                                                End If
                                            'JACA END***********************************************************************************************
                                            End If
                                     End If
                                  Next i
                                  'JACA 20110525*************************************************************
                                  If Me.cmbPersNatMagnitud.ListIndex = 1 And bPersNatMagnitud = False And i > 1 Then
                                        bEnc = False
                                        ind = 6
                                  End If
                                  'JACA END****************************************************************
                            Else
                                  For i = 1 To FEDocs.rows - 1
                                      If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
                                          If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC Then
                                              If FEDocs.rows = 1 Then
                                                  bEnc = True
                                              Else
                                                  bEnc = True
                                              End If
                                          End If
                                      End If
                                  Next i
                            End If
                'Persona Naturales con Nacionalidad Extranjera
                Else
                     ind = 3
                     For i = 1 To FEDocs.rows - 1
                       If Trim(FEDocs.TextMatrix(i, 1)) <> "" Then
                            If (i = 1) Then
                               '20200826LARI --SE AGREGARON VARIOS TIPOS DE DOCUMENTOS ADICIONALES EXTRANJEROS SEGÚN OFICIO 13323-2020-SBS
                               If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdExtranjeria Or _
                                   CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdCIEMRE Or _
                                   CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdCPTP Or _
                                   CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdCI Or _
                                   CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdCREF Or _
                                   CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdCEPR) And oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                                   bEnc = True
                               ElseIf (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdExtranjeria) And oPersona.ObtenerDocTipoAct(i - 1) = PersFilaEliminda Then 'JACA 20111003
                                    bEnc = False
                                    Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                    Exit For
                               Else
                                    If oPersona.ObtenerDocTipoAct(i - 1) <> PersFilaEliminda Then 'JACA 20111003
                                        bEnc = False
                                        Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                        Exit For
                                    Else
                                        bEnc = True
                                    End If
                               End If
'                            Else
'                               bEnc = False
'                               ind = 4
'                               Exit For
                            End If
                            'JACA 20110525*************************************************************************************
                                If (i = 2) Then
                                   If Me.cmbPersNatMagnitud.ListIndex = 1 Then
                                        If (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) <> gPersIdRUC) Then
                                            bEnc = False
                                            ind = 6
                                            Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                            Exit For
                                        ElseIf (CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC) And oPersona.ObtenerDocTipoAct(i - 1) = PersFilaEliminda Then 'JACA 20111004
                                            bEnc = False
                                            ind = 6
                                            Call oPersona.ActualizarDocsTipoAct(PersFilaSinCambios, FEDocs.row - 1) 'JACA 20111003 para actualizar la fila a sin cambios
                                            Exit For
                                        Else
                                            bPersNatMagnitud = True
                                        End If
                                    End If
                                End If
                            'JACA END***********************************************************************************************
                       End If
                     Next i
                      'JACA 20110525*************************************************************
                            If Me.cmbPersNatMagnitud.ListIndex = 1 And bPersNatMagnitud = False And i > 1 Then 'JACA 20111004
                                  bEnc = False
                                  ind = 6
                            End If
                    'JACA END****************************************************************
                End If
  End If

    If bEnc = False Then
        If ind = 1 Then
            CadTmp = "Debe registrar como documento principal DNI, Verifique"
        ElseIf ind = 2 Then
            CadTmp = "Debe registrar solamente como documento RUC Verifique"
        ElseIf ind = 3 Then
            'CadTmp = "Debe registrar solamente como documento Carnet de Extranjería, Verifíque"
            CadTmp = "Debe registrar como documento principal Carnet de Extranjería, Verifíque"
        ElseIf ind = 4 Then
            CadTmp = "Debe registrar solamente RUC o Partida de Nacimiento como documento adicional al DNI, Verifíque"
        ElseIf ind = 5 Then
            CadTmp = "Debe registrar como único documento RUC, Verifíque"
        'ElseIf ind = 6  Then
        ElseIf ind = 6 And oPersona.Personeria <> gPersonaNat Then 'ALPA 20120523
            CadTmp = "Debe registrar el RUC para la persona con magnitud con negocio, Verifíque" 'JACA 20110525
        
        End If
    End If

 End Function
'END MADM

Private Function ValidaControles() As Boolean
Dim CadTmp As String
    
    ValidaControles = True
    
    If Me.cmdPatVehAcepta.Visible Then
        MsgBox "Falta validar los datos del patrimonio para la DDJJ.", vbInformation, "Aviso"
        cmdPatVehAcepta.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    If Me.cmdRefComAcepta.Visible Then
        MsgBox "Falta validar los datos de las referencias personales/comerciales.", vbInformation, "Aviso"
        cmdRefComAcepta.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    
    If cmdPersIDAceptar.Visible Then
        MsgBox "Pulse Aceptar para Confirmar Documento", vbInformation, "Aviso"
        cmdPersIDAceptar.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    If cmbPersPersoneria.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Personeria", vbInformation, "Aviso"
        cmbPersPersoneria.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    If oPersona.Personeria = gPersonaNat Then
    'Valida Controles de Persona Natural
    'Comentado por WIOR SEGUN OYP-RFC060-2012
'        If Len(Trim(txtPersNombreAP.Text)) = 0 Then
'            MsgBox "Falta Ingresar el Apellido Paterno", vbInformation, "Aviso"
'            If txtPersNombreAP.Enabled = True Then
'                If Trim(txtPersNombreAP.Text) <> "" Then
'                    txtPersNombreAP.SetFocus
'                End If
'            End If
'            ValidaControles = False
'            SSTabs.Tab = 0
'            Exit Function
'        End If
'        If Len(Trim(txtPersNombreAM.Text)) = 0 Then
'            MsgBox "Falta Ingresar el Apellido Materno", vbInformation, "Aviso"
'            If txtPersNombreAM.Enabled Then
'               If Trim(txtPersNombreAM.Text) <> "" Then
'                    txtPersNombreAM.SetFocus
'               End If
'            End If
'            ValidaControles = False
'            SSTabs.Tab = 0
'            Exit Function
'        End If

        'WIOR 20120705 SEGUN OYP-RFC060-2012
        If oPersona.Nacionalidad = "04028" Then
        'PARA PERUANOS
        
            'VALIDAR APELLIDO PATERNO
            If Len(Trim(txtPersNombreAP.Text)) = 0 Then
                MsgBox "Falta Ingresar el Apellido Paterno", vbInformation, "Aviso"
                If txtPersNombreAP.Enabled = True Then
                    If Trim(txtPersNombreAP.Text) <> "" Then
                        If txtPersNombreAP.Visible Then 'WIOR 20120924
                            txtPersNombreAP.SetFocus
                        End If
                    End If
                End If
                ValidaControles = False
                SSTabs.Tab = 0
                Exit Function
            End If
            
            'VALIDAR APELLIDO MATERNO Y/O DE CASADA
            If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
                If Len(Trim(txtPersNombreAM.Text)) = 0 And Len(Trim(txtApellidoCasada.Text)) = 0 Then
                    MsgBox "Falta Ingresar el Apellido Materno y/o de Casada", vbInformation, "Aviso"
                    If txtPersNombreAM.Visible Then 'WIOR 20120924
                        txtPersNombreAM.SetFocus
                    End If
                    ValidaControles = False
                    SSTabs.Tab = 0
                    Exit Function
                Else
                   If txtPersNombreAM.Enabled Or txtApellidoCasada.Enabled Then
                       If Trim(txtPersNombreAM.Text) <> "" Then
                            If txtPersNombreAM.Visible Then 'WIOR 20120924
                                txtPersNombreAM.SetFocus
                            End If
                       ElseIf Trim(txtApellidoCasada.Text) <> "" Then
                            If txtApellidoCasada.Visible Then 'WIOR 20120924
                                txtApellidoCasada.SetFocus
                            End If
                       End If
                    End If
                End If
            Else
                If Len(Trim(txtPersNombreAM.Text)) = 0 Then
                    MsgBox "Falta Ingresar el Apellido Materno", vbInformation, "Aviso"
                    If txtPersNombreAM.Enabled Then
                       If Trim(txtPersNombreAM.Text) <> "" Then
                            If txtPersNombreAM.Visible Then 'WIOR 20120924
                                txtPersNombreAM.SetFocus
                            End If
                       End If
                    End If
                    ValidaControles = False
                    SSTabs.Tab = 0
                    Exit Function
                End If
            End If
        Else
        'PARA EXTRANJEROS
            If oPersona.Sexo = "M" Then
                If Len(Trim(txtPersNombreAP.Text)) = 0 And Len(Trim(txtPersNombreAM.Text)) = 0 Then
                        MsgBox "Falta Ingresar el Apellido Paterno y/o Materno", vbInformation, "Aviso"
                        If txtPersNombreAP.Visible Then 'WIOR 20120924
                            txtPersNombreAP.SetFocus
                        End If
                        ValidaControles = False
                        SSTabs.Tab = 0
                        Exit Function
                Else
                    If txtPersNombreAP.Enabled Or txtPersNombreAM.Enabled Then
                       If Trim(txtPersNombreAP.Text) <> "" Then
                            If txtPersNombreAP.Visible Then 'WIOR 20120924
                                txtPersNombreAP.SetFocus
                            End If
                       ElseIf Trim(txtPersNombreAM.Text) <> "" Then
                            If txtPersNombreAM.Visible Then 'WIOR 20120924
                                txtPersNombreAM.SetFocus
                            End If
                       End If
                    End If
                End If
            ElseIf oPersona.Sexo = "F" Then
                If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) Then
                    If Len(Trim(txtPersNombreAP.Text)) = 0 And Len(Trim(txtPersNombreAM.Text)) = 0 And Len(Trim(txtApellidoCasada.Text)) = 0 Then
                        MsgBox "Falta Ingresar el Apellido Paterno, Materno y/o de Casada.", vbInformation, "Aviso"
                        If txtPersNombreAP.Visible Then 'WIOR 20120924
                            txtPersNombreAP.SetFocus
                        End If
                        ValidaControles = False
                        SSTabs.Tab = 0
                        Exit Function
                    Else
                        If txtPersNombreAP.Enabled Or txtPersNombreAM.Enabled Or txtApellidoCasada.Enabled Then
                           If Trim(txtPersNombreAP.Text) <> "" Then
                                If txtPersNombreAP.Visible Then 'WIOR 20120924
                                    txtPersNombreAP.SetFocus
                                End If
                           ElseIf Trim(txtPersNombreAM.Text) <> "" Then
                                If txtPersNombreAM.Visible Then 'WIOR 20120924
                                    txtPersNombreAM.SetFocus
                                End If
                           ElseIf Trim(txtApellidoCasada.Text) <> "" Then
                                If txtApellidoCasada.Visible Then 'WIOR 20120924
                                    txtApellidoCasada.SetFocus
                                End If
                           End If
                        End If
                    End If
                Else
                    If Len(Trim(txtPersNombreAP.Text)) = 0 And Len(Trim(txtPersNombreAM.Text)) = 0 Then
                        MsgBox "Falta Ingresar el Apellido Paterno y/o Materno.", vbInformation, "Aviso"
                        If txtPersNombreAP.Visible Then 'WIOR 20120924
                            txtPersNombreAP.SetFocus
                        End If
                        ValidaControles = False
                        SSTabs.Tab = 0
                        Exit Function
                    Else
                        If txtPersNombreAP.Enabled Or txtPersNombreAM.Enabled Then
                           If Trim(txtPersNombreAP.Text) <> "" Then
                                If txtPersNombreAP.Visible Then 'WIOR 20120924
                                    txtPersNombreAP.SetFocus
                                End If
                           ElseIf Trim(txtPersNombreAM.Text) <> "" Then
                                If txtPersNombreAM.Visible Then 'WIOR 20120924
                                    txtPersNombreAM.SetFocus
                                End If
                           End If
                        End If
                    End If
                End If
            End If
        End If
        'WIOR FIN
        If Len(Trim(txtPersNombreN.Text)) = 0 Then
            MsgBox "Falta Ingresar Nombres de la Persona", vbInformation, "Aviso"
            'txtPersNombreN.SetFocus
            If txtPersNombreN.Enabled Then 'EJVG20120120
                txtPersNombreN.SetFocus
            End If
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        
        '*** PEAC 20080412
'        If Len(Trim(txtPersTelefono.Text)) = 0 Then
'            MsgBox "Falta Ingresar Teléfono", vbInformation, "Aviso"
'            txtPersTelefono.SetFocus
'            ValidaControles = False
'            SSTabs.Tab = 0
'            Exit Function
'        End If
        
        If cmbPersNatSexo.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Sexo de la Persona", vbInformation, "Aviso"
            'cmbPersNatSexo.SetFocus
            If cmbPersNatSexo.Enabled Then 'EJVG20120120
                cmbPersNatSexo.SetFocus
            End If
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If cmbPersNatEstCiv.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Estado Civil de la Persona", vbInformation, "Aviso"
            'cmbPersNatEstCiv.SetFocus
            If cmbPersNatEstCiv.Enabled Then 'EJVG20120120
                cmbPersNatEstCiv.SetFocus
            End If
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        If Len(Trim(txtPersNatHijos.Text)) = 0 Then
            MsgBox "Falta Ingresar el Numero de Hijos", vbInformation, "Aviso"
            txtPersNatHijos.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If

'        If Len(Trim(TxtTipoSangre.Text)) = 0 Then
'            MsgBox "Falta Ingresar el Tipo de Sangre de la Persona de la Persona", vbInformation, "Aviso"
'            ValidaControles = False
'            Exit Function
'        End If

         '*** madm 20091112
        If val(Trim(txtIngresoProm.Text)) = 0 Then
            MsgBox "Falta Registrar Ingreso Promedio", vbInformation, "Aviso"
            txtIngresoProm.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        
       If nTipoForm <> 1 Then 'add pti1 ers070-2018 26/12/2018
            bSensible = False
            If val(Trim(nIngPromedio)) <> val(Trim(txtIngresoProm.Text)) And BotonNuevo = False Then   'add pti1 ers070-2018 26/12/2018
                If MsgBox("¿Usted está seguro de actualizar los datos sensibles?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                        txtIngresoProm.SetFocus
                        ValidaControles = False
                        SSTabs.Tab = 0
                        Exit Function
                End If
                nIngPromedio = ""
                bSensible = True 'add pti1 ers070-2018 26/12/2018
            End If 'add pti1 ers070-2018 26/12/2018
        End If
        
        'madm 20100309
        If Len(Trim(txtActiGiro.Text)) = 0 Then
            MsgBox "Falta Ingresar Actividad Principal", vbInformation, "Aviso"
            txtActiGiro.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        'end madm
        If Me.chkcred.value = 0 And Me.chkaho.value = 0 And Me.chkser.value = 0 And Me.chkotro.value = 0 Then
            MsgBox "Falta Seleccionar Finalidad Relación CMAC", vbInformation, "Aviso"
            chkcred.SetFocus
            ValidaControles = False
            SSTabs.Tab = 0
            Exit Function
        End If
        'end madm

    'madm 20100503
     If Me.cboocupa.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Ocupacion de la Persona", vbInformation, "Aviso"
        cboocupa.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    If Trim(Right(Me.cboocupa.Text, 10)) = "116" And txtActiGiro.Text = "" Then
             MsgBox "Debe describir la actividad o giro", vbInformation, "Mensaje"
             ValidaControles = False
             Me.txtActiGiro.SetFocus
            SSTDatosGen.Tab = 0
            Exit Function
    End If

    'end madm
    '** Juez 20120326 ************************************
    If Me.cboCargos.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Cargo de la Persona", vbInformation, "Aviso"
        cboCargos.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    '** End Juez ******************************************
    'JACA 20110426*****************************************************************
        If cmbPersNatMagnitud.ListIndex = -1 Then
                MsgBox "Falta Seleccionar la Magnitud de la Persona Natural", vbInformation, "Aviso"
                cmbPersNatMagnitud.SetFocus
                ValidaControles = False
                SSTabs.Tab = 0
                Exit Function
        End If
    'JACA END**********************************************************************
    '** Juez 20120327 ************************************
    If Me.cboResidente.ListIndex = 1 And Me.cboPaisReside.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Pais de Residencia", vbInformation, "Aviso"
            If cboPaisReside.Enabled Then cboPaisReside.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
    End If
    '** End Juez ******************************************
    'EJVG20120813 ***
    If Me.cboPEPS.Visible = True And Me.cboPEPS.Enabled = True And Me.cboPEPS.ListIndex = -1 Then
        MsgBox "Falta seleccionar el PEPS de la Persona", vbInformation, "Aviso"
        Me.cboPEPS.SetFocus
        ValidaControles = False
        Exit Function
    End If
    'END EJVG *******
    If Me.CboAutoriazaUsoDatos.Visible = True And Me.CboAutoriazaUsoDatos.Enabled = True And Me.CboAutoriazaUsoDatos.ListIndex = -1 And nTipoForm <> 1 Then 'ADD por pti1 ers070-2018 se agrego  And nTipoForm <> 1
            MsgBox "Falta seleccionar la autorización de datos", vbInformation, "Aviso"
            Me.CboAutoriazaUsoDatos.SetFocus
            ValidaControles = False
        Exit Function
    End If
        
    Else
    'Valida Controles de Persona Juridica
        If Len(Trim(txtPersNombreRS.Text)) = 0 Then
            MsgBox "Falta Ingresar la razon Social de la Persona", vbInformation, "Aviso"
            txtPersNombreRS.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
    
        If Len(Trim(TxtSiglas.Text)) = 0 Then
            MsgBox "Falta Ingresar la Siglas de la Persona", vbInformation, "Aviso"
            TxtSiglas.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If cmbPersJurTpo.ListIndex = -1 Then
            MsgBox "Falta Seleccionar el Tipo de Persona Juridica", vbInformation, "Aviso"
            cmbPersJurTpo.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        
        If gsProyectoActual = "H" Then
            If cmbPersJurMagnitud.ListIndex = -1 Then
                MsgBox "Falta Seleccionar la Magnitud Empresarial de la Persona", vbInformation, "Aviso"
                cmbPersJurMagnitud.SetFocus
                ValidaControles = False
                SSTabs.Tab = 1
                Exit Function
            End If
        End If
        
        If Len(Trim(txtPersJurEmpleados.Text)) = 0 Then
            MsgBox "Falta Ingresar el Numero de Empleados ", vbInformation, "Aviso"
            txtPersJurEmpleados.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        '** Juez 20120328 *********************************************************
        If Len(Trim(Me.txtPersJurObjSocial.Text)) = 0 Then
            MsgBox "Falta Ingresar el Objeto Social ", vbInformation, "Aviso"
            txtPersJurObjSocial.SetFocus
            ValidaControles = False
            SSTabs.Tab = 1
            Exit Function
        End If
        '** End Juez **************************************************************
    End If
    
    'Valida Datos Generales
    CadTmp = ValidaFecha(txtPersNacCreac.Text)
    If Len(CadTmp) > 0 Then
        MsgBox CadTmp, vbInformation, "Aviso"
        'txtPersNacCreac.SetFocus
        If txtPersNacCreac.Enabled Then
            txtPersNacCreac.SetFocus
        End If
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    'Valida Datos Generales Fecha Insc Ruc
    CadTmp = ValidaFecha(txtPersFecInscRuc.Text)
'    If Len(CadTmp) > 0 Then
'        MsgBox CadTmp, vbInformation, "Aviso"
'        txtPersFecInscRuc.SetFocus
'        ValidaControles = False
'        SSTDatosGen.Tab = 0
'        Exit Function
'    End If
    
    If Len(CadTmp) > 0 Then
        'MsgBox CadTmp, vbInformation, "Aviso"
        txtPersFecInscRuc.Text = "01/01/1900"
'        ValidaControles = False
'        SSTDatosGen.Tab = 0
'        Exit Function
    End If
    
    
    'Valida Datos Generales Fecha Inicio actividad
    CadTmp = ValidaFecha(txtPersFecIniActi.Text)
'    If Len(CadTmp) > 0 Then
'        MsgBox CadTmp, vbInformation, "Aviso"
'        txtPersFecIniActi.SetFocus
'        ValidaControles = False
'        SSTDatosGen.Tab = 0
'        Exit Function
'    End If

    If Len(CadTmp) > 0 Then
'       MsgBox CadTmp, vbInformation, "Aviso"
        txtPersFecIniActi.Text = "01/01/1900"
'        ValidaControles = False
'        SSTDatosGen.Tab = 0
'        Exit Function
    End If
    
    
    
    If CboPersCiiu.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Ciiu de la Persona", vbInformation, "Aviso"
        CboPersCiiu.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
'*** PEAC 20080412
    If cboTipoComp.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Tipo de competencia.", vbInformation, "Aviso"
        cboTipoComp.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    If cboTipoSistInfor.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Tipo de sistema de información.", vbInformation, "Aviso"
        cboTipoSistInfor.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    If cboCadenaProd.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la cadena productiva", vbInformation, "Aviso"
        cboCadenaProd.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
'*** FIN PEAC


    
    If cmbPersEstado.ListIndex = -1 Then
        MsgBox "Falta Seleccionar el Estado de la Persona", vbInformation, "Aviso"
        cmbPersEstado.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    If cmbPersUbiGeo(4).ListIndex = -1 Then
        'MsgBox "Falta Seleccionar la Ubicacion Geografica de la Persona", vbInformation, "Aviso"
        MsgBox "Falta Seleccionar la Ubicacion Geografica del Domicilio de la Persona", vbInformation, "Aviso"
        cmbPersUbiGeo(4).SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 1
        Exit Function
    End If
    
    If Len(Trim(txtPersDireccDomicilio.Text)) = 0 Then
        MsgBox "Falta Ingresar el Domicilio de la Persona", vbInformation, "Aviso"
        txtPersDireccDomicilio.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 1
        Exit Function
    End If
    
    If cmbPersDireccCondicion.ListIndex = -1 Then
        MsgBox "Falta Seleccionar la Condicion del Domicilio de la Persona", vbInformation, "Aviso"
        cmbPersDireccCondicion.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 1
        Exit Function
    End If
    
    If Trim(FEDocs.TextMatrix(1, 0)) = "" And Trim(FEDocs.TextMatrix(1, 1)) = "" Then
        MsgBox "Falta Ingresar algun Documento de Identidad", vbInformation, "Aviso"
        SSTIdent.SetFocus
        ValidaControles = False
        Exit Function
    End If
    
    If Trim(txtPersNatHijos.Text) = "" Then txtPersNatHijos.Text = "0"
    '*** PEAC 20080412
    If Trim(txtPersNatNumEmp.Text) = "" Then txtPersNatNumEmp.Text = "0"
    
    If CInt(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
        If CInt(txtPersNatHijos.Text) > 15 Then
            MsgBox "Numero de Hijos Incorrecto", vbInformation, "Aviso"
            txtPersNatHijos.SetFocus
            ValidaControles = False
            Exit Function
        End If
    End If
    
    If CInt(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
        If cmbNacionalidad.ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Nacionalidad de la Persona", vbInformation, "Aviso"
            If cmbNacionalidad.Enabled Then cmbNacionalidad.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    
        If CmbRela.ListIndex = -1 Then
            MsgBox "Falta Seleccionar la Relacion de la Persona con la Institución", vbInformation, "Aviso"
            CmbRela.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
        
    End If
    'ORCR20140318 INICIO ***
    If Not (validarTelefonoControl(txtPersTelefono)) Then
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
    If Not (validarTelefonoControl(txtPersTelefono2)) Then
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
     If Not (validarTelefonoControl(txtCel1)) Then
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
     If Not (validarTelefonoControl(txtCel2)) Then
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
     If Not (validarTelefonoControl(txtCel3)) Then
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    
  
    If Len(Trim(txtPersTelefono.Text)) = 0 And Len(Trim(txtPersTelefono2.Text)) = 0 And Len(Trim(txtCel1.Text)) = 0 And Len(Trim(txtCel2.Text)) = 0 And Len(Trim(txtCel3.Text)) = 0 Then
        MsgBox "Debe Ingresar un Teléfono Fijo o Celular", vbInformation, "Aviso"
        txtPersTelefono.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    'EJVG20111207 *********************Obliga un Tel. Fijo o Celular, verifica los Emails
    'If Len(Trim(txtPersTelefono.Text)) = 0 And Len(Trim(txtPersTelefono2.Text)) = 0 And Len(Trim(txtCel1.Text)) = 0 And Len(Trim(txtCel2.Text)) = 0 And Len(Trim(txtCel3.Text)) = 0 Then
    '        MsgBox "Falta Ingresar un Teléfono Fijo o Celular", vbInformation, "Aviso"
    '        txtPersTelefono.SetFocus
    '        ValidaControles = False
    '        SSTDatosGen.Tab = 0
    '        Exit Function
    '    End If
    
    'ORCR20140318 FIN *******
    If TxtEmail.Text <> "" Then
        If EsEmailValido(TxtEmail.Text) = False Then
            MsgBox "Ingrese un Email válido", vbInformation, "Aviso"
            TxtEmail.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    End If
    If TxtEmail2.Text <> "" Then
        If EsEmailValido(TxtEmail2.Text) = False Then
            MsgBox "Ingrese un Email válido", vbInformation, "Aviso"
            TxtEmail2.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
    End If
    'EJVG20120724 *** Sujeto Obligado DJ
    If Me.cboSujetoObligado.ListIndex = -1 Then
        MsgBox "Falta seleccionar si la Persona es un Sujeto Obligado a DJ", vbInformation, "Aviso"
        If Me.cboSujetoObligado.Enabled Then Me.cboSujetoObligado.SetFocus
        ValidaControles = False
        Exit Function
    Else
        If CInt(Right(Me.cboSujetoObligado.Text, 1)) = 1 And Me.cboOfCumplimiento.ListIndex = -1 Then
            MsgBox "Falta seleccionar si cuenta con Oficial de Cumplimiento", vbInformation, "Aviso"
            If Me.cboOfCumplimiento.Enabled Then Me.cboOfCumplimiento.SetFocus
            ValidaControles = False
            Exit Function
        End If
    End If
    'END EJVG *******
    'JUEZ 20131007 *****************************************************************
    If Trim(cboRemInfoEmail.Text) = "" Then
        MsgBox "Falta seleccionar si desea Remisión de Información por Email", vbInformation, "Aviso"
        cboRemInfoEmail.SetFocus
        ValidaControles = False
        SSTDatosGen.Tab = 0
        Exit Function
    End If
    If Trim(Right(cboRemInfoEmail.Text, 2)) = "1" Then
        If Trim(TxtEmail.Text) = "" Then
            MsgBox "Falta ingresar el email de la Persona ", vbInformation, "Aviso"
            TxtEmail.SetFocus
            ValidaControles = False
            Exit Function
        End If
    End If
    If Trim(Right(cboMotivoActu.Text, 2)) = "2" Then
        If Not ValidaDetallesCampDatos Then
            ValidaControles = False
            Exit Function
        End If
        If gsCodUser = txtBUsuario.Text Then
            MsgBox "El usuario responsable de la información no puede ser el mismo que registra los datos", vbInformation, "Aviso"
            Screen.MousePointer = 0
            txtBUsuario.SetFocus
            ValidaControles = False
            Exit Function
        End If
        If Trim(txtActiGiro.Text) = "" Then
            MsgBox "Falta ingresar la Actividad/Giro de la Persona", vbInformation, "Aviso"
            txtActiGiro.SetFocus
            ValidaControles = False
            SSTDatosGen.Tab = 0
            Exit Function
        End If
        If Trim(txtBUsuario.Text) = "" Or Len(Trim(txtBUsuario.Text)) <> 4 Then
            MsgBox "Falta seleccionar el usuario responsable de la información", vbInformation, "Aviso"
            txtBUsuario.SetFocus
            ValidaControles = False
            Exit Function
        End If
    End If
    'END JUEZ **********************************************************************
End Function

Private Sub HabilitaControlesPersona(ByVal pbBloqueo As Boolean)
    
    cmbPersPersoneria.Enabled = pbBloqueo
    
    If Not oPersona Is Nothing Then
        If oPersona.Personeria = gPersonaNat Then
            'Ficha de Persona natural
            Call HabilitaFichaPersonaNat(pbBloqueo)
        Else
            'Ficha de Persona Juridica
            Call HabilitaFichaPersonaJur(pbBloqueo)
        End If
    Else
        'Ficha de Persona natural
        Call HabilitaFichaPersonaNat(pbBloqueo)
        'Ficha de Persona Juridica
        Call HabilitaFichaPersonaJur(Not pbBloqueo)
    End If
            
    TxtBCodPers.Enabled = Not pbBloqueo
    'Ficha de Relaciones de Persona
    FERelPers.lbEditarFlex = pbBloqueo
    cmdPersRelacNew.Enabled = pbBloqueo
    'APRI 20170630 TI-ERS025
    If pbBloqueo Then
        cmdPersRelacEditar.Enabled = bHabilitarBoton 'pbBloqueo
        cmdPersRelacDel.Enabled = bHabilitarBoton 'pbBloqueo
    Else
        cmdPersRelacEditar.Enabled = pbBloqueo
        cmdPersRelacDel.Enabled = pbBloqueo
    End If
    'END APRI
    cmdPersRelacAceptar.Enabled = pbBloqueo
    cmdPersRelacCancelar.Enabled = pbBloqueo
    
    'madm 2010327 ficha visitas
    FEVisitas.lbEditarFlex = pbBloqueo
    cmdVisitasNuevo.Enabled = pbBloqueo
    cmdVisitasEditar.Enabled = pbBloqueo
    cmdVisitasEliminar.Enabled = pbBloqueo
    cmdVisitasAceptar.Enabled = pbBloqueo
    cmdVisitasCancelar.Enabled = pbBloqueo
    'end madm
    
    'Ficha RefComercial
    feRefComercial.lbEditarFlex = pbBloqueo
    cmdRefComNuevo.Enabled = pbBloqueo
    cmdRefComEdita.Enabled = pbBloqueo
    cmdRefComElimina.Enabled = pbBloqueo
    cmdRefComAcepta.Enabled = pbBloqueo
    cmdRefComCancela.Enabled = pbBloqueo
    
    'Ficha Referencia Bancaria
    feRefBancaria.lbEditarFlex = pbBloqueo
    cmdRefBanNuevo.Enabled = pbBloqueo
    cmdRefBanEdita.Enabled = pbBloqueo
    cmdRefBanElimina.Enabled = pbBloqueo
    cmdRefBanAcepta.Enabled = pbBloqueo
    cmdRefBanCancela.Enabled = pbBloqueo
    
    'Ficha Patrimonio Vehicular
    fePatVehicular.lbEditarFlex = pbBloqueo
    '*** PEAC 20080412
    fePatOtros.lbEditarFlex = pbBloqueo
    fePatInmuebles.lbEditarFlex = pbBloqueo
    
    cmdPatVehNuevo.Enabled = pbBloqueo
    cmdPatVehEdita.Enabled = pbBloqueo
    cmdPatVehElimina.Enabled = pbBloqueo
    cmdPatVehAcepta.Enabled = pbBloqueo
    cmdPatVehCancela.Enabled = pbBloqueo
    
    'Ficha de Fuentes de Ingreso
    Call HabilitaControlesPersonaFtesIngreso(pbBloqueo)
    'Firma
    'CmdActFirma.Enabled = pbBloqueo
    cmdActualizarFirma.Enabled = pbBloqueo 'ARCV 25-10-2006
        
    'Ficha de Datos Generales
    txtPersNacCreac.Enabled = pbBloqueo
    ' *** PEAC
    txtPersFecInscRuc.Enabled = pbBloqueo
    txtPersFecIniActi.Enabled = pbBloqueo
    ' ***
    txtPersTelefono.Enabled = pbBloqueo
    txtPersTelefono2.Enabled = pbBloqueo
    TxtEmail.Enabled = pbBloqueo
    
    'EJVG20111207**********************
    txtCel1.Enabled = pbBloqueo
    txtCel2.Enabled = pbBloqueo
    txtCel3.Enabled = pbBloqueo
    TxtEmail2.Enabled = pbBloqueo
    
    '*** PEAC 20080412
    txtNumDependi.Enabled = pbBloqueo
    txtActComple.Enabled = pbBloqueo
    txtNumPtosVta.Enabled = pbBloqueo
    txtActiGiro.Enabled = pbBloqueo
    '*** FIN PEAC
    
    CboPersCiiu.Enabled = pbBloqueo
    '*** PEAC 20080412
        cboTipoComp.Enabled = pbBloqueo
        cboTipoSistInfor.Enabled = pbBloqueo
        cboCadenaProd.Enabled = pbBloqueo
        
        cboMonePatri.Enabled = pbBloqueo
    '*** FIN PEAC
    cmbPersEstado.Enabled = pbBloqueo
    
    'Ficha de Ubicacion Geografica
    cmbPersUbiGeo(0).Enabled = pbBloqueo
    cmbPersUbiGeo(1).Enabled = pbBloqueo
    cmbPersUbiGeo(2).Enabled = pbBloqueo
    cmbPersUbiGeo(3).Enabled = pbBloqueo
    cmbPersUbiGeo(4).Enabled = pbBloqueo
    txtPersDireccDomicilio.Enabled = pbBloqueo
    
    '*** PEAC 20080801
    txtRefDomicilio.Enabled = pbBloqueo
    
    'JUEZ 20131007 Ficha UbiGeo Negocio/Centro Laboral ***********
    cmbNegUbiGeo(0).Enabled = pbBloqueo
    cmbNegUbiGeo(1).Enabled = pbBloqueo
    cmbNegUbiGeo(2).Enabled = pbBloqueo
    cmbNegUbiGeo(3).Enabled = pbBloqueo
    cmbNegUbiGeo(4).Enabled = pbBloqueo
    txtNegDireccion.Enabled = pbBloqueo
    txtRefNegocio.Enabled = pbBloqueo
    txtNombreCentroLaboral.Enabled = pbBloqueo
    cboRemInfoEmail.Enabled = pbBloqueo
    txtBUsuario.Enabled = pbBloqueo
    'END JUEZ ****************************************************
    
    cmbPersDireccCondicion.Enabled = pbBloqueo
    txtValComercial.Enabled = pbBloqueo
    TxtSbs.Enabled = pbBloqueo
    CmbRela.Enabled = pbBloqueo
    txtIngresoProm.Enabled = pbBloqueo
    'MADM 20091113
    'txtcargo.Enabled = pbBloqueo
    'txtcentro.Enabled = pbBloqueo
    chkcred.Enabled = pbBloqueo
    chkaho.Enabled = pbBloqueo
    chkser.Enabled = pbBloqueo
    chkotro.Enabled = pbBloqueo
    
    cboocupa.Enabled = pbBloqueo 'madm 20100503
    'EJVG20120813 ***
    'chkpeps.Enabled = pbBloqueo 'madm 20100408
    Me.cboPEPS.Enabled = pbBloqueo
    Me.cboSujetoObligado.Enabled = pbBloqueo
    Me.cboOfCumplimiento.Enabled = pbBloqueo
    'END EJVG *******
    'end madm
    
    cboCargos.Enabled = pbBloqueo '** Juez 20120326
    
    cboMotivoActu.Enabled = pbBloqueo 'JUEZ 20131029
    
    'Ficha de Identificacion
    SSTIdent.Enabled = pbBloqueo
    FEDocs.lbEditarFlex = IIf(pbBloqueo, False, True)
    cmdPersIDnew.Enabled = pbBloqueo
    cmdPersIDedit.Enabled = pbBloqueo
    cmdPersIDDel.Enabled = pbBloqueo
    CmdPersAceptar.Enabled = True
    TxtCodCIIU.Enabled = pbBloqueo
    
End Sub
Private Sub HabilitaControlesPersonaFtesIngreso(ByVal pbBloqueo As Boolean)
    FEFteIng.lbEditarFlex = pbBloqueo
    CmdFteIngNuevo.Enabled = pbBloqueo
    CmdFteIngEditar.Enabled = pbBloqueo
    CmdFteIngEliminar.Enabled = pbBloqueo
    CmdPersFteConsultar.Enabled = pbBloqueo
End Sub
Private Sub HabilitaFichaPersonaNat(ByVal pbFicActiva As Boolean)

    lblPersNombreAP.Enabled = pbFicActiva
    txtPersNombreAP.Enabled = pbFicActiva
    lblPersNombreAM.Enabled = pbFicActiva
    txtPersNombreAM.Enabled = pbFicActiva
    lblApCasada.Enabled = pbFicActiva
    txtApellidoCasada.Enabled = pbFicActiva
    lblPersNombreN.Enabled = pbFicActiva
    txtPersNombreN.Enabled = pbFicActiva
    lblPersNatSexo.Enabled = pbFicActiva
    cmbPersNatSexo.Enabled = pbFicActiva
    cboMonePatri.Enabled = pbFicActiva '*** PEAC 20080412
    lblPersNatEstCiv.Enabled = pbFicActiva
    cmbPersNatEstCiv.Enabled = pbFicActiva
    lblPersNatHijos.Enabled = pbFicActiva
    txtPersNatHijos.Enabled = pbFicActiva
    '*** PEAC 20080412
    txtPersNatNumEmp.Enabled = pbFicActiva
    
    lblPeso.Enabled = pbFicActiva
    TxtPeso.Enabled = pbFicActiva
    LblTalla.Enabled = pbFicActiva
    TxtTalla.Enabled = pbFicActiva
    LblTpoSangre.Enabled = pbFicActiva
    CboTipoSangre.Enabled = pbFicActiva
    cmbNacionalidad.Enabled = pbFicActiva
    '*** PEAC 20080801
    'Me.cboMotivoActu.Enabled = pbFicActiva 'Comentado por JUEZ 20131029
    
    '*** CTI3
    cmdPerNatDatAdc.Enabled = pbFicActiva
    
    'chkResidente.Enabled = pbFicActiva '** Comentado por Juez 20120327
    lblNacionalidad.Enabled = pbFicActiva
    
    '** Juez 20120327 ************************************
    Me.lblResidente.Enabled = pbFicActiva
    Me.cboResidente.Enabled = pbFicActiva
    If pbFicActiva = True And Me.cboResidente.ListIndex = 1 Then
        Me.cboPaisReside.Enabled = pbFicActiva
    Else
        Me.cboPaisReside.Enabled = False
    End If
    '** End Juez *****************************************
    
     'MAVM 03042009
    txtPersFallec.Enabled = pbFicActiva
    'End MAVM
    
    'MADM 20091113
    'txtcargo.Enabled = pbFicActiva
    'txtcentro.Enabled = pbFicActiva
    
    'MARG 03052016: habilitacion rfc1603230001
    Me.chkaho.Enabled = True
    Me.chkcred.Enabled = True
    Me.chkser.Enabled = True
    Me.chkotro.Enabled = True
    
    Me.cboocupa.Enabled = pbFicActiva 'MADM 20100503
    'Me.chkpeps.Enabled = pbFicActiva 'MADM 20100408
    Me.cboPEPS.Enabled = pbFicActiva 'EJVG20120813
    'END MADM
    
    cboCargos.Enabled = pbFicActiva '** Juez 20120326
    
    Me.cmbPersNatMagnitud.Enabled = pbFicActiva 'JACA 20110427
    'cboAutoriazaUsoDatos.Enabled = pbFicActiva 'FRHU 20151130 ERS077-2015 'comentado por pti1 ers070-2018 17/12/2018
    
    Call ValidaAutorizardatos 'add pti1 ers070-2018
    If bPemisoAD Then 'ADD PTI1 ERS70-2018 26/12/2018
        If nTipoForm = 1 Then
        CboAutoriazaUsoDatos.Enabled = True
        Else
         CboAutoriazaUsoDatos.Enabled = False
        End If
    End If 'FIN PTI1
End Sub
Private Sub HabilitaFichaPersonaJur(ByVal pbFicActiva As Boolean)
    txtPersNombreRS.Enabled = pbFicActiva
    TxtSiglas.Enabled = pbFicActiva
    cmbPersJurTpo.Enabled = pbFicActiva
    txtPersJurEmpleados.Enabled = pbFicActiva
    cmbPersJurMagnitud.Enabled = pbFicActiva
    lblMagnitudEmpresarial.Enabled = pbFicActiva
    txtPersJurObjSocial.Enabled = pbFicActiva
    lblPersNombre.Enabled = pbFicActiva
    lblPersJurSiglas.Enabled = pbFicActiva
    lblPersJurTpo.Enabled = pbFicActiva
    lblPersJurEmpleados.Enabled = pbFicActiva
    lblPersJurMagnitud.Enabled = pbFicActiva
    lblPersJurObjSocial.Enabled = pbFicActiva
    FEVentas.lbEditarFlex = pbFicActiva
    cmdVentasNuevo.Enabled = pbFicActiva
    cmdVentasEditar.Enabled = pbFicActiva
    cmdVentasEliminar.Enabled = pbFicActiva
    cmdVentasAceptar.Enabled = pbFicActiva
    cmdVentasCancelar.Enabled = pbFicActiva
    '*** CTI3
    cmdAdicJuridica.Enabled = pbFicActiva
    
    '***
    'MARG 03052016: habilitacion rfc1603230001
    Me.chkaho.Enabled = True
    Me.chkcred.Enabled = True
    Me.chkser.Enabled = True
    Me.chkotro.Enabled = True
End Sub

Private Sub CargaControles()
'Dim sSql As String
'Dim Conn As DConecta
'Dim R As ADODB.Recordset
'Dim i As Integer
'Dim oConstante As DConstante
'Dim oCtasIF As NCajaCtaIF
'Dim sSql As String
'Dim Conn As COMConecta.DCOMConecta
'Dim R As ADODB.Recordset
'Dim i As Integer
'Dim oConstante As COMDConstantes.DCOMConstantes
'Dim oCtasIF As COMNCajaGeneral.NCOMCajaCtaIF
'Dim oPersonas As COMDPersona.DCOMPersonas

On Error GoTo ERRORCargaControles
    bEstadoCargando = True
    cmdPersDocEjecutado = 0
    cmdPersRelaEjecutado = 0
    cmdPersVisitasEjecutado = 0 'madm 20100293
    cmdPersFteIngresoEjecutado = 0
    NomMoverSSTabs = -1
    FERelPersNoMoverdeFila = -1
    FEDocsPersNoMoverdeFila = -1
    FEFtePersNoMoverdeFila = -1
    FEVisitasPersNoMoverdeFila = -1 'madm 20100293
    FEVentasPersNoMoverdeFila = -1 'MAVM 20100607
    
'    Set oConstante = New COMDConstantes.DCOMConstantes 'DConstante
'    Set Conn = New COMConecta.DCOMConecta 'DConecta
'    Conn.AbreConexion
    
    txtPersNacCreac.Text = Format(gdFecSis, "dd/mm/yyyy")
    '*** PEAC 20080412
    txtPersFecInscRuc.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtPersFecIniActi.Text = Format(gdFecSis, "dd/mm/yyyy")
    
    Call Cargar_Datos_Objetos_Persona
    
    'Carga Combo de Flex de Relaciones de Persona
    'FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
'    FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
       
    'Carga Combo de Documentos
'    FEDocs.CargaCombo oConstante.RecuperaConstantes(gPersIdTipo)
    
    'Carga Combo de Tipos de Referencia Comercial
'    feRefComercial.CargaCombo oConstante.RecuperaConstantes(3028)
    
    'Carga Combo de Condicion de Patrimonio Vehicular
'    fePatVehicular.CargaCombo oConstante.RecuperaConstantes(3029)
    
'    Set oConstante = Nothing
    
    'Carga TextBuscar Instituciones Financieras
'    Set oCtasIF = New COMNCajaGeneral.NCOMCajaCtaIF  'NCajaCtaIF
'    feRefBancaria.psRaiz = "BANCOS"
'    feRefBancaria.rsTextBuscar = oCtasIF.GetInstFinancieras("0[123]")
'    Set oCtasIF = Nothing
    
'    'Carga Tipos de Sangre
'    Call CargaComboConstante(gPersTpoSangre, CboTipoSangre)
'
'    'Carga Condiciones de Domicilio
'    cmbPersDireccCondicion.Clear
'    Call CargaComboConstante(gPersCondDomic, cmbPersDireccCondicion)
'
'    'Carga Personeria
'    cmbPersPersoneria.Clear
'    Call CargaComboConstante(gPersPersoneria, cmbPersPersoneria)
'
'    'Carga Tipos de Sexo de Personas
'    cmbPersNatSexo.AddItem "FEMENINO" & Space(50) & "F"
'    cmbPersNatSexo.AddItem "MASCULINO" & Space(50) & "M"
'
'    'Carga Magnitud Empresarial
'    cmbPersJurMagnitud.Clear
'    Call CargaComboConstante(gPersJurMagnitud, cmbPersJurMagnitud)
'
'    'Carga Condicion de Domicilio
'    cmbPersDireccCondicion.Clear
'    Call CargaComboConstante(gPersCondDomic, cmbPersDireccCondicion)
'
'    'Carga Estado Civil
'    cmbPersNatEstCiv.Clear
'    Call CargaComboConstante(gPersEstadoCivil, cmbPersNatEstCiv)
'
'    'Carga Combo Relaciones Con La Institucion
'    CmbRela.Clear
'    Call CargaComboConstante(gPersRelacionInst, CmbRela)
    
    'Carga Ubicaciones Geograficas
'    Call CargaUbicacionesGeograficas
    
    'Carga Ciiu
'    If gsCodCMAC = "102" Then
'        sSql = " Select cCIIUcod,cCIIUdescripcion, 1 Nro from CIIU  where cciiucod in ('O9309')  " _
'             & "  union all " _
'             & "  Select cCIIUcod ,cCIIUdescripcion, 2 nro  from CIIU where cciiucod not in ('O9309')" _
'             & " order by nro, cCIIUdescripcion "
'    Else
'        sSql = "Select cCIIUcod,cCIIUdescripcion from CIIU Order by cCIIUdescripcion"
'    End If
'    Set oPersonas = New COMDPersona.DCOMPersonas
'
'    Set R = oPersonas.Cargar_CIIU(gsCodCMAC)
'    Do While Not R.EOF
'        CboPersCiiu.AddItem Trim(R!cCIIUdescripcion) & Space(100) & Trim(R!cCIIUcod)
'        R.MoveNext
'    Loop
'    R.Close
'
'    'CARGA TIPOS DE PERSONA JURIDICA
'    cmbPersJurTpo.Clear
'
'    'sSql = "Select cPersJurTpoCod,cPersJurTpoDesc  from persjurtpo Order by cPersJurTpoDesc"
'    'Set R = Conn.CargaRecordSet(sSql)
'    Set R = oPersonas.CargarTipos_PerJuridicas
'    Do While Not R.EOF
'        cmbPersJurTpo.AddItem Trim(R!cPersJurTpoDesc) & Space(100) & Trim(R!cPersJurTpoCod)
'        R.MoveNext
'    Loop
'
'    R.Close
    
'    Set oPersonas = Nothing
'    Set R = Nothing
'    Conn.CierraConexion
'    Set Conn = Nothing
    bEstadoCargando = True
    Exit Sub
    
    
ERRORCargaControles:
    MsgBox Err.Description, vbExclamation, "Aviso"
    
End Sub

'Funcion que carga todos los Objetos necesarios para Cargar el Formulario

Sub Cargar_Datos_Objetos_Persona()

Dim lrsOcupa As ADODB.Recordset 'madm 20100305
Dim oPersonas As COMDPersona.DCOMPersonas
Dim oCtasIF As COMNCajaGeneral.NCOMCajaCtaIF

Dim lrsRelac As ADODB.Recordset
Dim lrsVisita As ADODB.Recordset 'madm 20100329
Dim lrsDocumentos As ADODB.Recordset
Dim lrsRefComercial As ADODB.Recordset
Dim lrsPatVehicular As ADODB.Recordset

Dim lrsTipoSangre As ADODB.Recordset
Dim lrsDireccCondic As ADODB.Recordset
Dim lrsPersoneria As ADODB.Recordset
Dim lrsMagnitudComer As ADODB.Recordset
Dim lrsMagnitudPersNat As ADODB.Recordset 'JACA 20110427
Dim lrsEstCivil As ADODB.Recordset
Dim lrsRelacInst As ADODB.Recordset

Dim lrsUbiGeo  As ADODB.Recordset
Dim i As Integer
Dim lrsCIIU As ADODB.Recordset
Dim lrsCargo As ADODB.Recordset '** Juez 20120326

'*** PEAC
Dim lrsTIPOCOMP As ADODB.Recordset
Dim lrsTIPOSISTINFOR As ADODB.Recordset
Dim lrsCADENAPROD As ADODB.Recordset
Dim lrsTipoPatri As ADODB.Recordset
Dim lrsMonePatri As ADODB.Recordset
Dim lrsAlterSiNo As ADODB.Recordset
Dim lrsRRPP As ADODB.Recordset
Dim lcCampo As String
Dim lrsMotivoActu As ADODB.Recordset '20080801
'*** FIN PEAC
Set lrsRRPP = New ADODB.Recordset

lrsRRPP.Fields.Append "Alternativa", adVarChar, 2
lrsRRPP.Open
For i = 0 To 1
    lcCampo = IIf(i = 0, "SI", "NO")
    lrsRRPP.AddNew
    lrsRRPP.Fields("Alternativa") = lcCampo
    lrsRRPP.MoveNext
Next

Dim lrsTipoPerJuri As ADODB.Recordset

On Error GoTo ErrCargarDatosObjetosPersona

Set oPersonas = New COMDPersona.DCOMPersonas

Call oPersonas.CargarDatosObjetosPersona(lrsRelac, lrsDocumentos, lrsRefComercial, lrsPatVehicular, lrsTipoSangre, lrsDireccCondic, lrsPersoneria, lrsMagnitudComer, lrsEstCivil, lrsRelacInst, _
                                        lrsUbiGeo, lrsCIIU, gsCodCMAC, lrsTipoPerJuri, lrsVisita, lrsTIPOCOMP, lrsTIPOSISTINFOR, lrsCADENAPROD, lrsTipoPatri, lrsMonePatri, lrsAlterSiNo, lrsMotivoActu, lrsOcupa, lrsCargo)

'Set oPersonas = Nothing JACA 20110427
'MADM 20100406 Carga Combo de Flex de Visitas de Persona
    FEVisitas.CargaCombo lrsVisita
'Carga Combo de Flex de Relaciones de Persona
    FERelPers.CargaCombo lrsRelac
'Carga Combo de Documentos
 'ALPA 20080922***************************************************************************
    nNumeroDoc = 0
    Do While Not lrsDocumentos.EOF
        nNumeroDoc = nNumeroDoc + 1
        ReDim Preserve MatrixTipoDoc(0 To 3, 0 To nNumeroDoc)
        MatrixTipoDoc(1, nNumeroDoc) = lrsDocumentos!cConsDescripcion
        MatrixTipoDoc(2, nNumeroDoc) = lrsDocumentos!nConsValor
        If lrsDocumentos!nConsValor = 1 Or lrsDocumentos!nConsValor = 3 Then
            MatrixTipoDoc(3, nNumeroDoc) = 1
        ElseIf lrsDocumentos!nConsValor = 4 Or lrsDocumentos!nConsValor = 11 Then
            MatrixTipoDoc(3, nNumeroDoc) = 2
        Else
            MatrixTipoDoc(3, nNumeroDoc) = 0
        End If
        lrsDocumentos.MoveNext
    Loop
'******************************************************************************************
   'nNumeroDoc = 0
    FEDocs.CargaCombo lrsDocumentos

'Carga Combo de Tipos de Referencia Comercial
    feRefComercial.CargaCombo lrsRefComercial
    
'Carga Combo de Condicion de Patrimonio Vehicular
    fePatVehicular.CargaCombo lrsPatVehicular
    fePatInmuebles.CargaCombo lrsAlterSiNo ''lrsRRPP

'cargar combo ocupaciones
    'Call Llenar_Combo_con_Recordset(lrsOcupa, cboocupa) 'madm 20100305
    '** Juez 20120326 ************************************ Agregar mas espacio a los valores del combo para evitar que se vea el numero
    Do While Not lrsOcupa.EOF
        cboocupa.AddItem Trim(lrsOcupa!cConsDescripcion) & Space(200) & Trim(lrsOcupa!nConsValor)
        lrsOcupa.MoveNext
    Loop
    '** End Juez *****************************************
'FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)

'Carga combo cargos
    Call Llenar_Combo_con_Recordset(lrsCargo, cboCargos) '** Juez 20120326
        
'***********Combos de Constantes
Set RsCIIUTemp = lrsCIIU


'Carga Tipos de Sangre
    Call Llenar_Combo_con_Recordset(lrsTipoSangre, CboTipoSangre)
'Carga Condiciones de Domicilio
    Call Llenar_Combo_con_Recordset(lrsDireccCondic, cmbPersDireccCondicion)
'Carga Personeria
    Call Llenar_Combo_con_Recordset(lrsPersoneria, cmbPersPersoneria)
'Carga Magnitud Empresarial
    Call Llenar_Combo_con_Recordset(lrsMagnitudComer, cmbPersJurMagnitud)
'JACA 20110427*********************************************************************
'Carga Magnitud Persona
    
    Call oPersonas.CargarMagnitudPersona(lrsMagnitudPersNat, 1)
    Set oPersonas = Nothing
    Call Llenar_Combo_con_Recordset(lrsMagnitudPersNat, cmbPersNatMagnitud)
'JACA END**************************************************************************
'Carga Estado Civil
    Call Llenar_Combo_con_Recordset(lrsEstCivil, cmbPersNatEstCiv)
'Carga Combo Relaciones Con La Institucion
    Call Llenar_Combo_con_Recordset(lrsRelacInst, CmbRela)

'*** PEAC 20080412
'Carga Combo Obtiene tipo competencia
    Call Llenar_Combo_con_Recordset(lrsTIPOCOMP, cboTipoComp)
'Carga Combo Obtiene tipo sistema informacion
    Call Llenar_Combo_con_Recordset(lrsTIPOSISTINFOR, cboTipoSistInfor)
'Carga Combo Obtiene cadena productiva
    Call Llenar_Combo_con_Recordset(lrsCADENAPROD, cboCadenaProd)
'Carga Combo tipos de patrimonio
    Call Llenar_Combo_con_Recordset(lrsTipoPatri, Me.cboTipoPatri)

'Carga Combo moneda patrimonio
    Call Llenar_Combo_con_Recordset(lrsMonePatri, Me.cboMonePatri)
    
'*** FIN PEAC
   
'*** PEAC 20080801
'Carga Combo motivo de actualizacion
    Call Llenar_Combo_con_Recordset(lrsMotivoActu, Me.cboMotivoActu)
    'cboMotivoActu.RemoveItem (2)
   
'Carga Tipos de Sexo de Personas
    cmbPersNatSexo.AddItem "FEMENINO" & Space(50) & "F"
    cmbPersNatSexo.AddItem "MASCULINO" & Space(50) & "M"

'Otros Objetos
    Set oCtasIF = New COMNCajaGeneral.NCOMCajaCtaIF  'NCajaCtaIF
    feRefBancaria.psRaiz = "BANCOS"
    feRefBancaria.rsTextBuscar = oCtasIF.GetInstFinancieras("0[123]")
    Set oCtasIF = Nothing
        
'Cargar CIIU
    Do While Not lrsCIIU.EOF
        CboPersCiiu.AddItem Trim(lrsCIIU!cCIIUdescripcion) & Space(200) & Trim(lrsCIIU!cCIIUcod)
        lrsCIIU.MoveNext
    Loop
      
'CARGA TIPOS DE PERSONA JURIDICA
    cmbPersJurTpo.Clear
    Do While Not lrsTipoPerJuri.EOF
        cmbPersJurTpo.AddItem Trim(lrsTipoPerJuri!cPersJurTpoDesc) & Space(100) & Trim(lrsTipoPerJuri!cPersJurTpoCod)
        lrsTipoPerJuri.MoveNext
    Loop


'Ubicaciones Geograficas
    
    While Not lrsUbiGeo.EOF
        cmbPersUbiGeo(0).AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        cmbNacionalidad.AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        lrsUbiGeo.MoveNext
    Wend
    
    'JUEZ 20131007 *****************************************
    lrsUbiGeo.MoveFirst
    While Not lrsUbiGeo.EOF
        cmbNegUbiGeo(0).AddItem Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
        lrsUbiGeo.MoveNext
    Wend
    cboRemInfoEmail.ListIndex = 0
    'END JUEZ **********************************************
    
    
'    ContNiv1 = 0
'    ContNiv2 = 0
'    ContNiv3 = 0
'    ContNiv4 = 0
'    ContNiv5 = 0
    
'    Do While Not lrsUbiGeo.EOF
'        Select Case lrsUbiGeo!P
'            Case 1 'Pais
'                ContNiv1 = ContNiv1 + 1
'                ReDim Preserve Nivel1(ContNiv1)
'                Nivel1(ContNiv1 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'            Case 2 ' Departamento
'                ContNiv2 = ContNiv2 + 1
'                ReDim Preserve Nivel2(ContNiv2)
'                Nivel2(ContNiv2 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'            Case 3 'Provincia
'                ContNiv3 = ContNiv3 + 1
'                ReDim Preserve Nivel3(ContNiv3)
'                Nivel3(ContNiv3 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'            Case 4 'Distrito
'                ContNiv4 = ContNiv4 + 1
'                ReDim Preserve Nivel4(ContNiv4)
'                Nivel4(ContNiv4 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'            Case 5 'Zona
'                ContNiv5 = ContNiv5 + 1
'                ReDim Preserve Nivel5(ContNiv5)
'                Nivel5(ContNiv5 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'        End Select
'        lrsUbiGeo.MoveNext
'    Loop
    
'    If lrsUbiGeo.EOF Then Exit Sub
'
'    Do While lrsUbiGeo!P = 1
'        ContNiv1 = ContNiv1 + 1
'        ReDim Preserve Nivel1(ContNiv1)
'        Nivel1(ContNiv1 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'        lrsUbiGeo.MoveNext
'    Loop
'
'    Do While lrsUbiGeo!P = 2
'        ContNiv2 = ContNiv2 + 1
'        ReDim Preserve Nivel2(ContNiv2)
'        Nivel2(ContNiv2 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'        lrsUbiGeo.MoveNext
'    Loop
'
'    Do While lrsUbiGeo!P = 3
'        ContNiv3 = ContNiv3 + 1
'        ReDim Preserve Nivel3(ContNiv3)
'        Nivel3(ContNiv3 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'        lrsUbiGeo.MoveNext
'    Loop
'
'    Do While lrsUbiGeo!P = 4
'        ContNiv4 = ContNiv4 + 1
'        ReDim Preserve Nivel4(ContNiv4)
'        Nivel4(ContNiv4 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'        lrsUbiGeo.MoveNext
'    Loop
'
'    Do While lrsUbiGeo!P = 5
'        ContNiv5 = ContNiv5 + 1
'        ReDim Preserve Nivel5(ContNiv5)
'        Nivel5(ContNiv5 - 1) = Trim(lrsUbiGeo!cUbiGeoDescripcion) & Space(50) & Trim(lrsUbiGeo!cUbiGeoCod)
'        lrsUbiGeo.MoveNext
'        If lrsUbiGeo.EOF Then Exit Do
'    Loop
'
'    'Carga el Nivel1 en el Control
'    cmbPersUbiGeo(0).Clear
'    For i = 0 To ContNiv1 - 1
'        cmbPersUbiGeo(0).AddItem Nivel1(i)
'        cmbNacionalidad.AddItem Nivel1(i)
'        If Trim(Right(Nivel1(i), 10)) = "04028" Then
'            nPos = i
'        End If
'    Next i
    If lrsUbiGeo.RecordCount > 0 Then lrsUbiGeo.MoveFirst
    For i = 0 To lrsUbiGeo.RecordCount
        If Trim(lrsUbiGeo!cUbiGeoCod) = "04028" Then
            nPos = i
        End If
    Next i

    cmbPersUbiGeo(0).ListIndex = nPos
'    cmbPersUbiGeo(2).Clear
'    cmbPersUbiGeo(3).Clear
'    cmbPersUbiGeo(4).Clear
    cmbNacionalidad.ListIndex = nPos
    
    '*** PEAC 20080801
    Me.cboMotivoActu.ListIndex = nPos
        
    Exit Sub
    
ErrCargarDatosObjetosPersona:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub
'Sub Llenar_Combo_con_Recordset(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
'
'pcboObjeto.Clear
'Do While Not pRs.EOF
'    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim(Str(pRs!nConsValor))
'    pRs.MoveNext
'Loop
'pRs.Close
'
'End Sub

Private Sub CargaControlEstadoPersona(ByVal pnTipoPers As Integer)
'Dim Conn As COMConecta.DCOMConecta  'DConecta
Dim ssql As String
Dim R As ADODB.Recordset
'
'On Error GoTo ERRORCargaControles
'    Set Conn = New COMConecta.DCOMConecta 'DConecta
'    Conn.AbreConexion
cmbPersEstado.Clear
Dim oPersonas As New COMDPersona.DCOMPersonas
    'Carga Estados de la Persona
    'sSql = "Select nConsValor,cConsDescripcion From Constante Where nConsCod = " & Trim(Str(gPersEstado)) & " and nConsValor <> " & Trim(Str(gPersEstado))
    'Set R = Conn.CargaRecordSet(sSql)
    Set R = oPersonas.CargarEstadosPersona()
    Do While Not R.EOF
        If pnTipoPers = gPersonaNat Then
            If Len(Trim(R!nConsValor)) = 1 Then
                cmbPersEstado.AddItem Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(str(R!nConsValor)), 2)
            End If
        Else
            If Len(Trim(R!nConsValor)) > 1 Then
                cmbPersEstado.AddItem Trim(R!cConsDescripcion) & Space(50) & Right("0" & Trim(R!nConsValor), 2)
            End If
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
'    Conn.CierraConexion
'    Set Conn = Nothing
    Exit Sub
    
'ERRORCargaControles:
'    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub DistribuyeApellidos(ByVal bApellCasada As Boolean)
   If bApellCasada = True Then
        lblApCasada.Visible = True
        txtApellidoCasada.Visible = True
        'JACA 20110426*************************************************
'        lblPersNombreAP.Top = 660
'        lblPersNombreAP.Left = 120
'        txtPersNombreAP.Top = 660
'        txtPersNombreAP.Left = 1680
'
'        lblPersNombreAM.Top = 1065
'        lblPersNombreAM.Left = 120
'        txtPersNombreAM.Top = 1065
'        txtPersNombreAM.Left = 1680
'
'        lblApCasada.Top = 1485
'        lblApCasada.Left = 135
'
'        txtApellidoCasada.Top = 1485
'        txtApellidoCasada.Left = 1680
'
'        lblPersNombreN.Top = 1875
'        lblPersNombreN.Left = 135
'        txtPersNombreN.Top = 1875
'        txtPersNombreN.Left = 1680
        
               
        lblPersNombreAM.Top = 840
        txtPersNombreAM.Top = 840
        
        lblApCasada.Top = 1200
        txtApellidoCasada.Top = 1200
        
        lblPersNombreN.Top = 1560
        txtPersNombreN.Top = 1560
        
        'JACA END********************************************************
    Else
        'JACA 20110426*****************************************************
'        lblPersNombreAP.Top = 660
'        lblPersNombreAP.Left = 120
'        txtPersNombreAP.Top = 660
'        txtPersNombreAP.Left = 1680
'
'        lblPersNombreAM.Top = 1140
'        lblPersNombreAM.Left = 120
'        txtPersNombreAM.Top = 1140
'        txtPersNombreAM.Left = 1680
'
'        lblApCasada.Visible = False
'        txtApellidoCasada.Visible = False
'
'        lblPersNombreN.Top = 1620
'        lblPersNombreN.Left = 120
'        txtPersNombreN.Top = 1620
'        txtPersNombreN.Left = 1680
        
        
        lblPersNombreAM.Top = 960
        txtPersNombreAM.Top = 960
                
        lblApCasada.Visible = False
        txtApellidoCasada.Visible = False
        
        lblPersNombreN.Top = 1440
        txtPersNombreN.Top = 1440
        
        'JACA END***********************************************************
    End If
End Sub

Private Sub CargaDocumentos()
Dim i As Integer
    Call LimpiaFlex(FEDocs)
    For i = 0 To oPersona.NumeroDocumentos - 1
    'For i = 0 To fnNumDocs - 1
        FEDocs.AdicionaFila
        'Columna de Tipo de Documento
        FEDocs.TextMatrix(i + 1, 1) = Trim(oPersona.ObtenerTipoDoc(i))
        'Columna de Numero de Documento
        FEDocs.TextMatrix(i + 1, 2) = oPersona.ObtenerNumeroDoc(i)
    Next i
End Sub
'FRHU 20151130 ERS077-2015
Private Sub CargaDocumentosParaActAutDatosIni()
    Dim i As Integer
    For i = 0 To oPersona.NumeroDocumentos - 1
        If Trim(Right(oPersona.ObtenerTipoDoc(i), 2)) = "1" Then
            'Columna de Tipo de Documento
            MatPersona(1).sPersIDTpo = Trim(Right(oPersona.ObtenerTipoDoc(i), 2))
            'Columna de Numero de Documento
            MatPersona(1).sPersIDnro = oPersona.ObtenerNumeroDoc(i)
            Exit For
        End If
        MatPersona(1).sPersIDTpo = Trim(Right(oPersona.ObtenerTipoDoc(i), 2))
        MatPersona(1).sPersIDnro = oPersona.ObtenerNumeroDoc(i)
    Next i
End Sub
Private Sub CargaDocumentosParaActAutDatos()
Dim i As Integer
    Call LimpiaFlex(FEDocs)
    For i = 0 To oPersona.NumeroDocumentos - 1
        If Trim(Right(oPersona.ObtenerTipoDoc(i), 2)) = "1" Then
            'Columna de Tipo de Documento
            MatPersona(2).sPersIDTpo = Trim(Right(oPersona.ObtenerTipoDoc(i), 2))
            'Columna de Numero de Documento
            MatPersona(2).sPersIDnro = oPersona.ObtenerNumeroDoc(i)
            Exit For
        End If
        MatPersona(2).sPersIDTpo = Trim(Right(oPersona.ObtenerTipoDoc(i), 2))
        MatPersona(2).sPersIDnro = oPersona.ObtenerNumeroDoc(i)
    Next i
End Sub
'FIN FRHU 20151130
Private Sub CargaRelacionesPersonas()
Dim i As Integer
    FERelPers.lbEditarFlex = True
    Call LimpiaFlex(FERelPers)
    For i = 0 To oPersona.NumeroRelacPers - 1
        FERelPers.AdicionaFila
        'Codigo
        FERelPers.TextMatrix(i + 1, 1) = oPersona.ObtenerRelaPersCodigo(i)
        'Apellidos y Nombres
        FERelPers.TextMatrix(i + 1, 2) = oPersona.ObtenerRelaPersNombres(i)
        'Relacion
        FERelPers.TextMatrix(i + 1, 3) = oPersona.ObtenerRelaPersRelacion(i)
        'Beneficiario
        FERelPers.TextMatrix(i + 1, 4) = oPersona.ObtenerRelaPersBenef(i)
        'Beneficiario Porcentaje
        FERelPers.TextMatrix(i + 1, 5) = Format(oPersona.ObtenerRelaPersBenefPorc(i), "#0.00")
        'Asistencia medica Privada
        FERelPers.TextMatrix(i + 1, 6) = oPersona.ObtenerRelaPersAMP(i)
    Next i
    FERelPers.lbEditarFlex = False
End Sub
'madm 20100329
Private Sub CargaVisitasPersonas()
Dim i As Integer
    FEVisitas.lbEditarFlex = True
    Call LimpiaFlex(FEVisitas)
    For i = 0 To oPersona.NumeroVisitaPers - 1
        FEVisitas.AdicionaFila
        'Codigo
        FEVisitas.TextMatrix(i + 1, 1) = oPersona.ObtenerVisitaPersCodOrig(i)
        'Apellidos y Nombres
        FEVisitas.TextMatrix(i + 1, 2) = oPersona.ObtenerVisitaApellNombres(i)
        'Relacion
        FEVisitas.TextMatrix(i + 1, 3) = oPersona.ObtenerVisitaDireccion(i)
        'Beneficiario
        FEVisitas.TextMatrix(i + 1, 4) = Format(oPersona.ObtenerVisitaFecha(i), "dd/mm/yyyy")
        'Beneficiario Porcentaje
        FEVisitas.TextMatrix(i + 1, 5) = oPersona.ObtenerVisitaUsual(i)
        'Asistencia medica Privada
        FEVisitas.TextMatrix(i + 1, 6) = oPersona.ObtenerVisitaObserva(i)
    Next i
    FEVisitas.lbEditarFlex = False
End Sub
'end madm

'MAVM 20100607 BAS II
Private Sub CargaVentasPersonas()
Dim i As Integer
    FEVentas.lbEditarFlex = True
    Call LimpiaFlex(FEVentas)
    For i = 0 To oPersona.NumeroVentasPers - 1
        FEVentas.AdicionaFila
        'Codigo
        FEVentas.TextMatrix(i + 1, 1) = oPersona.ObtenerVentasPersCod(i)
        'Apellidos y Nombres
        FEVentas.TextMatrix(i + 1, 2) = oPersona.ObtenerVentasApellNombres(i)
        
        'Monto
        FEVentas.TextMatrix(i + 1, 3) = Format$(oPersona.ObtenerVentasMonto(i), "###,###.00")
        'Fecha
        FEVentas.TextMatrix(i + 1, 4) = Format(oPersona.ObtenerVentasFecha(i), "dd/mm/yyyy")
        'Periodo
        FEVentas.TextMatrix(i + 1, 5) = oPersona.ObtenerVentasPeriodo(i)
    Next i
    FEVentas.lbEditarFlex = False
End Sub

Private Sub CargaRefComerciales()
Dim i As Integer

    feRefComercial.lbEditarFlex = True
    Call LimpiaFlex(feRefComercial)
    For i = 0 To oPersona.NumeroRefComercial - 1
        feRefComercial.AdicionaFila
        feRefComercial.TextMatrix(i + 1, 1) = oPersona.ObtenerRefComNombre(i) 'Nombre/Razón Social
        feRefComercial.TextMatrix(i + 1, 2) = oPersona.ObtenerRefComRelacion(i) 'Tipo de Relacion con la Referencia
        feRefComercial.TextMatrix(i + 1, 3) = oPersona.ObtenerRefComComentario(i) 'Comentario de la Referencia n
        feRefComercial.TextMatrix(i + 1, 4) = oPersona.ObtenerRefComFono(i) 'Telefono de la Referencia
        feRefComercial.TextMatrix(i + 1, 5) = oPersona.ObtenerRefComDireccion(i) 'Direccion de la Referencia
        feRefComercial.TextMatrix(i + 1, 6) = oPersona.ObtenerRefComNumRef(i) 'Número de Referencia
    Next i
    
    feRefComercial.lbEditarFlex = False
    
End Sub

Private Sub CargaRefBancarias()
Dim i As Integer

    feRefBancaria.lbEditarFlex = True
    Call LimpiaFlex(feRefBancaria)
    For i = 0 To oPersona.NumeroRefBancaria - 1
        feRefBancaria.AdicionaFila
        feRefBancaria.TextMatrix(i + 1, 1) = oPersona.ObtenerRefBanCodIF(i) 'Codigo Institución Financiera
        feRefBancaria.TextMatrix(i + 1, 2) = oPersona.ObtenerRefBanNombre(i) 'Tipo de Relacion con la Referencia
        feRefBancaria.TextMatrix(i + 1, 3) = oPersona.ObtenerRefBanNumCta(i) 'Número de Cuenta
        feRefBancaria.TextMatrix(i + 1, 4) = oPersona.ObtenerRefBanNumTar(i) 'Número de Tarjeta
        feRefBancaria.TextMatrix(i + 1, 5) = Format$(oPersona.ObtenerRefBanLinCred(i), "###,###.00") 'Monto de la Línea de Crédito
    Next i
    
    feRefBancaria.lbEditarFlex = False
    
End Sub

'*** PEAC 20080412
Private Sub CargaPatInmuebles()
Dim i As Integer

    fePatInmuebles.lbEditarFlex = True
    Call LimpiaFlex(fePatInmuebles)
    For i = 0 To oPersona.NumeroPatInmuebles - 1
        fePatInmuebles.AdicionaFila
        fePatInmuebles.TextMatrix(i + 1, 1) = oPersona.ObtenerPatInmueblesUbicacion(i) 'ubicacion
        fePatInmuebles.TextMatrix(i + 1, 2) = oPersona.ObtenerPatInmueblesAreaTerreno(i) 'ubicacion
        fePatInmuebles.TextMatrix(i + 1, 3) = oPersona.ObtenerPatInmueblesAreaConstru(i) 'ubicacion
        fePatInmuebles.TextMatrix(i + 1, 4) = oPersona.ObtenerPatInmueblesTipoUso(i) 'ubicacion
        fePatInmuebles.TextMatrix(i + 1, 5) = oPersona.ObtenerPatInmueblesRRPP(i) 'ubicacion
        fePatInmuebles.TextMatrix(i + 1, 6) = Format$(oPersona.ObtenerPatInmueblesValCom(i), "###,###.00") 'Valor Comercial
        fePatInmuebles.TextMatrix(i + 1, 7) = oPersona.ObtenerPatInmueblesCod(i) 'Codigo del Patrimonio Vehicular
    Next i
    
    fePatInmuebles.lbEditarFlex = False
    
End Sub

'*** PEAC 20080412
Private Sub CargaPatOtros()
Dim i As Integer

    fePatOtros.lbEditarFlex = True
    Call LimpiaFlex(fePatOtros)
    For i = 0 To oPersona.NumeroPatOtros - 1
        fePatOtros.AdicionaFila
        fePatOtros.TextMatrix(i + 1, 1) = oPersona.ObtenerPatOtrosDescripcion(i) 'Descripcion
        fePatOtros.TextMatrix(i + 1, 2) = Format$(oPersona.ObtenerPatOtrosValCom(i), "###,###.00") 'Valor Comercial
        fePatOtros.TextMatrix(i + 1, 3) = oPersona.ObtenerPatOtrosCod(i) 'Codigo del Patrimonio Vehicular
    Next i
    
    fePatOtros.lbEditarFlex = False
    
End Sub


Private Sub CargaPatVehicular()
Dim i As Integer

    fePatVehicular.lbEditarFlex = True
    Call LimpiaFlex(fePatVehicular)
    For i = 0 To oPersona.NumeroPatVehicular - 1
        fePatVehicular.AdicionaFila
        fePatVehicular.TextMatrix(i + 1, 1) = oPersona.ObtenerPatVehMarca(i) 'Marca del Patrimonio Vehicular
        fePatVehicular.TextMatrix(i + 1, 2) = oPersona.ObtenerPatVehFecFab(i) 'Fecha de Fabricacion
        fePatVehicular.TextMatrix(i + 1, 3) = Format$(oPersona.ObtenerPatVehValCom(i), "###,###.00") 'Valor Comercial
        'fePatVehicular.TextMatrix(i + 1, 4) = oPersona.ObtenerPatVehCondicion(i) 'Condicion del Patrimonio
        fePatVehicular.TextMatrix(i + 1, 4) = oPersona.ObtenerPatVehModelo(i) 'Condicion del Patrimonio
        fePatVehicular.TextMatrix(i + 1, 5) = oPersona.ObtenerPatVehPlaca(i) 'Condicion del Patrimonio
        fePatVehicular.TextMatrix(i + 1, 6) = oPersona.ObtenerPatVehCod(i) 'Codigo del Patrimonio Vehicular
    Next i
    
    fePatVehicular.lbEditarFlex = False
    
End Sub

Private Sub CargaFuentesIngreso()
Dim i As Integer
Dim MatFte As Variant

    Call LimpiaFlex(FEFteIng)
    
    MatFte = oPersona.FiltraFuentesIngresoPorRazonSocial
    
    'For i = 0 To oPersona.NumeroFtesIngreso - 1
    '    FEFteIng.AdicionaFila
    '    FEFteIng.TextMatrix(i + 1, 0) = i + 1
    '    FEFteIng.TextMatrix(i + 1, 1) = IIf(oPersona.ObtenerFteIngTipo(i) = "1", "D", "I") 'Tipo de Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 2) = oPersona.ObtenerFteIngRazonSocial(i) 'Razon Social de Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 3) = Format(oPersona.ObtenerFteIngFecCaducac(i), "dd/mm/yyyy") 'Fecha de Caducacion de la Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 4) = Format(oPersona.ObtenerFteIngFecEval(i), "dd/mm/yyyy") 'Fecha de Evaluacion de la Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 5) = IIf(oPersona.ObtenerFteIngMoneda(i) = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
    '    FEFteIng.TextMatrix(i + 1, 5) = IIf(oPersona.ObtenerFteIngMoneda(i) = gMonedaNacional, "SOLES", "DOLARES") 'Moneda de la Fte de Ingreso
    'Next i
    
    If IsArray(MatFte) Then
        For i = 0 To UBound(MatFte) - 1
            FEFteIng.AdicionaFila
            FEFteIng.TextMatrix(i + 1, 0) = MatFte(i, 0)
            FEFteIng.TextMatrix(i + 1, 1) = MatFte(i, 1)
            FEFteIng.TextMatrix(i + 1, 2) = MatFte(i, 2)
            FEFteIng.TextMatrix(i + 1, 3) = MatFte(i, 5)
            FEFteIng.TextMatrix(i + 1, 4) = MatFte(i, 6)
            FEFteIng.TextMatrix(i + 1, 5) = MatFte(i, 7)
        Next i
    End If
End Sub

'***********CTI3 10092018
Private Sub CargaDatosAdicionales()
Dim i As Integer
Dim MatDatAdc As Variant
'frmPersonaJurDatosAdic.Show 1
'frmPersonaJurDatosAdic.Hide
    Call LimpiaFlex(frmPersonaJurDatosAdic.fleAccionistas) '1
    Call LimpiaFlex(frmPersonaJurDatosAdic.fleDirectorio) '2
    Call LimpiaFlex(frmPersonaJurDatosAdic.fleGerencias) '3
    Call LimpiaFlex(frmPersonaJurDatosAdic.flePatrimonio) '4
    Call LimpiaFlex(frmPersonaJurDatosAdic.flePatOtrasEmpresa) '5
    Call LimpiaFlex(frmPersonaJurDatosAdic.fleCargos)  '6
    
   MatDatAdc = oPersona.RecuperaLlenado

    If IsArray(MatDatAdc) Then
        For i = 0 To UBound(MatDatAdc)
           If MatDatAdc(i, 9) = 1 Then
                frmPersonaJurDatosAdic.fleAccionistas.AdicionaFila
                frmPersonaJurDatosAdic.fleAccionistas.TextMatrix(frmPersonaJurDatosAdic.fleAccionistas.rows - 1, 1) = MatDatAdc(i, 0)
                frmPersonaJurDatosAdic.fleAccionistas.TextMatrix(frmPersonaJurDatosAdic.fleAccionistas.rows - 1, 2) = MatDatAdc(i, 1)
                frmPersonaJurDatosAdic.fleAccionistas.TextMatrix(frmPersonaJurDatosAdic.fleAccionistas.rows - 1, 3) = MatDatAdc(i, 2)
                frmPersonaJurDatosAdic.fleAccionistas.TextMatrix(frmPersonaJurDatosAdic.fleAccionistas.rows - 1, 4) = MatDatAdc(i, 3)
                frmPersonaJurDatosAdic.fleAccionistas.TextMatrix(frmPersonaJurDatosAdic.fleAccionistas.rows - 1, 5) = Format(MatDatAdc(i, 6), "#,#0.00")
                frmPersonaJurDatosAdic.fleAccionistas.TextMatrix(frmPersonaJurDatosAdic.fleAccionistas.rows - 1, 6) = MatDatAdc(i, 7)
           End If
           If MatDatAdc(i, 9) = 2 Then
                frmPersonaJurDatosAdic.fleDirectorio.AdicionaFila
                frmPersonaJurDatosAdic.fleDirectorio.TextMatrix(frmPersonaJurDatosAdic.fleDirectorio.rows - 1, 1) = MatDatAdc(i, 0)
                frmPersonaJurDatosAdic.fleDirectorio.TextMatrix(frmPersonaJurDatosAdic.fleDirectorio.rows - 1, 2) = MatDatAdc(i, 1)
                frmPersonaJurDatosAdic.fleDirectorio.TextMatrix(frmPersonaJurDatosAdic.fleDirectorio.rows - 1, 3) = MatDatAdc(i, 2)
                frmPersonaJurDatosAdic.fleDirectorio.TextMatrix(frmPersonaJurDatosAdic.fleDirectorio.rows - 1, 4) = MatDatAdc(i, 4)
                frmPersonaJurDatosAdic.fleDirectorio.TextMatrix(frmPersonaJurDatosAdic.fleDirectorio.rows - 1, 5) = MatDatAdc(i, 3)
           End If
           If MatDatAdc(i, 9) = 3 Then
                frmPersonaJurDatosAdic.fleGerencias.AdicionaFila
                frmPersonaJurDatosAdic.fleGerencias.TextMatrix(frmPersonaJurDatosAdic.fleGerencias.rows - 1, 1) = MatDatAdc(i, 0)
                frmPersonaJurDatosAdic.fleGerencias.TextMatrix(frmPersonaJurDatosAdic.fleGerencias.rows - 1, 2) = MatDatAdc(i, 1)
                frmPersonaJurDatosAdic.fleGerencias.TextMatrix(frmPersonaJurDatosAdic.fleGerencias.rows - 1, 3) = MatDatAdc(i, 2)
                frmPersonaJurDatosAdic.fleGerencias.TextMatrix(frmPersonaJurDatosAdic.fleGerencias.rows - 1, 4) = MatDatAdc(i, 4)

           End If
           If MatDatAdc(i, 9) = 4 Then
                frmPersonaJurDatosAdic.flePatrimonio.AdicionaFila
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 3) = MatDatAdc(i, 0)
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 4) = MatDatAdc(i, 5)
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 5) = MatDatAdc(i, 1)
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 6) = Format(MatDatAdc(i, 6), "#,#0.00")
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 7) = MatDatAdc(i, 7)
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 2) = MatDatAdc(i, 8)
                frmPersonaJurDatosAdic.flePatrimonio.TextMatrix(frmPersonaJurDatosAdic.flePatrimonio.rows - 1, 1) = "0"
                frmPersonaJurDatosAdic.flePatrimonio.lbEditarFlex = False
           End If
           If MatDatAdc(i, 9) = 5 Then
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.AdicionaFila
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 3) = MatDatAdc(i, 0)
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 4) = MatDatAdc(i, 5)
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 5) = MatDatAdc(i, 1)
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 6) = Format(MatDatAdc(i, 6), "#,#0.00")
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 7) = MatDatAdc(i, 7)
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 2) = MatDatAdc(i, 8)
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.TextMatrix(frmPersonaJurDatosAdic.flePatOtrasEmpresa.rows - 1, 1) = 0
                frmPersonaJurDatosAdic.flePatOtrasEmpresa.lbEditarFlex = False
           End If
           If MatDatAdc(i, 9) = 6 Then
                frmPersonaJurDatosAdic.fleCargos.AdicionaFila
                frmPersonaJurDatosAdic.fleCargos.TextMatrix(frmPersonaJurDatosAdic.fleCargos.rows - 1, 3) = MatDatAdc(i, 0)
                frmPersonaJurDatosAdic.fleCargos.TextMatrix(frmPersonaJurDatosAdic.fleCargos.rows - 1, 4) = MatDatAdc(i, 5)
                frmPersonaJurDatosAdic.fleCargos.TextMatrix(frmPersonaJurDatosAdic.fleCargos.rows - 1, 5) = MatDatAdc(i, 1)
                frmPersonaJurDatosAdic.fleCargos.TextMatrix(frmPersonaJurDatosAdic.fleCargos.rows - 1, 6) = MatDatAdc(i, 4)
                frmPersonaJurDatosAdic.fleCargos.TextMatrix(frmPersonaJurDatosAdic.fleCargos.rows - 1, 2) = MatDatAdc(i, 8)
                frmPersonaJurDatosAdic.fleCargos.TextMatrix(frmPersonaJurDatosAdic.fleCargos.rows - 1, 1) = "0"
                frmPersonaJurDatosAdic.fleCargos.lbEditarFlex = False
           End If
        Next i
    End If
End Sub

'************************
Private Sub CargaDatos()
Dim i As Integer
     
    bEstadoCargando = True
    SSTDatosGen.Tab = 0
    'If SSTabs.TabVisible(0) Then
    '    SSTabs.Tab = 0
    'Else
    '    SSTabs.Tab = 1
    'End If
    'Carga Personeria
    SSTabs.TabVisible(0) = True
    SSTabs.TabVisible(1) = True
    'SSTabs.TabVisible(9) = True 'FRHU 20150311 ERS013-2015
        
    cmbPersPersoneria.ListIndex = IndiceListaCombo(cmbPersPersoneria, Trim(str(oPersona.Personeria)))
    
    'Habilita o Deshabilita Ficha de Persona Juridica
    
    Call HabilitaFichaPersonaJur(False)
    Call HabilitaFichaPersonaNat(False)
    
    'Carga Ubicacion Georgrafica
    If Len(Trim(oPersona.UbicacionGeografica)) = 12 Then
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "04028")
        cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "1" & Mid(oPersona.UbicacionGeografica, 2, 2) & String(9, "0"))
        cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "2" & Mid(oPersona.UbicacionGeografica, 2, 4) & String(7, "0"))
        cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & "3" & Mid(oPersona.UbicacionGeografica, 2, 6) & String(5, "0"))
        cmbPersUbiGeo(4).ListIndex = IndiceListaCombo(cmbPersUbiGeo(4), Space(30) & oPersona.UbicacionGeografica)
    Else
        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & oPersona.UbicacionGeografica)
        cmbPersUbiGeo(1).Clear
        cmbPersUbiGeo(1).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(1).ListIndex = 0
        cmbPersUbiGeo(2).Clear
        cmbPersUbiGeo(2).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(2).ListIndex = 0
        cmbPersUbiGeo(3).Clear
        cmbPersUbiGeo(3).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(3).ListIndex = 0
        cmbPersUbiGeo(4).Clear
        cmbPersUbiGeo(4).AddItem cmbPersUbiGeo(0).Text
        cmbPersUbiGeo(4).ListIndex = 0
    End If
    
    'Carga Direccion
    txtPersDireccDomicilio.Text = oPersona.Domicilio
    '*** PEAC 20080801
    txtRefDomicilio.Text = oPersona.RefDomicilio
    
    'Selecciona la Condicion del Doicilio
    cmbPersDireccCondicion.ListIndex = IndiceListaCombo(cmbPersDireccCondicion, oPersona.CondicionDomicilio)
    
    txtValComercial.Text = Format(oPersona.ValComDomicilio, "#####,###.00")
    
    'JUEZ 20131007 Ficha UbiGeo Negocio/Centro Laboral ******************
    If Len(Trim(oPersona.UbiGeoNegocio)) = 12 Then
        cmbNegUbiGeo(0).ListIndex = IndiceListaCombo(cmbNegUbiGeo(0), Space(30) & "04028")
        cmbNegUbiGeo(1).ListIndex = IndiceListaCombo(cmbNegUbiGeo(1), Space(30) & "1" & Mid(oPersona.UbiGeoNegocio, 2, 2) & String(9, "0"))
        cmbNegUbiGeo(2).ListIndex = IndiceListaCombo(cmbNegUbiGeo(2), Space(30) & "2" & Mid(oPersona.UbiGeoNegocio, 2, 4) & String(7, "0"))
        cmbNegUbiGeo(3).ListIndex = IndiceListaCombo(cmbNegUbiGeo(3), Space(30) & "3" & Mid(oPersona.UbiGeoNegocio, 2, 6) & String(5, "0"))
        cmbNegUbiGeo(4).ListIndex = IndiceListaCombo(cmbNegUbiGeo(4), Space(30) & oPersona.UbiGeoNegocio)
    Else
        cmbNegUbiGeo(0).ListIndex = IndiceListaCombo(cmbNegUbiGeo(0), Space(30) & oPersona.UbiGeoNegocio)
        cmbNegUbiGeo(1).Clear
        cmbNegUbiGeo(1).AddItem cmbNegUbiGeo(0).Text
        cmbNegUbiGeo(1).ListIndex = 0
        cmbNegUbiGeo(2).Clear
        cmbNegUbiGeo(2).AddItem cmbNegUbiGeo(0).Text
        cmbNegUbiGeo(2).ListIndex = 0
        cmbNegUbiGeo(3).Clear
        cmbNegUbiGeo(3).AddItem cmbNegUbiGeo(0).Text
        cmbNegUbiGeo(3).ListIndex = 0
        cmbNegUbiGeo(4).Clear
        cmbNegUbiGeo(4).AddItem cmbNegUbiGeo(0).Text
        cmbNegUbiGeo(4).ListIndex = 0
    End If
    
    txtNegDireccion.Text = oPersona.NegocioDireccion
    txtRefNegocio.Text = oPersona.RefNegocio
    Me.txtNombreCentroLaboral.Text = oPersona.NombreCentroLaboral
    
    cboRemInfoEmail.ListIndex = IndiceListaCombo(cboRemInfoEmail, oPersona.RemisionInfoEmail)
    'END JUEZ ***********************************************************
    
    'Tipo de sangre
    If CboTipoSangre.ListCount > 0 Then
        CboTipoSangre.ListIndex = IndiceListaCombo(CboTipoSangre, oPersona.TipoSangre)
    End If
    
    'Carga Ficha 1
    If oPersona.Personeria = gPersonaNat Then
        txtPersNombreAP.Text = oPersona.ApellidoPaterno
        txtPersNombreAM.Text = oPersona.ApellidoMaterno
        txtPersNombreN.Text = oPersona.Nombres
    Else
        txtPersNombreRS.Text = oPersona.NombreCompleto
    End If
    TxtTalla.Text = Format(oPersona.Talla, "#0.00")
    TxtPeso.Text = Format(oPersona.Peso, "#0.00")
    TxtEmail.Text = oPersona.Email
    
    '*** PEAC 20080412
    txtNumDependi.Text = oPersona.NumDependi
    txtActComple.Text = oPersona.ActComple
    txtNumPtosVta.Text = oPersona.NumPtosVta
    txtActiGiro.Text = oPersona.ActiGiro
    '*** FIN PEAC
    
    txtPersTelefono2.Text = oPersona.Telefonos2
    If oPersona.Sexo = "F" Then
        txtApellidoCasada.Text = oPersona.ApellidoCasada
        cmbPersNatSexo.ListIndex = 0
        'Call DistribuyeApellidos(True)
        'JACA 20110428*************************************************
        If oPersona.ApellidoCasada = "" Then
            Call DistribuyeApellidos(False)
        Else
            Call DistribuyeApellidos(True)
        End If
        'JACA END******************************************************
    Else
        cmbPersNatSexo.ListIndex = 1
        Call DistribuyeApellidos(False)
    End If
    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, oPersona.EstadoCivil)
    txtPersNatHijos.Text = Trim(str(oPersona.Hijos))
    '*** PEAC 20080412
    txtPersNatNumEmp.Text = Trim(str(oPersona.NumEmp))
    
    cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, Space(30) & oPersona.Nacionalidad)
    '*** PEAC 20080801
    Me.cboMotivoActu.ListIndex = IndiceListaCombo(cboMotivoActu, Space(30) & oPersona.MotivoActu)
    
    '** Juez 20120327 ************************************
    'chkResidente.value = oPersona.Residencia
    cboResidente.ListIndex = IndiceListaCombo(cboResidente, oPersona.Residencia)
    If cboResidente.ListIndex = 1 Then
        'If oPersona.PaisReside <> "" Then
            Dim oPersUbigeo As New COMDPersona.DCOMPersonas
            Set oPersUbigeo = New COMDPersona.DCOMPersonas
            Dim prsUbiGeo As ADODB.Recordset
            Set prsUbiGeo = oPersUbigeo.CargarUbicacionesGeograficas(True, 0)
            Do While Not prsUbiGeo.EOF
                If Trim(prsUbiGeo!cUbiGeoCod) <> "04028" Then 'JUEZ 20131007 Para que Perú no sea opción a elegir
                    cboPaisReside.AddItem Trim(prsUbiGeo!cUbiGeoDescripcion) & Space(100) & Trim(prsUbiGeo!cUbiGeoCod)
                End If
                prsUbiGeo.MoveNext
            Loop
            cboPaisReside.ListIndex = IndiceListaCombo(cboPaisReside, oPersona.PaisReside)
        'End If
    End If
    '** End Juez *****************************************
    
    'Carga Datos Generales
    txtPersNacCreac.Text = Format(oPersona.FechaNacimiento, "dd/mm/yyyy")
    txtFecUltAct.Text = Format(oPersona.FechaActualizacion, "dd/mm/yyyy")
    
    '*** PEAC 20080412
    txtPersFecInscRuc.Text = Format(oPersona.FechaInscRuc, "dd/mm/yyyy")
    txtPersFecIniActi.Text = Format(oPersona.FechaIniActi, "dd/mm/yyyy")
    
    txtPersTelefono.Text = oPersona.Telefonos
    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, oPersona.CIIU)
    
    'EJVG20111209 **********************************
    txtCel1.Text = oPersona.Celular
    txtCel2.Text = oPersona.Celular2
    txtCel3.Text = oPersona.Celular3
    TxtEmail2.Text = oPersona.Email2
    '**********************************************
    '*** PEAC 20080412
    
    'MAVM 03042009 Cargar la Fecha de Fallecimiento
    'txtPersFallec.Text = Format(oPersona.FechaFallecimiento, "dd/mm/yyyy")
    txtPersFallec.Text = Format(IIf(oPersona.FechaFallecimiento = "01/01/1900", "__/__/____", oPersona.FechaFallecimiento), "dd/mm/yyyy")
    'MAVM 03042009
    
    'MADM 20091113
    'txtcargo.Text = oPersona.Cargo
    'txtcentro.Text = oPersona.CentroTrabajo
    chkcred.value = oPersona.RCredito
    chkaho.value = oPersona.RAhorro
    chkser.value = oPersona.RServicio
    chkotro.value = oPersona.ROtro

    Me.cboocupa.ListIndex = IndiceListaCombo(cboocupa, oPersona.RcActiGiro1) 'madm 20100503
    'Me.chkpeps.value = oPersona.RPeps 'madm 20100408
    'EJVG20120813 ***
    Me.cboPEPS.ListIndex = IndiceListaCombo(Me.cboPEPS, oPersona.RPeps)
    Me.cboSujetoObligado.ListIndex = IndiceListaCombo(Me.cboSujetoObligado, oPersona.SujetoObligado)
    Me.cboOfCumplimiento.ListIndex = IndiceListaCombo(Me.cboOfCumplimiento, oPersona.OfCumplimiento)
    'END EJVG *******
    'FRHU 20151130 ERS077-2015
    If oPersona.AutorizaUsoDatos = -1 Then
        lblAutorizarUsoDatos.Visible = True   'modificado pti1 se cambio a true ERS070-2018 26/12/2018
        CboAutoriazaUsoDatos.Visible = True  'modificado pti1 se cambio a true ERS070-2018 26/12/2018
    Else
        Me.CboAutoriazaUsoDatos.ListIndex = IndiceListaCombo(Me.CboAutoriazaUsoDatos, oPersona.AutorizaUsoDatos)
        lblAutorizarUsoDatos.Visible = True
        CboAutoriazaUsoDatos.Visible = True
        CboAutoriazaUsoDatos.Enabled = False
    End If
    'FIN FRHU 20151130
    'end madm
        
    Me.cboCargos.ListIndex = IndiceListaCombo(cboCargos, oPersona.RnCargo) '** Juez 20120326
        
    cboTipoComp.ListIndex = IndiceListaCombo(cboTipoComp, oPersona.TIPOCOMPDescripcion)
    cboTipoSistInfor.ListIndex = IndiceListaCombo(cboTipoSistInfor, oPersona.TIPOSISTINFORDescripcion)
    cboCadenaProd.ListIndex = IndiceListaCombo(cboCadenaProd, oPersona.CADENAPRODDescripcion)
    
    cboMonePatri.ListIndex = IndiceListaCombo(cboMonePatri, oPersona.MonedaPatri)
    '*** FIN PEAC
    
    '*** PEAC 20080211
    Me.TxtCodCIIU.Text = Int(val(Trim(Right(CboPersCiiu.Text, 4))))
    
    Call CargaControlEstadoPersona(oPersona.Personeria)
    cmbPersEstado.ListIndex = IndiceListaCombo(cmbPersEstado, Right("0" & oPersona.Estado, 2))
    
    TxtSiglas.Text = oPersona.Siglas     'Carga Razon Social
    
    TxtSbs.Text = oPersona.PersCodSbs    'Carga Codigo SBS
    
    'Selecciona el Tipo de Persona Juridica
    cmbPersJurTpo.ListIndex = IndiceListaCombo(cmbPersJurTpo, Trim(str(IIf(oPersona.TipoPersonaJur = "", -1, oPersona.TipoPersonaJur))))

    'Selecciona la relacion Con la Persona
    CmbRela.ListIndex = IndiceListaCombo(CmbRela, oPersona.PersRelInst)

    'Selecciona la magnitud Empresarial
    cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, Trim(oPersona.MagnitudEmpresarial))
    cmbPersNatMagnitud.ListIndex = IndiceListaCombo(cmbPersNatMagnitud, Trim(oPersona.MagnitudPersNat)) 'JACA 20110427
    lblMagnitudEmpresarial.Caption = cmbPersJurMagnitud.Text
    'Carga Numero de Empleados
    txtPersJurEmpleados.Text = Trim(str(oPersona.NumerosEmpleados))
    'Carga Objeto Social
    txtPersJurObjSocial.Text = oPersona.ObjetoSocial '** Juez 20120328
    
    
    'CUSCO
    txtIngresoProm.Text = Format(oPersona.IngresoPromedio, "0.00")
    
    Call CargaDocumentos                    'Carga Los Documentos de la Personas
    
    Call CargaRelacionesPersonas            'Carga las Relaciones de las Personas
    
    Call CargaVisitasPersonas               'Carga las Relaciones de las Personas
    
    Call CargaVentasPersonas                'Cargar Ventas
    
    Call CargaFuentesIngreso                'Carga las Fuentes de Ingresos de las Personas
    
    Call CargaRefComerciales                'Carga las Referencias Comerciales
    lnNumRefCom = oPersona.MaxRefComercial  'Carga el max Ref Comercial
     
    Call CargaRefBancarias                  'Carga las Referencias Bancarias
    
    Call CargaPatVehicular                  'Carga el Patrimonio Vehicular
    lnNumPatVeh = oPersona.MaxPatVehicular  'Carga el max Pat Vehicular
    
    '*** PEAC 20080412
    Call CargaPatOtros                  'Carga el Patrimonio oTROS
    lnNumPatOtros = oPersona.MaxPatOtros  'Carga el max Pat Vehicular
    
    '*** PEAC 20080412
    Call CargaPatInmuebles                  'Carga el Patrimonio inmuebles
    lnNumPatInmuebles = oPersona.MaxPatInmuebles  'Carga el max Pat Inmuebles
    
    '*** CTI3 10092018
    Call CargaDatosAdicionales
    
    'Carga Firma
    'Call IDBFirma.CargarFirma(oPersona.RFirma)
    
    If Not oPersona Is Nothing Then
    If oPersona.Personeria = gPersonaNat Then
        SSTabs.Tab = 0
    Else
        SSTabs.Tab = 1
    End If
    End If
    bEstadoCargando = False
    CmdPersFteConsultar.Enabled = True
    'WIOR 20130827 *******************************
    fsNombreActual = oPersona.NombreCompleto
    Set rsDocPersActual = FEDocs.GetRsNew(0)
    'WIOR FIN ************************************
    'FRHU 20151130 ERS077-2015
    If oPersona.Personeria = gPersonaNat Then
        MatPersona(1).sNombres = oPersona.Nombres
        MatPersona(1).sApePat = oPersona.ApellidoPaterno
        MatPersona(1).sApeMat = oPersona.Nombres
        MatPersona(1).sApeCas = oPersona.ApellidoCasada
        MatPersona(1).sSexo = oPersona.Sexo
        MatPersona(1).sEstadoCivil = oPersona.EstadoCivil
        MatPersona(1).cNacionalidad = oPersona.Nacionalidad
        MatPersona(1).sDomicilio = oPersona.Domicilio
        MatPersona(1).sRefDomicilio = oPersona.RefDomicilio
        MatPersona(1).sUbicGeografica = oPersona.UbicacionGeografica
        MatPersona(1).sCelular = oPersona.Celular
        MatPersona(1).sTelefonos = oPersona.Telefonos
        MatPersona(1).sEmail = oPersona.Email
        Call CargaDocumentosParaActAutDatosIni
    End If
    'FIN FRHU 20151130
End Sub

'Private Sub CargaUbicacionesGeograficas()
''Dim Conn As COMConecta.DCOMConecta   'DConecta
'Dim sSql As String
'Dim R As ADODB.Recordset
'Dim i As Integer
'Dim oPersonas As New COMDPersona.DCOMPersonas
'
'On Error GoTo ErrCargaUbicacionesGeograficas
''    Set Conn = New COMConecta.DCOMConecta 'DConecta
'    'Carga Niveles
''    sSql = "select *, 1 p from UbicacionGeografica where cUbiGeoCod like '0%' "
''    sSql = sSql & " Union "
''    sSql = sSql & " Select *, 2 p from UbicacionGeografica where cUbiGeoCod like '1%'"
''    sSql = sSql & " Union "
''    sSql = sSql & " select *, 3 p from UbicacionGeografica where cUbiGeoCod like '2%' "
''    sSql = sSql & " Union "
''    sSql = sSql & " select *, 4 p from UbicacionGeografica where cUbiGeoCod like '3%' "
''    sSql = sSql & " Union "
''    sSql = sSql & " select *, 5 p from UbicacionGeografica where cUbiGeoCod like '4%' order by p,cUbiGeoDescripcion "
'    ContNiv1 = 0
'    ContNiv2 = 0
'    ContNiv3 = 0
'    ContNiv4 = 0
'    ContNiv5 = 0
'
'    'Conn.AbreConexion
'    'Set R = Conn.CargaRecordSet(sSql)
'    Set R = oPersonas.CargarUbicacionesGeograficas
'    Do While Not R.EOF
'        Select Case R!P
'            Case 1 'Pais
'                ContNiv1 = ContNiv1 + 1
'                ReDim Preserve Nivel1(ContNiv1)
'                Nivel1(ContNiv1 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 2 ' Departamento
'                ContNiv2 = ContNiv2 + 1
'                ReDim Preserve Nivel2(ContNiv2)
'                Nivel2(ContNiv2 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 3 'Provincia
'                ContNiv3 = ContNiv3 + 1
'                ReDim Preserve Nivel3(ContNiv3)
'                Nivel3(ContNiv3 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 4 'Distrito
'                ContNiv4 = ContNiv4 + 1
'                ReDim Preserve Nivel4(ContNiv4)
'                Nivel4(ContNiv4 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'            Case 5 'Zona
'                ContNiv5 = ContNiv5 + 1
'                ReDim Preserve Nivel5(ContNiv5)
'                Nivel5(ContNiv5 - 1) = Trim(R!cUbiGeoDescripcion) & Space(50) & Trim(R!cUbiGeoCod)
'        End Select
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    'Conn.CierraConexion
'    'Set Conn = Nothing
'
'    'Carga el Nivel1 en el Control
'    cmbPersUbiGeo(0).Clear
'    For i = 0 To ContNiv1 - 1
'        cmbPersUbiGeo(0).AddItem Nivel1(i)
'        cmbNacionalidad.AddItem Nivel1(i)
'        If Trim(Right(Nivel1(i), 10)) = "04028" Then
'            nPos = i
'        End If
'    Next i
'    cmbPersUbiGeo(0).ListIndex = nPos
'    cmbPersUbiGeo(2).Clear
'    cmbPersUbiGeo(3).Clear
'    cmbPersUbiGeo(4).Clear
'    cmbNacionalidad.ListIndex = nPos
'    Exit Sub
'
'ErrCargaUbicacionesGeograficas:
'    MsgBox Err.Description, vbInformation, "Aviso"
'
'End Sub
Private Sub ActualizaCombo(ByVal psValor As String, ByVal TipoCombo As TTipoCombo)
Dim i As Long
Dim sCodigo As String
    
    sCodigo = Trim(Right(psValor, 15))
    Select Case TipoCombo
        Case ComboDpto
            cmbPersUbiGeo(1).Clear
            If sCodigo = "04028" Then
                For i = 0 To ContNiv2 - 1
                    cmbPersUbiGeo(1).AddItem Nivel2(i)
                Next i
            Else
                cmbPersUbiGeo(1).AddItem psValor
            End If
        Case ComboProv
            cmbPersUbiGeo(2).Clear
            If Len(sCodigo) > 3 Then
                cmbPersUbiGeo(2).Clear
                For i = 0 To ContNiv3 - 1
                    If Mid(sCodigo, 2, 2) = Mid(Trim(Right(Nivel3(i), 15)), 2, 2) Then
                        cmbPersUbiGeo(2).AddItem Nivel3(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(2).AddItem psValor
            End If
        Case ComboDist
            cmbPersUbiGeo(3).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv4 - 1
                    If Mid(sCodigo, 2, 4) = Mid(Trim(Right(Nivel4(i), 15)), 2, 4) Then
                        cmbPersUbiGeo(3).AddItem Nivel4(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(3).AddItem psValor
            End If
        Case ComboZona
            cmbPersUbiGeo(4).Clear
            If Len(sCodigo) > 3 Then
                For i = 0 To ContNiv5 - 1
                    If Mid(sCodigo, 2, 6) = Mid(Trim(Right(Nivel5(i), 15)), 2, 6) Then
                        cmbPersUbiGeo(4).AddItem Nivel5(i)
                    End If
                Next i
            Else
                cmbPersUbiGeo(4).AddItem psValor
            End If
    End Select
End Sub
Private Sub cboCadenaProd_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        cboTipoSistInfor.SetFocus
  End If
End Sub
Private Sub cboMonePatri_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.MonedaPatri = Trim(Right(cboMonePatri.Text, 10))
  End If
End Sub

Private Sub cboMonePatri_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.MonedaPatri = Trim(Right(cboMonePatri.Text, 10))
            
    End If

End Sub

Private Sub cboMotivoActu_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        'oPersona.Nacionalidad = Trim(Right(cboMotivoActu.Text, 12))
        oPersona.MotivoActu = Trim(Right(cboMotivoActu.Text, 12))
    End If
    'JUEZ 20131007 Campaña Actualiza Datos *************
    If oPersona.MotivoActu = "2" Then
        If bValidaCampDatos Then
            If Not ValidaDetallesCampDatos Then Exit Sub
        End If
        lblBUsuario.Visible = True
        txtBUsuario.Visible = True
        'txtActiGiro.BackColor = &HC0FFFF
    Else
        lblBUsuario.Visible = False
        txtBUsuario.Visible = False
        txtBUsuario.Text = ""
        'txtActiGiro.BackColor = &H80000005
    End If
    'END JUEZ ******************************************
End Sub

Private Sub cboMotivoActu_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        'oPersona.Nacionalidad = Trim(Right(cboMotivoActu.Text, 12))
        oPersona.MotivoActu = Trim(Right(cboMotivoActu.Text, 12))
    End If
    'JUEZ 20131007 Campaña Actualiza Datos *************
    If Trim(Right(cboMotivoActu.Text, 12)) = "2" Then
        If bValidaCampDatos Then
            If Not ValidaDetallesCampDatos Then Exit Sub
        End If
        lblBUsuario.Visible = True
        txtBUsuario.Visible = True
        txtActiGiro.BackColor = &HC0FFFF
    Else
        lblBUsuario.Visible = False
        txtBUsuario.Visible = False
        txtBUsuario.Text = ""
        txtActiGiro.BackColor = &HC0FFFF
    End If
    'END JUEZ ******************************************
End Sub
'madm 20100320
Private Sub cboocupa_Change()
If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
            
            oPersona.RcActiGiro1 = Trim(Right(Me.cboocupa.Text, 10))
        
    End If
End Sub

Private Sub cboocupa_Click()
If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
            oPersona.RcActiGiro1 = Trim(Right(Me.cboocupa.Text, 10))
        
    End If
End Sub

Private Sub cboocupa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
         If Trim(Right(Me.cboocupa.Text, 10)) = "116" Then
            Me.txtActiGiro.SetFocus
         Else
            Me.cboCargos.SetFocus '** Juez 20120326
         End If
    End If
End Sub
'end madm 20100320

'** Juez 20120326 ************************************
Private Sub cboCargos_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
    oPersona.RnCargo = Trim(Right(Me.cboCargos.Text, 10))
    End If
End Sub

Private Sub cboCargos_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
    oPersona.RnCargo = Trim(Right(Me.cboCargos.Text, 10))
    End If
End Sub

Private Sub cboCargos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbPersEstado.SetFocus
    End If
End Sub
'**End Juez ******************************************

Private Sub CboPersCiiu_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 20))
        
        '***PEAC 20080211
        Me.TxtCodCIIU.Text = Int(val(Trim(Right(CboPersCiiu.Text, 4))))
        
    End If
End Sub

'JUEZ 20131007 *********************************************************
Private Sub cboRemInfoEmail_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RemisionInfoEmail = CInt(Trim(Right(cboRemInfoEmail.Text, 2)))
        If oPersona.RemisionInfoEmail = 1 Then
            TxtEmail.BackColor = &HC0FFFF
        Else
            TxtEmail.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub cboRemInfoEmail_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RemisionInfoEmail = CInt(Trim(Right(cboRemInfoEmail.Text, 2)))
        If oPersona.RemisionInfoEmail = 1 Then
            TxtEmail.BackColor = &HC0FFFF
        Else
            TxtEmail.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub cboRemInfoEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtEmail.SetFocus
    End If
End Sub
'END JUEZ **************************************************************

'*** PEAC 20080412
Private Sub CboTipoComp_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TIPOCOMPDescripcion = Trim(Right(cboTipoComp.Text, 20))
        
    End If
End Sub
'MADM 20091114
Private Sub cboTipoComp_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        CmbRela.SetFocus
    End If
End Sub
'END MADM


Private Sub cboTipoPatri_Click()
If Not Me.cmdPatVehAcepta.Visible Then
    If Len(cboTipoPatri) > 0 Then
        If CInt(Trim(Right(cboTipoPatri, 5))) = 1 Then
            Me.fePatInmuebles.Visible = True
            Me.fePatVehicular.Visible = False
            Me.fePatOtros.Visible = False
        ElseIf CInt(Trim(Right(cboTipoPatri, 5))) = 2 Then
            Me.fePatVehicular.Visible = True
            Me.fePatOtros.Visible = False
            Me.fePatInmuebles.Visible = False
        Else
            Me.fePatOtros.Visible = True
            Me.fePatInmuebles.Visible = False
            Me.fePatVehicular.Visible = False
        End If
    End If
    'lcTexto = ""
Else
    MsgBox "Debe validar los datos ingresados presionando <Aceptar>.", vbInformation, "Mensaje"
    'cboTipoPatri.SetFocus
    'cboTipoPatri.Text = lcTexto
End If
End Sub

Private Sub cboTipoPatri_GotFocus()
'lcTexto = cboTipoPatri.Text
End Sub

'*** PEAC 20080412
Private Sub CboTipoSistInfor_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TIPOSISTINFORDescripcion = Trim(Right(cboTipoSistInfor.Text, 20))
        
    End If
End Sub
'*** PEAC 20080412
Private Sub CboCadenaProd_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CADENAPRODDescripcion = Trim(Right(cboCadenaProd.Text, 20))
        
    End If
End Sub


Private Sub CboPersCiiu_Click()
       
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 10))
        
        '***PEAC 20080211
        Me.TxtCodCIIU.Text = Int(val(Trim(Right(CboPersCiiu.Text, 4))))
    
    End If
End Sub

'*** PEAC 20080412
Private Sub Cbotipocomp_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TIPOCOMPDescripcion = Trim(Right(cboTipoComp.Text, 10))
        
    End If
End Sub
'*** PEAC 20080412
Private Sub CboTipoSistInfor_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TIPOSISTINFORDescripcion = Trim(Right(cboTipoSistInfor.Text, 10))
            
    End If
End Sub
'*** PEAC 20080412
Private Sub CboCadenaProd_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CADENAPRODDescripcion = Trim(Right(cboCadenaProd.Text, 10))
        
    End If
End Sub
Private Sub CboPersCiiu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cboocupa.Enabled = True Then
            Me.cboocupa.SetFocus
        Else
            Me.cmbPersEstado.SetFocus
        End If
    End If
End Sub

Private Sub CboTipoSangre_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoSangre = Trim(Right(CboTipoSangre.Text, 2))
    End If
End Sub

Private Sub CboTipoSangre_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoSangre = Trim(Right(CboTipoSangre.Text, 2))
    End If
End Sub

Private Sub CboTipoSangre_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            If (TxtPeso.Enabled) Then
                TxtPeso.SetFocus
            End If
        End If
End Sub
'MADM 20091114
Private Sub cboTipoSistInfor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SSTDatosGen.Tab = 1
        cmbPersUbiGeo(0).SetFocus
  End If
End Sub

Private Sub chkaho_Click()
   If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RAhorro = chkaho.value
    End If
End Sub

Private Sub chkaho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RAhorro = chkaho.value
    End If
End If
End Sub

Private Sub chkcred_Click()
 If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RCredito = chkcred.value
    End If
End Sub

Private Sub chkcred_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RCredito = chkcred.value
    End If
End If
End Sub

Private Sub chkotro_Click()
 If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ROtro = chkotro.value
    End If
End Sub

Private Sub chkotro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ROtro = chkotro.value
    End If
End If
End Sub
'EJVG20120813 ***
'Private Sub chkpeps_Click()
'    If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
'        oPersona.RPeps = chkpeps.value
'    End If
'End Sub
'
'Private Sub chkpeps_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'        If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
'        oPersona.RPeps = chkpeps.value
'    End If
'End If
'
'End Sub
Private Sub cboPEPS_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RPeps = CInt(Right(Me.cboPEPS.Text, 1))
    End If
End Sub
Private Sub cboPEPS_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RPeps = CInt(Right(Me.cboPEPS.Text, 1))
    End If
End Sub
Private Sub cboPEPS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            If Not IsNumeric(Right(Me.cboPEPS.Text, 1)) Then
                Exit Sub
            End If
            oPersona.RPeps = CInt(Right(Me.cboPEPS.Text, 1))
            If Me.cboSujetoObligado.Visible = True And Me.cboSujetoObligado.Enabled = True Then
                Me.cboSujetoObligado.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cboSujetoObligado_Click()
    Dim bSujetoObligadoDJ As Boolean
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.SujetoObligado = CInt(Right(Me.cboSujetoObligado.Text, 1))
        bSujetoObligadoDJ = IIf(oPersona.SujetoObligado = 1, True, False)
        Me.lblOfCumplimiento.Visible = bSujetoObligadoDJ
        Me.cboOfCumplimiento.Visible = bSujetoObligadoDJ
    End If
End Sub
Private Sub cboSujetoObligado_Change()
    Dim bSujetoObligadoDJ As Boolean
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.SujetoObligado = CInt(Right(Me.cboSujetoObligado.Text, 1))
        bSujetoObligadoDJ = IIf(oPersona.SujetoObligado = 1, True, False)
        Me.lblOfCumplimiento.Visible = bSujetoObligadoDJ
        Me.cboOfCumplimiento.Visible = bSujetoObligadoDJ
    End If
End Sub
Private Sub cboSujetoObligado_KeyPress(KeyAscii As Integer)
    Dim bSujetoObligadoDJ As Boolean
    If KeyAscii = 13 Then
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            If Not IsNumeric(Right(Me.cboSujetoObligado.Text, 1)) Then
                Exit Sub
            End If
            oPersona.SujetoObligado = CInt(Right(Me.cboSujetoObligado.Text, 1))
            bSujetoObligadoDJ = IIf(oPersona.SujetoObligado = 1, True, False)
            Me.lblOfCumplimiento.Visible = bSujetoObligadoDJ
            Me.cboOfCumplimiento.Visible = bSujetoObligadoDJ
            If bSujetoObligadoDJ Then
                If Me.cboOfCumplimiento.Visible And Me.cboOfCumplimiento.Enabled Then Me.cboOfCumplimiento.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cboOfCumplimiento_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.OfCumplimiento = CInt(Right(Me.cboOfCumplimiento.Text, 1))
    End If
End Sub
Private Sub cboOfCumplimiento_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.OfCumplimiento = CInt(Right(Me.cboOfCumplimiento.Text, 1))
    End If
End Sub
Private Sub cboOfCumplimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            If Not IsNumeric(Right(Me.cboOfCumplimiento.Text, 1)) Then
                Exit Sub
            End If
            oPersona.OfCumplimiento = CInt(Right(Me.cboOfCumplimiento.Text, 1))
        End If
    End If
End Sub
'END EJVG *******
'END MADM
'FRHU 20151204 ERS077-2015
Private Sub cboAutoriazaUsoDatos_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.AutorizaUsoDatos = CInt(Right(Me.CboAutoriazaUsoDatos.Text, 1))
    End If
End Sub
Private Sub cboAutoriazaUsoDatos_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.AutorizaUsoDatos = CInt(Right(Me.CboAutoriazaUsoDatos.Text, 1))
    End If
End Sub
Private Sub cboAutoriazaUsoDatos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            If Not IsNumeric(Right(Me.CboAutoriazaUsoDatos.Text, 1)) Then
                Exit Sub
            End If
            oPersona.AutorizaUsoDatos = CInt(Right(Me.CboAutoriazaUsoDatos.Text, 1))
        End If
    End If
End Sub
'FIN FRHU 20151204
'** Comentado por Juez 20120327 *****************************
'Private Sub chkResidente_Click()
'    'On Error Resume Next
'    If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
'        oPersona.Residencia = chkResidente.value
'    End If
'End Sub
'
'Private Sub chkResidente_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
''    On Error Resume Next
'        If Not bEstadoCargando Then
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'        End If
'        oPersona.Residencia = chkResidente.value
'    End If
'    'If TxtTalla.Enabled Then TxtTalla.SetFocus
'End If
'End Sub

'** Juez 20120327 ************************************
Private Sub cboResidente_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Residencia = Trim(Right(Me.cboResidente.Text, 1))
        
        If Trim(Right(Me.cboResidente.Text, 1)) = 0 Then
            Dim oPersonas As New COMDPersona.DCOMPersonas
            Set oPersonas = New COMDPersona.DCOMPersonas
            Dim prsUbiGeo As ADODB.Recordset

            'Call oPersonas.CargarDatosObjetosPersona(lrsRelac, lrsDocumentos, lrsRefComercial, lrsPatVehicular, lrsTipoSangre, lrsDireccCondic, lrsPersoneria, lrsMagnitudComer, lrsEstCivil, lrsRelacInst, _
                                        lrsUbiGeo, lrsCIIU, gsCodCMAC, lrsTipoPerJuri, lrsVisita, lrsTIPOCOMP, lrsTIPOSISTINFOR, lrsCADENAPROD, lrsTipoPatri, lrsMonePatri, lrsAlterSiNo, lrsMotivoActu, lrsOcupa, lrsCargo)
            Set prsUbiGeo = oPersonas.CargarUbicacionesGeograficas(True, 0)
            Me.cboPaisReside.Enabled = True
            'Me.lblResidente.Enabled = True
            Me.cboPaisReside.SetFocus
            Do While Not prsUbiGeo.EOF
                If Trim(prsUbiGeo!cUbiGeoCod) <> "04028" Then 'JUEZ 20131007 Para que Perú no sea opción a elegir
                    cboPaisReside.AddItem Trim(prsUbiGeo!cUbiGeoDescripcion) & Space(100) & Trim(prsUbiGeo!cUbiGeoCod)
                End If
                prsUbiGeo.MoveNext
            Loop
        Else
            Me.cboPaisReside.Enabled = False
            'Me.lblResidente.Enabled = False
            cboPaisReside.Clear
            Me.TxtTalla.SetFocus
            oPersona.PaisReside = ""
            'cboPaisReside.AddItem "" 'UltimaModificacion
        End If
    End If
End Sub

Private Sub cboResidente_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Residencia = Trim(Right(Me.cboResidente.Text, 1))
        
        If Trim(Right(Me.cboResidente.Text, 1)) = 0 Then
            Dim oPersonas As New COMDPersona.DCOMPersonas
            Set oPersonas = New COMDPersona.DCOMPersonas
            Dim prsUbiGeo As ADODB.Recordset

            'Call oPersonas.CargarDatosObjetosPersona(lrsRelac, lrsDocumentos, lrsRefComercial, lrsPatVehicular, lrsTipoSangre, lrsDireccCondic, lrsPersoneria, lrsMagnitudComer, lrsEstCivil, lrsRelacInst, _
                                        lrsUbiGeo, lrsCIIU, gsCodCMAC, lrsTipoPerJuri, lrsVisita, lrsTIPOCOMP, lrsTIPOSISTINFOR, lrsCADENAPROD, lrsTipoPatri, lrsMonePatri, lrsAlterSiNo, lrsMotivoActu, lrsOcupa, lrsCargo)
            Set prsUbiGeo = oPersonas.CargarUbicacionesGeograficas(True, 0)
            Me.cboPaisReside.Enabled = True
            'Me.lblResidente.Enabled = True
            Me.cboPaisReside.SetFocus
            Do While Not prsUbiGeo.EOF
                If Trim(prsUbiGeo!cUbiGeoCod) <> "04028" Then 'JUEZ 20131007 Para que Perú no sea opción a elegir
                    cboPaisReside.AddItem Trim(prsUbiGeo!cUbiGeoDescripcion) & Space(100) & Trim(prsUbiGeo!cUbiGeoCod)
                End If
                prsUbiGeo.MoveNext
            Loop
        Else
            Me.cboPaisReside.Enabled = False
            'Me.lblResidente.Enabled = False
            cboPaisReside.Clear
            Me.TxtTalla.SetFocus
            oPersona.PaisReside = ""
        End If
    End If
End Sub

Private Sub cboResidente_KeyPress(KeyAscii As Integer)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Residencia = Trim(Right(Me.cboResidente.Text, 1))
        
        If Trim(Right(Me.cboResidente.Text, 1)) = 0 Then
            Dim oPersonas As New COMDPersona.DCOMPersonas
            Set oPersonas = New COMDPersona.DCOMPersonas
            Dim prsUbiGeo As ADODB.Recordset

            'Call oPersonas.CargarDatosObjetosPersona(lrsRelac, lrsDocumentos, lrsRefComercial, lrsPatVehicular, lrsTipoSangre, lrsDireccCondic, lrsPersoneria, lrsMagnitudComer, lrsEstCivil, lrsRelacInst, _
                                        lrsUbiGeo, lrsCIIU, gsCodCMAC, lrsTipoPerJuri, lrsVisita, lrsTIPOCOMP, lrsTIPOSISTINFOR, lrsCADENAPROD, lrsTipoPatri, lrsMonePatri, lrsAlterSiNo, lrsMotivoActu, lrsOcupa, lrsCargo)
            Set prsUbiGeo = oPersonas.CargarUbicacionesGeograficas(True, 0)
            Me.cboPaisReside.Enabled = True
            'Me.lblResidente.Enabled = True
            Me.cboPaisReside.SetFocus
            Do While Not prsUbiGeo.EOF
                If Trim(prsUbiGeo!cUbiGeoCod) <> "04028" Then 'JUEZ 20131007 Para que Perú no sea opción a elegir
                    cboPaisReside.AddItem Trim(prsUbiGeo!cUbiGeoDescripcion) & Space(100) & Trim(prsUbiGeo!cUbiGeoCod)
                End If
                prsUbiGeo.MoveNext
            Loop
        Else
            Me.cboPaisReside.Enabled = False
            'Me.lblResidente.Enabled = False
            cboPaisReside.Clear
            Me.TxtTalla.SetFocus
            oPersona.PaisReside = ""
        End If
    End If
End Sub

Private Sub cboPaisReside_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PaisReside = Trim(Right(cboPaisReside.Text, 12))
    End If
End Sub

Private Sub cboPaisReside_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PaisReside = Trim(Right(cboPaisReside.Text, 12))
    End If
End Sub

Private Sub cboPaisReside_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not bEstadoCargando Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
            oPersona.PaisReside = Trim(Right(cboPaisReside.Text, 12))
        End If
        Me.TxtTalla.SetFocus
    End If
End Sub
'** End Juez *****************************************

'MADM 20091114
Private Sub chkser_Click()
 If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RServicio = chkser.value
    End If
End Sub

Private Sub chkser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RServicio = chkser.value
    End If
End If
End Sub

'END MADM
Private Sub cmbNacionalidad_Change()

    'On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
    End If
    
End Sub

Private Sub cmbNacionalidad_Click()
    'On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
        If Trim(Right(cmbNacionalidad.Text, 12)) = "04028" Then
            'chkResidente.value = 1 'Comentado por Juez 20120327
            cboResidente.ListIndex = 0 '** Juez 20120327
        Else
            'chkResidente.value = 0 'Comentado por Juez 20120327
            cboResidente.ListIndex = 1 '** Juez 20120327
        End If
        cboResidente.SetFocus
        Call LlenarComboTipoDocumento
    End If
    
End Sub

Private Sub cmbNacionalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'chkResidente.SetFocus 'Comentado por Juez 20120327
    cboResidente.SetFocus '** Juez 20120327
End If
End Sub

'JUEZ 20131007 **************************************************************
Private Sub cmbNegUbiGeo_Change(Index As Integer)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbiGeoNegocio = Trim(Right(cmbNegUbiGeo(4).Text, 15))
    End If
End Sub

Private Sub cmbNegUbiGeo_Click(Index As Integer)
Dim oUbic As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim i As Integer

If Index <> 4 Then
    Set oUbic = New COMDPersona.DCOMPersonas
    Set rs = oUbic.CargarUbicacionesGeograficas(True, Index + 1, Trim(Right(cmbNegUbiGeo(Index).Text, 15)))

    If Trim(Right(cmbNegUbiGeo(0).Text, 12)) <> "04028" Then
         If Index = 0 Then
            For i = 1 To cmbNegUbiGeo.count - 1
                cmbNegUbiGeo(i).Clear
                cmbNegUbiGeo(i).AddItem Trim(Trim(cmbNegUbiGeo(0).Text)) & Space(50) & Trim(Right(cmbNegUbiGeo(0).Text, 12))
            Next i
         End If
    Else
        For i = Index + 1 To cmbNegUbiGeo.count - 1
        cmbNegUbiGeo(i).Clear
        Next
        
        While Not rs.EOF
            cmbNegUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
            rs.MoveNext
        Wend
    End If
    Set oUbic = Nothing
End If

If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.UbiGeoNegocio = Trim(Right(cmbNegUbiGeo(4).Text, 15))
End If
End Sub

Private Sub cmbNegUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            cmbNegUbiGeo(Index + 1).SetFocus
        Else
            txtNegDireccion.SetFocus
        End If
    End If
End Sub
'END JUEZ *******************************************************************

Private Sub cmbPersDireccCondicion_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CondicionDomicilio = Trim(Right(cmbPersDireccCondicion.Text, 10))
    End If
End Sub

Private Sub cmbPersDireccCondicion_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.CondicionDomicilio = Trim(Right(cmbPersDireccCondicion.Text, 10))
    End If
End Sub

Private Sub cmbPersDireccCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtValComercial.SetFocus
    End If
End Sub

Private Sub cmbPersEstado_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Estado = Trim(Right(cmbPersEstado.Text, 10))
      End If
End Sub

Private Sub cmbPersEstado_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Estado = Trim(Right(cmbPersEstado.Text, 10))

    End If
End Sub

Private Sub cmbPersEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Me.chkcred.Enabled = True Then
            Me.chkcred.SetFocus
       Else
            cboTipoComp.SetFocus
       End If
    End If
End Sub

Private Sub cmbPersJurMagnitud_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
            oPersona.MagnitudEmpresarial = Trim(Right(cmbPersJurMagnitud.Text, 15))
        End If
    End If
End Sub

Private Sub cmbPersJurMagnitud_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.MagnitudEmpresarial = Trim(Right(cmbPersJurMagnitud.Text, 15))
    End If
End Sub

Private Sub cmbPersJurMagnitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNacCreac.SetFocus
    End If
End Sub

Private Sub cmbPersJurTpo_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersJurTpo.Text, 10))
    End If
End Sub

Private Sub cmbPersJurTpo_Click()
    
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersJurTpo.Text, 10))
    End If
End Sub

Private Sub cmbPersJurTpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersJurEmpleados.SetFocus
    End If
End Sub



Private Sub cmdAdicJuridica_Click()

  frmPersonaJurDatosAdic.SSTabs.TabVisible(0) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(1) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(2) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(3) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(4) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(5) = True
  frmPersonaJurDatosAdic.lblPersElejida.Caption = 2
  frmPersonaJurDatosAdic.Show 1
  
End Sub

Private Sub cmdPerNatDatAdc_Click()
  frmPersonaJurDatosAdic.SSTabs.TabVisible(0) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(1) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(2) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(3) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(4) = True
  frmPersonaJurDatosAdic.SSTabs.TabVisible(5) = True
  frmPersonaJurDatosAdic.lblPersElejida.Caption = 1
  frmPersonaJurDatosAdic.Show 1
 
End Sub

Private Sub txtBUsuario_EmiteDatos()
    'If txtBUsuario.Text = "DEX" Then txtBUsuario.Text = ""
End Sub

'RECO20140217 ERS160-2013******************************************************
Private Sub txtCel1_LostFocus()
    If Len(Trim(txtCel1.Text)) > 0 Then
        txtPersTelefono.BackColor = 16777215
        txtCel1.BackColor = 12648447
    Else
        txtPersTelefono.BackColor = 12648447
        txtCel1.BackColor = 16777215
    End If
    If Len(Trim(txtPersTelefono.Text)) > 0 And Len(Trim(txtCel1.Text)) > 0 Then
        txtPersTelefono.BackColor = 12648447
        txtCel1.BackColor = 12648447
    End If
End Sub
'RECO FIN**********************************************************************




'** Juez 20120328 *****************************************
Private Sub txtPersJurObjSocial_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ObjetoSocial = Trim(txtPersJurObjSocial.Text)
    End If
End Sub

Private Sub txtPersJurObjSocial_GotFocus()
    fEnfoque txtPersJurObjSocial
End Sub
Private Sub txtPersJurObjSocial_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Me.cmbPersJurTpo.SetFocus
    End If
End Sub
'** End Juez **********************************************

Private Sub cmbPersNatEstCiv_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
    
  End If
End Sub

Private Sub cmbPersNatEstCiv_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
        If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
            txtApellidoCasada.Text = ""
        End If
    End If
End Sub

Private Sub cmbPersNatEstCiv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPersNatHijos.SetFocus
    End If
End Sub

Private Sub cmbPersNatSexo_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
  End If
End Sub

Private Sub cmbPersNatSexo_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Sexo = Trim(Right(cmbPersNatSexo.Text, 10))
        If oPersona.EstadoCivil = "" Then
            Exit Sub
        End If
        If (CInt(oPersona.EstadoCivil) = gPersEstadoCivilCasado Or CInt(oPersona.EstadoCivil) = gPersEstadoCivilViudo) And oPersona.Sexo = "F" Then
            Call DistribuyeApellidos(True)
        Else
            Call DistribuyeApellidos(False)
            txtApellidoCasada.Text = ""
        End If
    End If
End Sub

Private Sub cmbPersNatSexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbPersNatEstCiv.SetFocus
    End If
End Sub

Private Sub cmbPersPersoneria_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.TipoPersonaJur = Trim(Right(cmbPersPersoneria.Text, 15))
    End If
End Sub

Private Sub cmbPersPersoneria_Click()
    
    If Not bEstadoCargando Then
        If oPersona.Personeria <> CInt(Trim(Right(cmbPersPersoneria.Text, 15))) Then
            oPersona.Personeria = Trim(Right(IIf(cmbPersPersoneria.Text = "", Trim(str(gPersonaNat)), cmbPersPersoneria.Text), 15))
            If oPersona.Personeria <> gPersonaNat Then
                Call HabilitaFichaPersonaJur(True)
                Call HabilitaFichaPersonaNat(False)
                SSTabs.TabVisible(1) = True
                'SSTabs.TabVisible(9) = True 'FRHU 20150311 ERS013-2015
                SSTabs.Tab = 1
                CboAutoriazaUsoDatos.Visible = False 'add pti1 ers070-2018 26/12/2018
                lblAutorizarUsoDatos.Visible = False 'add pti1 ers070-2018 26/12/2018
            Else
                Call HabilitaFichaPersonaJur(False)
                Call HabilitaFichaPersonaNat(True)
                SSTabs.TabVisible(0) = True
                SSTabs.Tab = 0
                CboAutoriazaUsoDatos.Visible = True 'add pti1 ers070-2018 26/12/2018
                lblAutorizarUsoDatos.Visible = True 'add pti1 ers070-2018 26/12/2018
            End If
            cmbPersEstado.ListIndex = -1
        End If
    End If
    
    'CUSCO
    If oPersona.Personeria <> gPersonaNat Then
        SSTabs.TabVisible(0) = False
    Else
        SSTabs.TabVisible(1) = False
        SSTabs.TabVisible(9) = False
    End If
    '''''''''''''''''''''''''
    
    Call CargaControlEstadoPersona(oPersona.Personeria)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Personeria = Trim(Right(IIf(cmbPersPersoneria.Text = "", Trim(str(gPersonaNat)), cmbPersPersoneria.Text), 15))
    End If
    LlenarComboTipoDocumento 'madm 20100903
End Sub

Private Sub cmbPersPersoneria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPersNombreAP.Enabled Then
            txtPersNombreAP.SetFocus
        Else
            txtPersNombreRS.SetFocus
        End If
    End If
End Sub

Private Sub cmbPersUbiGeo_Change(Index As Integer)
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
    End If
End Sub

Private Sub cmbPersUbiGeo_Click(Index As Integer)
'        Select Case Index
'            Case 0 'Combo Pais
'                Call ActualizaCombo(cmbPersUbiGeo(0).Text, ComboDpto)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(2).Clear
'                    cmbPersUbiGeo(3).Clear
'                    cmbPersUbiGeo(4).Clear
'                End If
'            Case 1 'Combo Dpto
'                Call ActualizaCombo(cmbPersUbiGeo(1).Text, ComboProv)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(3).Clear
'                    cmbPersUbiGeo(4).Clear
'                End If
'            Case 2 'Combo Provincia
'                Call ActualizaCombo(cmbPersUbiGeo(2).Text, ComboDist)
'                If Not bEstadoCargando Then
'                    cmbPersUbiGeo(4).Clear
'                End If
'            Case 3 'Combo Distrito
'                Call ActualizaCombo(cmbPersUbiGeo(3).Text, ComboZona)
'        End Select
        
Dim oUbic As COMDPersona.DCOMPersonas
Dim rs As ADODB.Recordset
Dim i As Integer

If Index <> 4 Then

    Set oUbic = New COMDPersona.DCOMPersonas

    Set rs = oUbic.CargarUbicacionesGeograficas(True, Index + 1, Trim(Right(cmbPersUbiGeo(Index).Text, 15)))

    If Trim(Right(cmbPersUbiGeo(0).Text, 12)) <> "04028" Then
     'MADM 20101228
         If Index = 0 Then
            For i = 1 To cmbPersUbiGeo.count - 1
                cmbPersUbiGeo(i).Clear
                cmbPersUbiGeo(i).AddItem Trim(Trim(cmbPersUbiGeo(0).Text)) & Space(50) & Trim(Right(cmbPersUbiGeo(0).Text, 12))
            Next i
         End If
    'END MADM
    Else
        For i = Index + 1 To cmbPersUbiGeo.count - 1
        cmbPersUbiGeo(i).Clear
        Next
        
        While Not rs.EOF
            cmbPersUbiGeo(Index + 1).AddItem Trim(rs!cUbiGeoDescripcion) & Space(50) & Trim(rs!cUbiGeoCod)
            rs.MoveNext
        Wend
    End If
    Set oUbic = Nothing
End If

If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.UbicacionGeografica = Trim(Right(cmbPersUbiGeo(4).Text, 15))
End If




End Sub

Private Sub cmbPersUbiGeo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index < 4 Then
            cmbPersUbiGeo(Index + 1).SetFocus
        Else
            txtPersDireccDomicilio.SetFocus
        End If
    End If
End Sub

Private Sub CmbRela_Change()
    Call CmbRela_Click
End Sub

Private Sub CmbRela_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PersRelInst = CInt(Trim(Right(CmbRela.Text, 10)))
    End If
End Sub

Private Sub CmbRela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboCadenaProd.SetFocus
         End If
End Sub

Private Sub CmdActFirma_Click()
'Dim sRuta As String
'    CdlgImg.nHwd = Me.hwnd
'    CdlgImg.Show
'    sRuta = CdlgImg.Ruta
'    If Len(Trim(sRuta)) > 0 Then
'        IDBFirma.RutaImagen = sRuta
'        If oPersona.TipoActualizacion <> PersFilaNueva Then
'            oPersona.TipoActualizacion = PersFilaModificada
'            Call IDBFirma.GrabarFirma(oPersona.RFirma, oPersona.PersCodigo, "")
'        Else
'            Call IDBFirma.GrabarFirma(oPersona.RFirma, oPersona.PersCodigo, "")
'        End If
'    End If
End Sub

Private Sub cmdActualizarFirma_Click()

    If TxtBCodPers.Text = "" Then
        MsgBox "Debe registrar a la persona primero", vbInformation, "Mensaje"
        Exit Sub
    End If
    Call frmPersonaFirma.Inicio(oPersona.PersCodigo, oPersona.sCodage, , , 1)
    
    'ande 20171011 para considerar la fecha de ultima actualizacion de firma
    If Not bEstadoCargando Then
        If gbFirmaActualizada Then
            If oPersona.TipoActualizacion <> PersFilaNueva Then
                oPersona.TipoActualizacion = PersFilaModificada
            End If
        End If
    End If
    'end ande
End Sub

Private Sub cmdEditar_Click()
    If oPersona Is Nothing Then
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPersona.PersCodigo = "" Then
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    If Me.TxtBCodPers = "" Then '** Juez 20120328
        MsgBox "No se Puede Editar la persona", vbInformation, "Aviso"
        Exit Sub
    End If
    CmdPersAceptar.Enabled = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Enabled = True
    CmdPersCancelar.Visible = True
    cmdEditar.Enabled = False
    cmdnuevo.Enabled = False
    
    Call HabilitaControlesPersona(True)
    bNuevaPersona = False
    bPersonaAct = True
'    cmdVentasEditar_Click

 'MADM 20101020
        If oPersona.Personeria = 1 And oPersona.PersRRHH <> 0 Then
             'SSTabs.TabVisible(2) = False
            FERelPers.lbEditarFlex = bHabilitarBoton 'Not False
            cmdPersRelacNew.Enabled = Not False
            cmdPersRelacEditar.Enabled = bHabilitarBoton 'Not False
            cmdPersRelacDel.Enabled = bHabilitarBoton 'Not False
            cmdPersRelacAceptar.Enabled = Not False
            cmdPersRelacCancelar.Enabled = Not False
        Else
            FERelPers.lbEditarFlex = False
            cmdPersRelacNew.Enabled = False
            cmdPersRelacEditar.Enabled = False
            cmdPersRelacDel.Enabled = False
            cmdPersRelacAceptar.Enabled = False
            cmdPersRelacCancelar.Enabled = False
        End If
    cmbPersPersoneria.Enabled = False
    'END MADM
 
    'RIRO20140128 - Comentado
    ''Add By GITU 2011-04-26 Valida para ver o actualizar la firma de acuerdo al cargo
    'If (gsCodCargo <> "002001" And gsCodCargo <> "002002" And gsCodCargo <> "002003" And gsCodCargo <> "003001" And gsCodCargo <> "003002" _
    '    And gsCodCargo <> "004001" And gsCodCargo <> "004002" And gsCodCargo <> "006024" And gsCodCargo <> "007012" And gsCodCargo <> "006005") Then
    '
    '    cmdVerFirma.Visible = True
    '    cmdVerFirma.Enabled = True
    '
    '    cmdActualizarFirma.Visible = False
    '    cmdActualizarFirma.Enabled = False
    'End If
    ''End GITU
    
    'RIRO20140128 Segun Peticion TIC1401200017 ***************
    Dim oConst As COMDConstSistema.DCOMGeneral
    Dim sConstante As String
    Set oConst = New COMDConstSistema.DCOMGeneral
    sConstante = oConst.LeeConstSistema(465)
    Set oConst = Nothing
    If InStr(1, sConstante, gsCodCargo) = 0 Then
        cmdVerFirma.Visible = True
        cmdVerFirma.Enabled = True
        cmdActualizarFirma.Visible = False
        cmdActualizarFirma.Enabled = False
    End If
    'END RIRO ************************************************
    
    EstableceEdicionSujetoObligado (fbPermisoEditarSujetoObligadoDJ) 'EJVG20120815
    'EJVG20111217 **********************************
    bPermisoEditarTodo = True
    If oPersona.Personeria = gPersonaNat Then 'EJVG20120120
        If ObtenerVecesCreditoyAhorroPersona(oPersona.PersCodigo) > 0 Then
            If Not validaPermisoEditarPersona(gsCodCargo, gsCodPersUser) Then 'RIRO20141106 ERS159 gsCodPersUser
                bPermisoEditarTodo = False
                HabilitarControlesDatosPrincipalesPersonas (bPermisoEditarTodo)
            End If
        End If
    End If

    'ANDE 20190910  comprobar si uso de Caja Maynas Online está habilitado
    If oPersona.BIHabilitado = True Then
        Dim NomBI As String, mensaje As String
        Set oConst = New COMDConstSistema.DCOMGeneral
        NomBI = oConst.LeeConstSistema(116)
        mensaje = "Usuario está afiliado al uso de " & NomBI & ", motivo por el cual no podrá editar el Celular 1 y Email 1."

        MsgBox mensaje, vbInformation + vbOKOnly, "Aviso"
        txtCel1.Enabled = Not oPersona.BIHabilitado
        TxtEmail.Enabled = Not oPersona.BIHabilitado
    End If
    'end ande 20190910
    'WIOR 20130827 *********************************************
    Dim oGen As COMDConstSistema.DCOMGeneral
     Set oGen = New COMDConstSistema.DCOMGeneral
     'fbPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gEdicionPersonas)
     fbPermisoCargo = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gEdicionPersonas, gsCodPersUser) 'RIRO20141027 ERS159
     
     If Not bPermisoEditarTodo Or oPersona.Personeria <> gPersonaNat Then
          HabilitarControlesDatosBasicos (fbPermisoCargo)
     End If
     If bPermisoEditarTodo Then
        fbPermisoCargo = True
     End If
    'WIOR FIN **************************************************
    'FRHU 20151204 ERS077-2015
    If fbPermisoCargo Then
        lblAutorizarUsoDatos.Visible = True
        CboAutoriazaUsoDatos.Visible = True
        'cboAutoriazaUsoDatos.Enabled = True' comentado por pti1 ers070-2018
        'add pti1 ers070-2018
        If oPersona.AutorizaUsoDatos = 1 Then
         CboAutoriazaUsoDatos.Enabled = False
        Else
         CboAutoriazaUsoDatos.Enabled = True
        End If
        'end pti1
    Else
        CboAutoriazaUsoDatos.Enabled = False
    End If
    'FIN FRHU 20151204
    bValidaCampDatos = True 'JUEZ 20131024
    'JUEZ 20140605 *******************************************
    Dim oDPersGen As COMDPersona.DCOMPersGeneral
    Set oDPersGen = New COMDPersona.DCOMPersGeneral
    Dim RSPersGen As ADODB.Recordset
    Set RSPersGen = oDPersGen.VerificaSolicitudAutorizacionRiesgos(oPersona.PersCodigo)
    If Not (RSPersGen.EOF And RSPersGen.BOF) Then
        TxtCodCIIU.Enabled = False
        CboPersCiiu.Enabled = False
    End If
    'END JUEZ ************************************************
    
    Me.lblTopera.Caption = "M" '*** CTI3 07092018
End Sub
'EJVG20120815 ***
Private Sub EstableceEdicionSujetoObligado(ByVal pbPermisoEditar As Boolean)
    Dim bPermiso As Boolean
    bPermiso = False
    If oPersona.SujetoObligadoIni <> -1 Then
        bPermiso = pbPermisoEditar
    Else
        bPermiso = True
    End If
    Me.cboSujetoObligado.Enabled = bPermiso
    Me.cboOfCumplimiento.Enabled = bPermiso
End Sub
'END EJVG *******
Private Sub HabilitarControlesDatosPrincipalesPersonas(ByVal pbHabilita As Boolean)
    cmbPersNatSexo.Enabled = pbHabilita
    cmbPersNatEstCiv.Enabled = pbHabilita
    txtPersNombreAP.Enabled = pbHabilita
    txtPersNombreAM.Enabled = pbHabilita
    txtApellidoCasada.Enabled = pbHabilita
    txtPersNombreN.Enabled = pbHabilita
    txtPersNacCreac.Enabled = pbHabilita
    cmdPersIDAceptar.Enabled = pbHabilita
    cmdPersIDedit.Enabled = pbHabilita
    cmdPersIDDel.Enabled = pbHabilita
End Sub

'WIOR 20130826 *****TI-ERS119-2013**********************************************
Private Sub HabilitarControlesDatosBasicos(ByVal pbHabilita As Boolean)
If oPersona.Personeria = gPersonaNat Then
    txtPersNombreAP.Enabled = pbHabilita
    txtPersNombreAM.Enabled = pbHabilita
    txtApellidoCasada.Enabled = pbHabilita
    txtPersNombreN.Enabled = pbHabilita
Else
    txtPersNombreRS.Enabled = pbHabilita
End If
cmdPersIDAceptar.Enabled = pbHabilita
cmdPersIDedit.Enabled = pbHabilita
cmdPersIDDel.Enabled = pbHabilita
End Sub
'WIOR FIN *******************************************************************

Private Sub CmdFteIngEditar_Click()
    If Trim(FEFteIng.TextMatrix(0, 1)) = "" Then
        MsgBox "No Existen Registros para Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If cmdPersFteIngresoEjecutado = 1 Then  'Se ingreso Fte de Ingreso
        MsgBox "Debe grabar los datos antes de editar la Fuente de Ingreso", vbInformation, "Mensaje"
        Exit Sub
    End If
    If oPersona.NumeroFtesIngreso > 0 Then
        If gsProyectoActual = "C" Then
            Call frmFteIngresosCS.Editar(FEFteIng.row - 1, oPersona)
            cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, CStr(frmFteIngresosCS.nPersMagnitudEmp))
            lblMagnitudEmpresarial.Caption = cmbPersJurMagnitud.Text
        Else
            Call frmFteIngresos.Editar(FEFteIng.row - 1, oPersona, rsHojEval)
        End If
    Else
        MsgBox "No Existe Fuentes de Ingreso para Editar", vbInformation, "Aviso"
    End If
End Sub


Private Sub CmdFteIngEliminar_Click()
    If Trim(FEFteIng.TextMatrix(0, 1)) = "" Then
        MsgBox "No Existen Registros para Eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Esta Seguro que Desea Eliminar esta Fuente de Ingreso", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarFteIngTipoAct(PersFilaEliminda, FEFteIng.row - 1)
        FEFteIng.EliminaFila FEFteIng.row
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            Call CmdPersAceptar_Click
        Else
            oPersona.LimpiaEliminados
        End If
        Call CargaFuentesIngreso
    End If
End Sub

Private Sub CmdFteIngNuevo_Click()
Dim sNombreCompleto As String
    
    'If bNuevaPersona Then
    '    MsgBox "Se debe Crear la Persona "
    'End If
    
    If Not ValidaControles Then
        'bPersonaAGrabar = True
        Exit Sub
    End If
    '05-06-2006
    If FEFteIng.TextMatrix(0, 1) = "" And bPersonaAGrabar And TxtBCodPers.Text = "" Then  'Si no existen Fuentes de Ingreso
        Call CmdPersAceptar_Click           'Grabar los Datos de la Persona
        bPersonaAGrabar = False
        Exit Sub
    End If
    '--------------------
    If Trim(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
        If Trim(Right(cmbPersNatSexo.Text, 2)) <> "F" And Len(Trim(txtApellidoCasada.Text)) > 0 Then
            sNombreCompleto = txtPersNombreAP.Text & "/" & txtPersNombreAM.Text & "\" & txtApellidoCasada.Text & "," & txtPersNombreN.Text
        Else
            sNombreCompleto = txtPersNombreAP.Text & "/" & txtPersNombreAM.Text & "," & txtPersNombreN.Text
        End If
    Else
        sNombreCompleto = txtPersNombreRS.Text
    End If
    oPersona.NombreCompleto = sNombreCompleto
    If Trim(FEFteIng.TextMatrix(1, 1)) = "" Then
        If gsProyectoActual <> "C" Then
            Call frmFteIngresos.NuevaFteIngreso(oPersona)
        Else
            Call frmFteIngresosCS.NuevaFteIngreso(oPersona)
            cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, CStr(frmFteIngresosCS.nPersMagnitudEmp))
            lblMagnitudEmpresarial.Caption = cmbPersJurMagnitud.Text
        End If
    Else
        If gsProyectoActual <> "C" Then
            Call frmFteIngresos.NuevaFteIngreso(oPersona, CInt(Trim(FEFteIng.TextMatrix(FEFteIng.row, 5))))
        Else
            Call frmFteIngresosCS.NuevaFteIngreso(oPersona, CInt(Trim(FEFteIng.TextMatrix(FEFteIng.row, 5))))
            cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, CStr(frmFteIngresosCS.nPersMagnitudEmp))
            lblMagnitudEmpresarial.Caption = cmbPersJurMagnitud.Text
        End If
    End If
    cmdPersFteIngresoEjecutado = 1
    Call CargaFuentesIngreso
End Sub

Private Sub cmdNuevo_Click()
    Call HabilitaControles_BotonNuevo
    SSTabs.Tab = 0
    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
    oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 10))
    
    '*** PEAC 20080412
    oPersona.TIPOCOMPDescripcion = Trim(Right(cboTipoComp.Text, 10))
    oPersona.TIPOSISTINFORDescripcion = Trim(Right(cboTipoSistInfor.Text, 10))
    oPersona.CADENAPRODDescripcion = Trim(Right(cboCadenaProd.Text, 10))
    
    oPersona.MonedaPatri = Trim(Right(cboMonePatri.Text, 10))
    '*** FIN PEAC
    
    oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
    '*** PEAC 20080801
    oPersona.MotivoActu = Trim(Right(Me.cboMotivoActu.Text, 12))
    
    '** Juez 20120327 *************************************
    'oPersona.Residencia = chkResidente.value
    oPersona.Residencia = Trim(Right(cboResidente.Text, 1))
    If Trim(Right(cboResidente.Text, 1)) = 0 Then
        oPersona.PaisReside = Trim(Right(cboPaisReside.Text, 12))
    End If
    '**End Juez *******************************************
    
    bNuevaPersona = True
    bPersonaAct = True
    Call LlenarComboTipoDocumento
    
    'Add By GITU 2011-04-26 Valida para ver o actualizar la firma de acuerdo al cargo
    If (gsCodCargo <> "002001" And gsCodCargo <> "002002" And gsCodCargo <> "002003" And gsCodCargo <> "003001" And gsCodCargo <> "003002" _
        And gsCodCargo <> "004001" And gsCodCargo <> "004002" And gsCodCargo <> "006024" And gsCodCargo <> "007012" And gsCodCargo <> "006005") Then
        
        cmdVerFirma.Visible = True
        cmdVerFirma.Enabled = False
        
        cmdActualizarFirma.Visible = False
        cmdActualizarFirma.Enabled = True
    End If
    'End GITU
    
'    Call HabilitaControlesPersona(True)
'    Call LimpiarPantalla
'    bEstadoCargando = True
'    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, Trim(Str(gPersEstadoCivilCasado)))
'    cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, "04028")
'    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, "O9309")
'    chkResidente.value = 1
'    bEstadoCargando = False
'    'Inicializa Controles
'    TxtBCodPers.Enabled = False
'    cmdNuevo.Enabled = False
'    cmdEditar.Enabled = False
'    SSTabs.Enabled = True
'    SSTabs.TabVisible(0) = True
'    SSTabs.TabVisible(1) = True
'    SSTabs.Tab = 0
'    SSTDatosGen.Enabled = True
'    SSTDatosGen.Tab = 0
'    SSTIdent.Enabled = True
'    CmdPersAceptar.Visible = True
'    CmdPersCancelar.Visible = True
'
'    If oPersona Is Nothing Then
'       Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona
'    End If
'    Call oPersona.NuevaPersona
'    oPersona.TipoActualizacion = PersFilaNueva
'
'    Call HabilitaControlesPersonaFtesIngreso(True)
'    cmbPersNatSexo.ListIndex = 0
'    cmbPersPersoneria.ListIndex = 0
'
'    If Not bBuscaNuevo Then
'        cmbPersPersoneria.SetFocus
'    End If
'    If txtPersNombreAP.Enabled And txtPersNombreAP.Visible Then
'        txtPersNombreAP.SetFocus
'    End If
'    oPersona.EstadoCivil = Trim(Right(cmbPersNatEstCiv.Text, 10))
'    oPersona.CIIU = Trim(Right(CboPersCiiu.Text, 10))
'    oPersona.Nacionalidad = Trim(Right(cmbNacionalidad.Text, 12))
'    oPersona.Residencia = chkResidente.value
'    bNuevaPersona = True
'    bPersonaAct = True
'
'    FEFteIng.Enabled = False
'    CmdFteIngNuevo.Enabled = False
'    CmdFteIngEditar.Enabled = False
'    CmdFteIngEliminar.Enabled = False
'    CmdPersFteConsultar.Enabled = False
'    TxtBCodPers.Text = ""
    'WIOR 20130827 ***************************
    Set rsDocPersActual = Nothing
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN ********************************
    
    Me.lblTopera.Caption = "N" '*** CTI3 07092018
    
    Dim frmCargado As Integer
    frmCargado = olimpiar.IsFormLoaded(frmPersonaJurDatosAdic)
    If frmCargado Then Call olimpiar.LimpiarGrillasAccionariales(frmPersonaJurDatosAdic.fleAccionistas, frmPersonaJurDatosAdic.fleDirectorio, frmPersonaJurDatosAdic.fleGerencias, frmPersonaJurDatosAdic.flePatrimonio, frmPersonaJurDatosAdic.flePatOtrasEmpresa, frmPersonaJurDatosAdic.fleCargos)
    
End Sub

Private Sub HabilitaControles_BotonNuevo()
    
    Call HabilitaControlesPersona(True)
    
    '*** PEAC 20080801
    Me.cboMotivoActu.Enabled = False
    
    Call LimpiarPantalla
    bEstadoCargando = True
    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, Trim(str(gPersEstadoCivilCasado)))
    cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, "04028")
    
    
    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, "O9309")
    
    '*** PEAC 20080412
    cboTipoComp.ListIndex = IndiceListaCombo(cboTipoComp, "O9309")
    cboTipoSistInfor.ListIndex = IndiceListaCombo(cboTipoSistInfor, "O9309")
    cboCadenaProd.ListIndex = IndiceListaCombo(cboCadenaProd, "O9309")
    
    cboMonePatri.ListIndex = IndiceListaCombo(cboMonePatri, "O9309")
    '*** FIN PEAC
    
    '** Juez 20120327 ************
    'chkResidente.value = 1
    cboResidente.ListIndex = 0
    cboPaisReside.Enabled = False
    '** End Juez *****************
    
    bEstadoCargando = False
    'Inicializa Controles
    TxtBCodPers.Enabled = False
    cmdnuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTabs.Enabled = True
    SSTabs.TabVisible(0) = True
    SSTabs.TabVisible(1) = True
    'SSTabs.TabVisible(9) = True 'MAVM 20100606 BAS II 'FRHU 20150311 ERS013-2015
    'SSTabs.Tab = 0
    SSTDatosGen.Enabled = True
    SSTDatosGen.Tab = 0
    SSTIdent.Enabled = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    If oPersona Is Nothing Then
       Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona  'DPersona
    End If
    Call oPersona.NuevaPersona
    oPersona.TipoActualizacion = PersFilaNueva
    
    Call HabilitaControlesPersonaFtesIngreso(True)
    cmbPersNatSexo.ListIndex = 0
    cmbPersPersoneria.ListIndex = 0
        
    If Not bBuscaNuevo Then
        cmbPersPersoneria.SetFocus
    End If
    If txtPersNombreAP.Enabled And txtPersNombreAP.Visible Then
        txtPersNombreAP.SetFocus
    End If

    Me.TxtCodCIIU.Enabled = True
    FEFteIng.Enabled = False
    CmdFteIngNuevo.Enabled = False
    CmdFteIngEditar.Enabled = False
    CmdFteIngEliminar.Enabled = False
    CmdPersFteConsultar.Enabled = False
    TxtBCodPers.Text = ""
    SSTabs.Tab = 0
    TxtSbs.Text = "0000000000"
    TxtSbs.Enabled = False
End Sub

Private Sub cmdPatVehAcepta_Click()

If Len(Trim(cboMonePatri.Text)) = 0 Then
    MsgBox "Seleccionar la moneda de la DDJJ.", vbInformation, "Aviso"
    cboMonePatri.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End If

If Me.fePatVehicular.Visible = True Then
    If Not ValidaDatosPatVehicular Then
        Exit Sub
    End If
   If cmdPersPatVehicularEjecutado = 1 Then
        Call oPersona.AdicionaPatVehicular
        Call oPersona.ActualizarPatVehTipoAct(PersFilaNueva, fePatVehicular.row - 1)
        lnNumPatVeh = lnNumPatVeh + 1
        fePatVehicular.TextMatrix(fePatVehicular.row, 6) = lnNumPatVeh
    Else
        If cmdPersPatVehicularEjecutado = 2 Then
            If oPersona.ObtenerPatVehTipoAct(fePatVehicular.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarPatVehTipoAct(PersFilaModificada, fePatVehicular.row - 1)
            End If
        End If
    End If
    'Marca
    Call oPersona.ActualizaPatVehMarca(UCase(fePatVehicular.TextMatrix(fePatVehicular.row, 1)), fePatVehicular.row - 1)
    'Fecha Fabricación
    If fePatVehicular.TextMatrix(fePatVehicular.row, 2) <> "" Then
        Call oPersona.ActualizaPatVehFecFab(CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)), fePatVehicular.row - 1)
    End If
    'Valor de Comercializacion
    Call oPersona.ActualizaPatVehValCom(fePatVehicular.TextMatrix(fePatVehicular.row, 3), fePatVehicular.row - 1)
    'Valor de Modelo
    Call oPersona.ActualizaPatVehModelo(UCase(fePatVehicular.TextMatrix(fePatVehicular.row, 4)), fePatVehicular.row - 1)
    'Valor de Placa
    Call oPersona.ActualizaPatVehPlaca(UCase(fePatVehicular.TextMatrix(fePatVehicular.row, 5)), fePatVehicular.row - 1)
    '*** PEAC 20080412
    'Condicion del Vehiculo
    'Call oPersona.ActualizaPatVehCondicion(fePatVehicular.TextMatrix(fePatVehicular.Row, 4), fePatVehicular.Row - 1)
    Call oPersona.ActualizaPatVehCod(fePatVehicular.TextMatrix(fePatVehicular.row, 6), fePatVehicular.row - 1)
    'Habilitar Controles
    cmdPersPatVehicularEjecutado = 0
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatVehicular.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    fePatVehicular.SetFocus
End If

If Me.fePatOtros.Visible = True Then
    If Not ValidaDatosPatOtros Then
        Exit Sub
    End If
   If cmdPersPatVehicularEjecutado = 1 Then
        Call oPersona.AdicionaPatOtros
        Call oPersona.ActualizarPatOtrosTipoAct(PersFilaNueva, fePatOtros.row - 1)
        lnNumPatOtros = lnNumPatOtros + 1
        fePatOtros.TextMatrix(fePatOtros.row, 3) = lnNumPatOtros
    Else
        If cmdPersPatVehicularEjecutado = 2 Then
            If oPersona.ObtenerPatOtrosTipoAct(fePatOtros.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarPatOtrosTipoAct(PersFilaModificada, fePatOtros.row - 1)
            End If
        End If
    End If
    
    'descripcion
    Call oPersona.ActualizaPatOtrosDescripcion(UCase(fePatOtros.TextMatrix(fePatOtros.row, 1)), fePatOtros.row - 1)
    
    'Valor de Comercializacion
    Call oPersona.ActualizaPatOtrosValCom(fePatOtros.TextMatrix(fePatOtros.row, 2), fePatOtros.row - 1)
    
    Call oPersona.ActualizaPatOtrosCod(fePatOtros.TextMatrix(fePatOtros.row, 3), fePatOtros.row - 1)
    
    'Habilitar Controles
    cmdPersPatVehicularEjecutado = 0
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatOtros.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    fePatOtros.SetFocus
End If

If Me.fePatInmuebles.Visible = True Then
    If Not ValidaDatosPatInmuebles Then
        Exit Sub
    End If
   If cmdPersPatVehicularEjecutado = 1 Then
        Call oPersona.AdicionaPatInmuebles
        Call oPersona.ActualizarPatInmueblesTipoAct(PersFilaNueva, fePatInmuebles.row - 1)
        lnNumPatInmuebles = lnNumPatInmuebles + 1
        fePatInmuebles.TextMatrix(fePatInmuebles.row, 7) = lnNumPatInmuebles
    Else
        If cmdPersPatVehicularEjecutado = 2 Then
            If oPersona.ObtenerPatInmueblesTipoAct(fePatInmuebles.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarPatInmueblesTipoAct(PersFilaModificada, fePatInmuebles.row - 1)
            End If
        End If
    End If
    
    'Ubicacion
    Call oPersona.ActualizaPatInmueblesUbicacion(UCase(fePatInmuebles.TextMatrix(fePatInmuebles.row, 1)), fePatInmuebles.row - 1)
    
    'Area terreno
    Call oPersona.ActualizaPatInmueblesAreaTerreno(fePatInmuebles.TextMatrix(fePatInmuebles.row, 2), fePatInmuebles.row - 1)
    
    'Area construida
    Call oPersona.ActualizaPatInmueblesAreaConstru(fePatInmuebles.TextMatrix(fePatInmuebles.row, 3), fePatInmuebles.row - 1)
    
    'tipo uso
    Call oPersona.ActualizaPatInmueblesTipoUso(UCase(fePatInmuebles.TextMatrix(fePatInmuebles.row, 4)), fePatInmuebles.row - 1)
    
    'rrpp
    Call oPersona.ActualizaPatInmueblesRRPP(UCase(fePatInmuebles.TextMatrix(fePatInmuebles.row, 5)), fePatInmuebles.row - 1)
    
    'Valor de Comercializacion
    Call oPersona.ActualizaPatInmueblesValCom(fePatInmuebles.TextMatrix(fePatInmuebles.row, 6), fePatInmuebles.row - 1)
    
    Call oPersona.ActualizaPatInmueblesCod(fePatInmuebles.TextMatrix(fePatInmuebles.row, 7), fePatInmuebles.row - 1)
    
    'Habilitar Controles
    cmdPersPatVehicularEjecutado = 0
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatInmuebles.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    fePatInmuebles.SetFocus
End If


End Sub

Private Sub cmdPatVehCancela_Click()
    
If Me.fePatVehicular.Visible = True Then
    CargaPatVehicular
    'Habilitar Controles
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatVehicular.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    
    fePatVehicular.SetFocus
End If

If Me.fePatOtros.Visible = True Then
    CargaPatOtros
    'Habilitar Controles
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatVehicular.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    
    fePatOtros.SetFocus
End If

If Me.fePatInmuebles.Visible = True Then
    CargaPatInmuebles
    'Habilitar Controles
    FEPatVehPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    fePatVehicular.lbEditarFlex = False
    cmdPatVehNuevo.Enabled = True
    cmdPatVehEdita.Enabled = True
    cmdPatVehElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdPatVehAcepta.Visible = False
    cmdPatVehCancela.Visible = False
    
    fePatInmuebles.SetFocus
End If

End Sub

Private Sub cmdPatVehEdita_Click()

If Me.fePatVehicular.Visible = True Then

   If oPersona.NumeroPatVehicular > 0 Then
        cmdPersPatVehicularEjecutado = 2
        FEPatVehPersNoMoverdeFila = fePatVehicular.row
        NomMoverSSTabs = SSTabs.Tab
        fePatVehicular.lbEditarFlex = True
        fePatVehicular.SetFocus

        cmdPatVehNuevo.Enabled = False
        cmdPatVehEdita.Enabled = False
        cmdPatVehElimina.Enabled = False
        cmdPatVehAcepta.Visible = True
        cmdPatVehCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True

        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False

        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False

        fePatVehicular.SetFocus

    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End If

If Me.fePatOtros.Visible = True Then

   If oPersona.NumeroPatOtros > 0 Then
        cmdPersPatVehicularEjecutado = 2
        FEPatVehPersNoMoverdeFila = fePatOtros.row
        NomMoverSSTabs = SSTabs.Tab
        fePatOtros.lbEditarFlex = True
        fePatOtros.SetFocus

        cmdPatVehNuevo.Enabled = False
        cmdPatVehEdita.Enabled = False
        cmdPatVehElimina.Enabled = False
        cmdPatVehAcepta.Visible = True
        cmdPatVehCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True

        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False

        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False

        fePatOtros.SetFocus

    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End If

If Me.fePatInmuebles.Visible = True Then

   If oPersona.NumeroPatInmuebles > 0 Then
        cmdPersPatVehicularEjecutado = 2
        FEPatVehPersNoMoverdeFila = fePatInmuebles.row
        NomMoverSSTabs = SSTabs.Tab
        fePatInmuebles.lbEditarFlex = True
        fePatInmuebles.SetFocus

        cmdPatVehNuevo.Enabled = False
        cmdPatVehEdita.Enabled = False
        cmdPatVehElimina.Enabled = False
        cmdPatVehAcepta.Visible = True
        cmdPatVehCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True

        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False

        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False

        fePatInmuebles.SetFocus

    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End If


End Sub

Private Sub cmdPatVehElimina_Click()

If Me.fePatVehicular.Visible = True Then
    If MsgBox("Esta Seguro que Desea Eliminar el Patrimonio Vehicular", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarPatVehTipoAct(PersFilaEliminda, fePatVehicular.row - 1)
        Call CmdPersAceptar_Click
        Call CargaPatVehicular
    End If
End If

If Me.fePatOtros.Visible = True Then
    If MsgBox("Esta Seguro que Desea Eliminar el Patrimonio.", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarPatOtrosTipoAct(PersFilaEliminda, fePatOtros.row - 1)
        Call CmdPersAceptar_Click
        Call CargaPatOtros
    End If
End If

If Me.fePatInmuebles.Visible = True Then
    If MsgBox("Esta Seguro que Desea Eliminar el Patrimonio.", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarPatInmueblesTipoAct(PersFilaEliminda, fePatInmuebles.row - 1)
        Call CmdPersAceptar_Click
        Call CargaPatInmuebles
    End If
End If


End Sub

Private Sub cmdPatVehNuevo_Click()
       
If Me.fePatVehicular.Visible = True Then
    cmdPatVehAcepta.Visible = True
    cmdPatVehCancela.Visible = True
    cmdPatVehNuevo.Enabled = False
    cmdPatVehElimina.Enabled = False
    cmdPatVehEdita.Enabled = False
    
    fePatVehicular.lbEditarFlex = True
    fePatVehicular.AdicionaFila
    cmdPersPatVehicularEjecutado = 1
    FEPatVehPersNoMoverdeFila = fePatVehicular.rows - 1
    fePatVehicular.SetFocus
End If

If Me.fePatOtros.Visible = True Then
    cmdPatVehAcepta.Visible = True
    cmdPatVehCancela.Visible = True
    cmdPatVehNuevo.Enabled = False
    cmdPatVehElimina.Enabled = False
    cmdPatVehEdita.Enabled = False
    
    fePatOtros.lbEditarFlex = True
    fePatOtros.AdicionaFila
    cmdPersPatVehicularEjecutado = 1
    FEPatVehPersNoMoverdeFila = fePatOtros.rows - 1
    fePatOtros.SetFocus
End If

If Me.fePatInmuebles.Visible = True Then
    cmdPatVehAcepta.Visible = True
    cmdPatVehCancela.Visible = True
    cmdPatVehNuevo.Enabled = False
    cmdPatVehElimina.Enabled = False
    cmdPatVehEdita.Enabled = False
    
    fePatInmuebles.lbEditarFlex = True
    fePatInmuebles.AdicionaFila
    cmdPersPatVehicularEjecutado = 1
    FEPatVehPersNoMoverdeFila = fePatInmuebles.rows - 1
    fePatInmuebles.SetFocus
End If


End Sub

Private Sub CmdPersAceptar_Click()
Dim oPersonaNeg As COMNPersona.NCOMPersona    'npersona
Dim R As ADODB.Recordset
Dim nVerDuplicadoc As Integer
Dim nVerTamanioDoc As Integer
'ALPA 20080922********************************************************
Dim nSalir As Integer
'*********************************************************************
'ARCV 07-06-2006
Dim lsPersCodGrabar As String
'madm 20100408 ------------------------------------------------------
Dim lbResultadoVisto As Boolean
Dim loVistoElectronico As frmVistoElectronico
Set loVistoElectronico = New frmVistoElectronico
Dim loPersNegativa As frmPersNegativas
Set loPersNegativa = New frmPersNegativas
'--------------------------------------------------------------------
Dim J As Integer
Dim bRUC4, bRUC5 As Boolean 'ALPA 20120423
Dim mensaje As String
Dim parMensaje As String
Dim lnPEPS As Integer 'EJVG20120813
Dim oDPers As COMDPersona.DCOMPersonas 'JUEZ 20131007
Dim bDocExt, bDocPas As Boolean 'JUEZ 20131007
Dim lnCondicionPNeg As Integer '-->JGPA20191210 ACTA N° 106 - 2019
Dim lsParMensaje As String '-->JGPA20191210 ACTA N° 106 - 2019
Dim lsCondicion As String '-->JGPA20191210 ACTA N° 106 - 2019

    If Not ValidaRestriccion(mensaje) Then
        If Len(mensaje) > 0 Then
            MsgBox mensaje, vbInformation
            Exit Sub
        End If
    End If
    
    If Not ValidaControles Then
        Exit Sub
    End If
    
    On Error GoTo ErrorCmdPersAceptar

    Screen.MousePointer = 11
    
    'OJO a revisar: 10-12
    'oPersona.CampoActualizacion = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    oPersona.dFechaHoy = gdFecSis
    
    nSalir = ValidarNombresApellidos(txtPersNombreAP.Text, Len(txtPersNombreAP.Text), nSalir, 1)
    nSalir = ValidarNombresApellidos(txtPersNombreAM.Text, Len(txtPersNombreAM.Text), nSalir, 2)
    nSalir = ValidarNombresApellidos(txtPersNombreN.Text, Len(txtPersNombreN.Text), nSalir, 3)
    If nSalir <> 0 Then
        MsgBox ("Nombre tiene caracteres no aceptados")
        If nSalir = 1 Then
            txtPersNombreAP.SetFocus
        ElseIf nSalir = 2 Then
            txtPersNombreAM.SetFocus
        ElseIf nSalir = 3 Then
            txtPersNombreN.SetFocus
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If
    'ALPA 20120423*************************************************************
    bRUC4 = True 'False es que no cumple
    bRUC5 = False
    If Trim(Right(cmbPersPersoneria.Text, 4)) = "1" Then
        If FEDocs.rows > 0 Then
            If Trim(Right(cmbPersNatMagnitud.Text, 5)) = "5" Or Trim(Right(cmbPersNatMagnitud.Text, 5)) = "4" Then
                For J = 1 To FEDocs.rows - 1
                    If Trim(Right(cmbPersNatMagnitud.Text, 5)) = "4" And Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = "2" Then
                        bRUC4 = False
                    End If
                    If Trim(Right(cmbPersNatMagnitud.Text, 5)) = "5" And Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = "2" Then
                        bRUC5 = True
                    End If
                Next J
                If bRUC4 = False And Trim(Right(cmbPersNatMagnitud.Text, 5)) = "4" Then
                    MsgBox ("La magnitud Persona natural sin negocio no debe tener RUC")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                If bRUC5 = False And Trim(Right(cmbPersNatMagnitud.Text, 5)) = "5" Then
                    MsgBox ("La magnitud Persona natural con negocio independiente debe tener RUC")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
    End If
    '****************************************************************************
    
    'JUEZ 20131007 *******************************************
    bDocExt = False
    bDocPas = False
    Set oDPers = New COMDPersona.DCOMPersonas
    If Trim(Right(cmbPersPersoneria.Text, 4)) = gPersonaNat And Trim(Right(cmbNacionalidad.Text, 12)) <> "04028" And BotonEditar Then
        If FEDocs.rows > 0 Then
            If oDPers.VerificaPersonaConCreditoActivo(oPersona.PersCodigo) Then
                For J = 1 To FEDocs.rows - 1
                    'comentado Por '202105256LARI --SE AGREGARON VARIOS TIPOS DE DOCUMENTOS ADICIONALES EXTRANJEROS SEGÚN OFICIO 13323-2020-SBS

                    'If Trim(Right(FEDocs.TextMatrix(j, 1), 2)) = gPersIdExtranjeria Then
                    '        bDocExt = True
                    '    End If
                    If Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = gPersIdExtranjeria Or _
                       Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = gPersIdCIEMRE Or _
                       Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = gPersIdCPTP Or _
                       Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = gPersIdCREF Or _
                       Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = gPersIdCEPR Then
                        bDocExt = True
                    End If
                    '******************************************************************************************
                    If Trim(Right(FEDocs.TextMatrix(J, 1), 2)) = gPersIdPasaporte Then
                        bDocPas = True
                    End If
                Next J
                If bDocExt = False Then
                    MsgBox ("Es obligatorio registrar Carnet de Extranjeria para personas extranjeras con crédito activo")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                If bDocPas = False Then
                    MsgBox ("Es obligatorio registrar Pasaporte para personas extranjeras con crédito activo")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
    End If

    If oPersona.MotivoActu = "2" Then
        If oDPers.VerificaPersonaConCreditoActivo(oPersona.PersCodigo, True) Then
            If MsgBox("El cliente tiene calificación D o E, para actualizar su dirección debe contar con algún documento que lo sustente el mismo que deberá ser adjuntado a su expediente, ¿Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If
    'END JUEZ ************************************************
    
    'PEPS madm 20100408
    'MADM 20101221 EXTRAN - NO RESIDENTES
    'If (Me.chkpeps.Visible = True And Me.chkpeps.Enabled = True And Me.chkpeps.value = 1) Or (Trim(Right(cmbNacionalidad.Text, 12)) <> "04028") Or (chkResidente.value = 0) And (oPersona.Personeria = gPersonaNat) Then
    'EJVG20120813 ***
    'If (Me.chkpeps.Visible = True And Me.chkpeps.Enabled = True And Me.chkpeps.value = 1) Or (Trim(Right(cmbNacionalidad.Text, 12)) <> "04028") Or (Trim(Right(cboResidente.Text, 1)) = 0) And (oPersona.Personeria = gPersonaNat) Then '** Juez 20120327
    lnPEPS = CInt(IIf(IsNumeric(Trim(Right(Me.cboPEPS.Text, 1))), Trim(Right(Me.cboPEPS.Text, 1)), 0))
    If (Me.cboPEPS.Visible = True And Me.cboPEPS.Enabled = True And lnPEPS = 1) Or (Trim(Right(cmbNacionalidad.Text, 12)) <> "04028") Or (Trim(Right(cboResidente.Text, 1)) = 0) Or (lnCondicionBNeg = 1) And (oPersona.Personeria = gPersonaNat) Then 'EJVG20120724'WIOR 20121122 AGREGO lnCondicionBNeg
        'If Me.chkpeps.value = 1 And oPersona.TipoActualizacion = PersFilaNueva Then
        If lnPEPS = 1 And oPersona.TipoActualizacion = PersFilaNueva Then
    'END EJVG *******
             parMensaje = "PEPS"
             If VerificarAutorizacion(parMensaje) = False Then Exit Sub
        ElseIf (Trim(Right(cmbNacionalidad.Text, 12)) <> "04028" And oPersona.TipoActualizacion = PersFilaNueva And oPersona.Personeria = gPersonaNat) Then
            parMensaje = "EXTRANJERO"
            If VerificarAutorizacion(parMensaje) = False Then Exit Sub
        'ElseIf (chkResidente.value = 0 And oPersona.TipoActualizacion = PersFilaNueva And oPersona.Personeria = gPersonaNat) Then
        ElseIf (Trim(Right(cboResidente.Text, 1)) = 0 And oPersona.TipoActualizacion = PersFilaNueva And oPersona.Personeria = gPersonaNat) Then '** Juez 20120327
            parMensaje = "NO RESIDENTE"
            If VerificarAutorizacion(parMensaje) = False Then Exit Sub
        'WIOR 20121122 VALIDA EL REGISTRO DE PERSONA NEGATIVA ***************************
        ElseIf lnCondicionBNeg = 1 Then
            parMensaje = "NEGATIVO"
            If VerificarAutorizacion(parMensaje) = False Then Exit Sub
            lnCondicionBNeg = 0
        'WIOR FIN ************************************************************************
        End If
    'end madm
    End If
    '***JGPA20191210 ACTA N° 106 - 2019
    Set oPersonaNeg = New COMNPersona.NCOMPersona
    
    '20210116LARI: VALIDA POR EL NUMDOCID
    'SI ENCUENTRA EL NUMDOCID ENVIA MENSAJE
    'SINO ENCUENTRA EL NUMDOCID HACE LA CONSULTA POR EL NOMBRE DE LA PERSONA REGISTRADA
    If FEDocs.rows > 0 Then 'JGPA20200117
        For J = 1 To FEDocs.rows - 1
            lnCondicionPNeg = oPersonaNeg.VerificaPersonaListaNegativa(lsCondicion, lsParMensaje, FormateaTexto(txtPersNombreN.Text), FormateaTexto(txtPersNombreAP.Text), Trim(FEDocs.TextMatrix(J, 2)), 1, FormateaTexto(txtPersNombreAM.Text), FormateaTexto(txtApellidoCasada.Text)) 'JGPA20200117 Added [Trim(FEDocs.TextMatrix(FEDocs.row, 2)),1]
            If lnCondicionPNeg > 0 Then
                Exit For
            '20210116LARI: VALIDA POR NOMBRE DE LA PERSONA: TAMBIEN PASA POR PARAMETRO EL DOI***********
            Else
                lnCondicionPNeg = oPersonaNeg.VerificaPersonaListaNegativa(lsCondicion, lsParMensaje, FormateaTexto(txtPersNombreN.Text), FormateaTexto(txtPersNombreAP.Text), Trim(FEDocs.TextMatrix(J, 2)), 2, FormateaTexto(txtPersNombreAM.Text), FormateaTexto(txtApellidoCasada.Text)) 'JGPA20200117 Added ["", 2]
                If lnCondicionPNeg > 0 Then
                    MsgBox lsParMensaje, vbInformation, "Aviso"
                    If nTipoAccion = 1 Then
                        If VerificarAutorizacion(lsCondicion) = False Then Exit Sub
                    End If
                    Exit For
                End If
            '*******************************************************************************************
            End If
        Next J
    End If
    '20210116LARI: COMENTADO POR LARI*******************************
    'lnCondicionPNeg = oPersonaNeg.VerificaPersonaListaNegativa(lsCondicion, lsParMensaje, FormateaTexto(txtPersNombreN.Text), FormateaTexto(txtPersNombreAP.Text), "", 2, FormateaTexto(txtPersNombreAM.Text), FormateaTexto(txtApellidoCasada.Text)) 'JGPA20200117 Added ["", 2]
    'If lnCondicionPNeg > 0 Then
    '    MsgBox lsParMensaje, vbInformation, "Aviso"
    '    If nTipoAccion = 1 Then 'JGPA20200117 Observación [GECA] solo para registro
    '        If VerificarAutorizacion(lsCondicion) = False Then Exit Sub
    '   End If
    'End If
    '****************************************************************
    Set oPersonaNeg = Nothing
    '***End JGPA20191210
    'GIPO
   If TxtSbs.Text <> "0000000000" Then
        Set oDPers = New COMDPersona.DCOMPersonas
        Dim rsComprobar As ADODB.Recordset
        Set rsComprobar = oDPers.seleccionarCodigoSBSRepetido(TxtSbs.Text, TxtBCodPers.Text)
        If rsComprobar.RecordCount > 0 Then
             MsgBox "El Código SBS ya existe.", vbInformation, "Aviso"
             TxtSbs.SetFocus
             Screen.MousePointer = 0
             Exit Sub
        End If
   End If
   
'**************** cti3 17092018
Set oValidaGrilla = New UPersona_Cli
Dim sResult As String
nValor = frmPersonaJurDatosAdic.lblPersElejida.Caption

If nValor >= 0 And SSTabs.TabVisible(0) = True Then  'persona
 sGrillaLleno = oValidaGrilla.validallenado(frmPersonaJurDatosAdic.fleAccionistas)
  If sGrillaLleno = False Then
  
        sResult = MsgBox("No se a registrado Datos Accionariales: Desea registrar?", vbYesNo + vbInformation, "AVISO")
        Screen.MousePointer = 0
        If sResult = 6 Then
        
            frmPersonaJurDatosAdic.SSTabs.TabVisible(0) = True
            frmPersonaJurDatosAdic.SSTabs.TabVisible(1) = True
            frmPersonaJurDatosAdic.SSTabs.TabVisible(2) = True
            frmPersonaJurDatosAdic.SSTabs.TabVisible(3) = True
            frmPersonaJurDatosAdic.SSTabs.TabVisible(4) = True
            frmPersonaJurDatosAdic.SSTabs.TabVisible(5) = True
            frmPersonaJurDatosAdic.lblPersElejida.Caption = 1
            frmPersonaJurDatosAdic.Show 1
            Exit Sub
        End If
  End If

Else
 If nValor >= 0 And SSTabs.TabVisible(1) = True Then  'juridico
    sGrillaLleno = oValidaGrilla.validallenado(frmPersonaJurDatosAdic.fleAccionistas)
    If sGrillaLleno = False Then
        MsgBox "Debe agregar Datos Accionariales: Los Datos Son Obligatorios", vbInformation, "AVISO"
        Screen.MousePointer = 0
        frmPersonaJurDatosAdic.SSTabs.TabVisible(0) = True
        frmPersonaJurDatosAdic.SSTabs.TabVisible(1) = True
        frmPersonaJurDatosAdic.SSTabs.TabVisible(2) = True
        frmPersonaJurDatosAdic.SSTabs.TabVisible(3) = True
        frmPersonaJurDatosAdic.SSTabs.TabVisible(4) = True
        frmPersonaJurDatosAdic.SSTabs.TabVisible(5) = True
        frmPersonaJurDatosAdic.lblPersElejida.Caption = 2
        frmPersonaJurDatosAdic.Show 1
        Exit Sub
    End If
 End If
End If

'******************************
   
    If MsgBox("Se va a proceder a guardar los datos de la Persona, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Call oPersona.GrabarCambiosPersona(gsCodCMAC, gsCodAge, gdFecSis, _
                                       gsCodUser, R, nVerDuplicadoc, nVerTamanioDoc, , lsPersCodGrabar, IIf(oPersona.MotivoActu = "2", txtBUsuario.Text, "")) 'JUEZ 20131007 Se agregó IIf(oPersona.MotivoActu = "2", Me.txtBUsuario, "")
    
  '  Call oPersona.GrabarCambiosPersona(gsCodCMAC, gsCodAge, gdFecSis, _
                                       gsCodUser, R, nVerDuplicadoc, nVerTamanioDoc, , lsPersCodGrabar, IIf(oPersona.MotivoActu = "2", txtBUsuario.Text, "")) 'JUEZ 20131007 Se agregó IIf(oPersona.MotivoActu = "2", Me.txtBUsuario, "")
    
    'Verifica Homonimia
    If oPersona.TipoActualizacion = PersFilaNueva Then
        If Not R.BOF And Not R.EOF Then
            Call frmMuestraHomonimia.Inicio(R)
            Screen.MousePointer = 0
            If MsgBox("Existen posibles personas Homonimas. Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmacion") = vbNo Then
                Exit Sub
            Else
                    Call oPersona.GrabarCambiosPersona(gsCodCMAC, gsCodAge, gdFecSis, _
                                       gsCodUser, R, nVerDuplicadoc, nVerTamanioDoc, False, lsPersCodGrabar, IIf(oPersona.MotivoActu = "2", txtBUsuario.Text, "")) 'JUEZ 20131007 Se agregó IIf(oPersona.MotivoActu = "2", Me.txtBUsuario, "")
            End If
        End If
    End If
            
    'Verificando Duplicidad de Documento
    If nVerDuplicadoc <> -1 Then
        MsgBox "Documento " & Trim(Left(FEDocs.TextMatrix(FEDocs.row, 1), 30)) & " se Encuentra Duplicado", vbInformation, "Aviso"
        cmdPersIDedit.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Verificando Tamaño de Documento
    If nVerTamanioDoc <> -1 Then
        MsgBox "Documento " & Trim(Left(FEDocs.TextMatrix(FEDocs.row, 1), 30)) & " numero de digitos Incorrecto", vbInformation, "Aviso"
        cmdPersIDedit.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'PEPS madm 20100408
     'If Me.chkpeps.Enabled = True And Me.chkpeps.Visible = True And Me.chkpeps.value = 1 Then 'And oPersona.TipoActualizacion = PersFilaNueva
     If Me.cboPEPS.Visible = True And Me.cboPEPS.Enabled = True And lnPEPS = 1 Then 'EJVG20120813
        loPersNegativa.Inicio FEDocs.TextMatrix(FEDocs.row, 2), Right(FEDocs.TextMatrix(FEDocs.row, 1), 2), txtPersNombreN.Text, txtPersNombreAP.Text, txtPersNombreAM.Text, IIf(txtApellidoCasada.Visible = True, Trim(txtApellidoCasada.Text), ""), 1 'WIOR 20130201 AGREGO txtApellidoCasada 'JGPA20191217 - Obs. [NALG] Added [1]
     End If
    'end madm
    
    '***MODIFICADO POR ELRO 20110803, según ACTA Nº 168-2011/TI-D
    If parMensaje <> "" Then
     Dim ofrmClienteProcesoReforzado As New frmPersClienteSensible
     Dim lsMovNroPersonaAtencion, lsMovNroPersonaAutoriza, sNombreCompletox As String
     Dim lsCentroLaboral As String
     Dim oClienteProcesoReforzado As New COMNPersona.NCOMPersona
     Dim lrsPersNegativaAutorizacion As New ADODB.Recordset
     Dim lrsPersListaNegativa  As New ADODB.Recordset
     Dim oNCOMContFunciones As New COMNContabilidad.NCOMContFunciones

      If Right(Me.cmbPersPersoneria.Text, 1) = "1" Then
        If Right(Me.cmbPersNatSexo.Text, 1) = "F" And Len(Trim(txtApellidoCasada)) > 0 Then
            If Right(Me.cmbPersNatEstCiv.Text, 1) = "3" Then
                sNombreCompletox = txtPersNombreAP & "/" & txtPersNombreAM & "\VDA " & txtApellidoCasada & "," & txtPersNombreN
            Else
                sNombreCompletox = txtPersNombreAP & "/" & txtPersNombreAM & "\" & txtApellidoCasada & "," & txtPersNombreN
            End If
        Else
            sNombreCompletox = txtPersNombreAP & "/" & txtPersNombreAM & "," & txtPersNombreN
        End If
     End If

     Set lrsPersNegativaAutorizacion = oClienteProcesoReforzado.mostrarPersNegativaAutorizacion(gdFecSis, sNombreCompletox, gsCodUser)
     lsMovNroPersonaAutoriza = lrsPersNegativaAutorizacion.Fields(8)
     lrsPersNegativaAutorizacion.Close: Set lrsPersNegativaAutorizacion = Nothing

     lsMovNroPersonaAtencion = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

      If parMensaje = "PEPS" Then
         Set lrsPersListaNegativa = oClienteProcesoReforzado.mostrarPersListaNegativa(Right(FEDocs.TextMatrix(FEDocs.row, 1), 2), FEDocs.TextMatrix(FEDocs.row, 2), 3)
         lsCentroLaboral = lrsPersListaNegativa.Fields(10)
         lrsPersListaNegativa.Close: Set lrsPersListaNegativa = Nothing
      Else
         lsCentroLaboral = ""
      End If


     Call ofrmClienteProcesoReforzado.Inicio(lsPersCodGrabar, lsMovNroPersonaAtencion, lsMovNroPersonaAutoriza, parMensaje, Right(Me.cmbPersNatEstCiv.Text, 1), lsCentroLaboral)

     Set oNCOMContFunciones = Nothing
     Set oClienteProcesoReforzado = Nothing
    End If
    '************************************************************
    
    'EJVG20120813 *** Adecuación Sujeto Obligado DJ
    ''JACA 20110730 Verifica si es Sujeto Obligado a Declarar a la UIF
    Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim ldFechaHoraDJ As Date
    'Dim bSujetoObligado As Boolean
    
    'Dim nPersoneriaDJ As Integer
    'bSujetoObligado = False
    'nPersoneriaDJ = oPersona.Personeria
    Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios

    'If nPersoneriaDJ = 4 Or nPersoneriaDJ = 5 Or nPersoneriaDJ = 6 Or nPersoneriaDJ = 7 Or nPersoneriaDJ = 8 Or nPersoneriaDJ = 9 Then
    '    bSujetoObligado = True
    'ElseIf oPersona.RcActiGiro1 = "21" Then'Comentado by JACA 20111019 para no incluir a los cambistas
    '     'bSujetoObligado = True
    'ElseIf clsServ.obtenerSectorDJSujetoObligado(oPersona.CIIU) Then
    '    bSujetoObligado = True
    'End If
    
    'If bSujetoObligado = True Then
    '    If Not clsServ.obtenerDJSujetoObligado(lsPersCodGrabar) Then ' verifica si ya esta registrado
    '       MsgBox "Esta Persona se Encuentra Obligado a Declarar ante la UIF,Coloque papel para imprimir la Declaracion Jurada", vbInformation, "DJ Sujeto Obligado"
    '        clsServ.guardarDJSujetoObligado lsPersCodGrabar, gdFecSis, gsCodUser
    '        imprimirDJSujetoObligado
    '
    '    End If
    'End If
    ''JACA END***************************************************
    If oPersona.SujetoObligadoIni <> oPersona.SujetoObligado Then
        ldFechaHoraDJ = CDate(gdFecSis & " " & Format(Time, "hh:mm:ss"))
        clsServ.guardarDJSujetoObligado lsPersCodGrabar, ldFechaHoraDJ, gsCodUser, Right(gsCodAge, 2), oPersona.SujetoObligado, oPersona.OfCumplimiento
        If oPersona.SujetoObligado = 1 Then
            MsgBox "Esta Persona se encuentra Obligado a Declarar ante la UIF, se va a emitir la Declaración Jurada", vbInformation, "DJ Sujeto Obligado"
            imprimirDJSujetoObligado2 (ldFechaHoraDJ)
        End If
    End If
    'END EJVG *******
    
    Call HabilitaControlesPersona(False)
    TxtBCodPers.Text = oPersona.PersCodigo
    
    '***Agreado por ELRO el 20130219, según INC1302150010
    lsDireccionActualizada = txtPersDireccDomicilio.Text
    '***Fin Agreado por ELRO el 20130219*****************
    
    'WIOR 20130827 **************************************
    
    'ADD PTI1 ERS070-2018
    If bPemisoAD And nTipoForm = 1 Then
    fbPermisoCargo = True
    End If
    'FIN PTI1
    
   ' If fbPermisoCargo Then 'COMENTADO POR PTI1 ERS070-2018
     If fbPermisoCargo And CboAutoriazaUsoDatos.ListIndex <> -1 Then 'ADD POR PTI1 ERS070-2018
        Dim cMovCambio As String
        Dim oCambio As COMDPersona.DCOMPersonas
        Dim bCabecera As Boolean
        Dim nI, nJ, nZ, nAux As Integer
        Set oCambio = New COMDPersona.DCOMPersonas
        
        Dim nTpoDocAnt, nTpoDocAct As Integer
        Dim cNumDocAnt, cNumDocAct As String
        bCabecera = False
        Set rsDocPersUlt = FEDocs.GetRsNew(0)
        
        'add pti1 ers070-2018
        If nTipoForm = 1 Then
          Set rsDocPersActual = FEDocs.GetRsNew(0)
        End If
        'END add pti1 ers070-2018
      
        nI = rsDocPersActual.RecordCount
        nJ = rsDocPersUlt.RecordCount
        
        nZ = IIf(nI > nJ, nI, nJ)
        
        cMovCambio = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        
        If Trim(fsNombreActual) <> Trim(oPersona.NombreCompleto) Then
            Call oCambio.RegistroDatosCambiosDatosPrinc(False, cMovCambio, oPersona.PersCodigo, Trim(fsNombreActual), Trim(oPersona.NombreCompleto))
            bCabecera = True
        End If
        
        For nAux = 1 To nZ
            If nI >= nAux Then
                nTpoDocAnt = CInt(Trim(Right(rsDocPersActual!tipo, 5)))
                cNumDocAnt = Trim(rsDocPersActual!Numero)
                rsDocPersActual.MoveNext
            Else
                nTpoDocAnt = 0
                cNumDocAnt = ""
            End If
            
            If nJ >= nAux Then
                nTpoDocAct = CInt(Trim(Right(rsDocPersUlt!tipo, 5)))
                cNumDocAct = Trim(rsDocPersUlt!Numero)
                rsDocPersUlt.MoveNext
            Else
                nTpoDocAct = 0
                cNumDocAct = ""
            End If
            
            If (nTpoDocAnt <> nTpoDocAct) Or (cNumDocAnt <> cNumDocAct) Then
                Call oCambio.RegistroDatosCambiosDatosPrinc(True, cMovCambio, oPersona.PersCodigo, , , nTpoDocAnt, nTpoDocAct, cNumDocAnt, cNumDocAct)
                bCabecera = True
            End If
        Next nAux
        
        If bCabecera Then
            Call oCambio.RegistroCabeceraDatosCambiosDatosPrinc(cMovCambio, oPersona.PersCodigo, 1)
        End If
        'FRHU 20151130 ERS077-2015
        If oPersona.AutorizaUsoDatosIni <> oPersona.AutorizaUsoDatos And oPersona.Personeria = gPersonaNat Then
            Dim oNPersona As New COMNPersona.NCOMPersona
            ldFechaHoraDJ = CDate(gdFecSis & " " & Format(Time, "hh:mm:ss"))
            If nTipoForm = 1 Then 'add pti1 ERS070-2018 26/12/2018
                MatPersona(1).sNombres = oPersona.Nombres
                MatPersona(1).sApePat = oPersona.ApellidoPaterno
                MatPersona(1).sApeMat = oPersona.ApellidoMaterno
                MatPersona(1).sApeCas = oPersona.ApellidoCasada
                MatPersona(1).sSexo = oPersona.Sexo
                MatPersona(1).sEstadoCivil = oPersona.EstadoCivil
                MatPersona(1).cNacionalidad = oPersona.Nacionalidad
                MatPersona(1).sDomicilio = oPersona.Domicilio
                MatPersona(1).sRefDomicilio = oPersona.RefDomicilio
                MatPersona(1).sUbicGeografica = oPersona.UbicacionGeografica
                MatPersona(1).sCelular = oPersona.Celular
                MatPersona(1).sTelefonos = oPersona.Telefonos
                MatPersona(1).sEmail = oPersona.Email
            End If
                MatPersona(2).sNombres = oPersona.Nombres
                MatPersona(2).sApePat = oPersona.ApellidoPaterno
                MatPersona(2).sApeMat = oPersona.ApellidoMaterno
                MatPersona(2).sApeCas = oPersona.ApellidoCasada
                MatPersona(2).sSexo = oPersona.Sexo
                MatPersona(2).sEstadoCivil = oPersona.EstadoCivil
                MatPersona(2).cNacionalidad = oPersona.Nacionalidad
                MatPersona(2).sDomicilio = oPersona.Domicilio
                MatPersona(2).sRefDomicilio = oPersona.RefDomicilio
                MatPersona(2).sUbicGeografica = oPersona.UbicacionGeografica
                MatPersona(2).sCelular = oPersona.Celular
                MatPersona(2).sTelefonos = oPersona.Telefonos
                MatPersona(2).sEmail = oPersona.Email
            Call CargaDocumentosParaActAutDatos
            Dim tregisgro As Integer
            Dim nSensibles As Integer
            tregisgro = -1
            If bNuevaPersona = False Or BotonEditar Then
                tregisgro = 2
            Else
                If BotonEditar = False Or bNuevaPersona Then
                tregisgro = 1
                End If
            End If
           
            
            If bSensible Then 'add pti1 ERS070-2018 26/12/2018
             nSensibles = 1
            Else
             nSensibles = 0
            End If
            
            'Call oNPersona.InsertarPersActAutDatos(cMovCambio, 1, oPersona.PersCodigo, MatPersona(), oPersona.AutorizaUsoDatos, 2, 1) 'comentado por pti1 ers070-2018
            Call oNPersona.InsertarPersActAutDatos(cMovCambio, nSensibles, oPersona.PersCodigo, MatPersona(), oPersona.AutorizaUsoDatos, 2, 1, 0, 1, tregisgro) 'ADD PTI1 ERS070-2018
         
            If nTipoForm = 1 Then
             Call ImprimirPdfCartillaAutorizacion
            Else
             Call ImprimirPdfCartilla
             Call ImprimirPdfCartillaAutorizacion
            End If
            
'COMENTADO POR PTI1 ERS070-2018
'            MatPersona(2).sNombres = oPersona.Nombres
'            MatPersona(2).sApePat = oPersona.ApellidoPaterno
'            MatPersona(2).sApeMat = oPersona.Nombres
'            MatPersona(2).sApeCas = oPersona.ApellidoCasada
'            MatPersona(2).sSexo = oPersona.Sexo
'            MatPersona(2).sEstadoCivil = oPersona.EstadoCivil
'            MatPersona(2).cNacionalidad = oPersona.Nacionalidad
'            MatPersona(2).sDomicilio = oPersona.Domicilio
'            MatPersona(2).sRefDomicilio = oPersona.RefDomicilio
'            MatPersona(2).sUbicGeografica = oPersona.UbicacionGeografica
'            MatPersona(2).sCelular = oPersona.Celular
'            MatPersona(2).sTelefonos = oPersona.Telefonos
'            MatPersona(2).sEmail = oPersona.Email
'            Call CargaDocumentosParaActAutDatos
'            Call oNPersona.InsertarPersActAutDatos(cMovCambio, 1, oPersona.PersCodigo, MatPersona(), oPersona.AutorizaUsoDatos, 2, 1)
'            Call ImprimirPdfCartilla
' FIN COMENTADO POR PTI1
        Else 'add else pti1 ers070-2018 14/12/2018
            MatPersona(2).sNombres = oPersona.Nombres
            MatPersona(2).sApePat = oPersona.ApellidoPaterno
            MatPersona(2).sApeMat = oPersona.ApellidoMaterno
            MatPersona(2).sApeCas = oPersona.ApellidoCasada
            MatPersona(2).sSexo = oPersona.Sexo
            MatPersona(2).sEstadoCivil = oPersona.EstadoCivil
            MatPersona(2).cNacionalidad = oPersona.Nacionalidad
            MatPersona(2).sDomicilio = oPersona.Domicilio
            MatPersona(2).sRefDomicilio = oPersona.RefDomicilio
            MatPersona(2).sUbicGeografica = oPersona.UbicacionGeografica
            MatPersona(2).sCelular = oPersona.Celular
            MatPersona(2).sTelefonos = oPersona.Telefonos
            MatPersona(2).sEmail = oPersona.Email
            Call ImprimirPdfCartilla
        End If
        'FIN FRHU 20151130
    End If
    Set rsDocPersActual = Nothing
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN *******************************************
    
    'LUCV20181220, Anexo01 de Acta 199-2018
    Set objPista = New COMManejador.Pista
    lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, IIf(nTipoAccion = 1, gInsertar, gModificar), Me.Caption, TxtBCodPers.Text, gCodigoPersona
    Set objPista = Nothing
    'Fin LUCV20181220
    
    Screen.MousePointer = 0
    
    MsgBox "Datos Grabados", vbInformation, "Aviso"
    Call HabilitaControlesPersona(False)
    CmdPersAceptar.Visible = False
    CmdPersCancelar.value = False
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    NomMoverSSTabs = -1
    
    'Habilita Todos los Controles
    SSTabs.Enabled = True
    'cmdPersRelacNew.Enabled = True
    'cmdPersRelacEditar.Enabled = True
    'cmdPersRelacDel.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdSalir.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Visible = False
    cmdnuevo.Enabled = BotonNuevo
    cmdEditar.Enabled = BotonEditar
    cmdPersFteIngresoEjecutado = 0
    
    bPersonaAct = False
    'FRHU 20151130 ERS077-2015
    'lblAutorizarUsoDatos.Visible = False 'comentado por pti1 14/12/2018 ers070-2018
    'cboAutoriazaUsoDatos.Visible = False 'comentado por pti1 14/12/2018 ers070-2018
    'FIN FRHU 20151130
    If bBuscaNuevo Then
        sPersCodNombre = IIf(oPersona.PersCodigo = "", lsPersCodGrabar, oPersona.PersCodigo) & oPersona.NombreCompleto
        bBuscaNuevo = False
        Unload Me
    End If
    'EJVG20120323
    If bRealizaMantenimiento Then
        bUsuarioRealizoMantenimiento = True
        Unload Me
    End If
    Call TxtBCodPers_EmiteDatos ' ARCV 28-08-2006

    
    Exit Sub
    
ErrorCmdPersAceptar:
    MsgBox Err.Description, vbExclamation, "Aviso"
    Call CmdPersCancelar_Click
End Sub
'EJVG20120815 ***
Private Sub imprimirDJSujetoObligado2(ByVal pdFechaDJ As Date)
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    Dim sArchivo As String
    Dim lsDoc As String
    Dim lnTpoDoc As Integer
    Dim lsTpoDoc As String
    Dim lsNroDoc As String
    Dim lsPersCodRepresentante As String, lsNombreRepresentante As String, lsIdRepresentante As String, lsDirRepresentante As String
    Dim oPer As COMDPersona.DCOMPersonas
    Dim rsPersona As ADODB.Recordset
    Dim lsSI As String, lsNO As String
    
    Set oPer = New COMDPersona.DCOMPersonas
    Set rsPersona = New ADODB.Recordset
    Set oWord = CreateObject("Word.Application")
    
    If oPersona.Personeria = 1 Then
        'Plantilla Natural
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\DJ_SO_NAT.doc")
        lsDoc = oPersona.ObtenerRUC
        If lsDoc = "" Then
            lsDoc = oPersona.ObtenerDNI
            If lsDoc = "" Then
                Call oPersona.ObtenerDatosDocumentoxPos(0, lnTpoDoc, lsTpoDoc, lsNroDoc)
                lsDoc = Trim(Left(lsTpoDoc, Len(lsTpoDoc) - 3)) & ": " & lsNroDoc
            Else
                lsDoc = "DNI: " & lsDoc
            End If
        Else
            lsDoc = "RUC: " & lsDoc
        End If
        With oWord.Selection.Find
            .Text = "<<Identificacion>>"
            .Replacement.Text = lsDoc
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    Else
        'Plantilla Juridica
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\DJ_SO_JUR.doc")
        lsPersCodRepresentante = oPersona.ObtenerCodPersonaRelacionado(10)
        If lsPersCodRepresentante <> "" Then
            Set rsPersona = oPer.RecuperaDatosPersona_Basic(lsPersCodRepresentante)
            If Not RSVacio(rsPersona) Then
                lsNombreRepresentante = rsPersona!cPersNombre
                lsIdRepresentante = IIf(rsPersona!Ruc <> "", "RUC: " & rsPersona!Ruc, "DNI: " & rsPersona!Dni)
                lsDirRepresentante = rsPersona!cPersDireccDomicilio
            End If
            Set rsPersona = Nothing
        End If
        With oWord.Selection.Find
            .Text = "<<NombreRep>>"
            .Replacement.Text = IIf(lsNombreRepresentante <> "", lsNombreRepresentante, "_____________________________________________________________________")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<IdentificacionRep>>"
            .Replacement.Text = IIf(lsIdRepresentante <> "", lsIdRepresentante, "________________")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<DireccionRep>>"
            .Replacement.Text = IIf(lsDirRepresentante <> "", lsDirRepresentante, "______________________________")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<RUC>>"
            .Replacement.Text = oPersona.ObtenerRUC
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    With oWord.Selection.Find
        .Text = "<<Fecha>>"
        .Replacement.Text = Format(pdFechaDJ, "dd/mm/yyyy")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<Nombre>>"
        .Replacement.Text = oPersona.NombreCompleto
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<Direccion>>"
        .Replacement.Text = oPersona.Domicilio
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<ActividadEco>>"
        .Replacement.Text = IIf(oPersona.ActiGiro <> "", oPersona.ActiGiro, "___________________________________")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    If oPersona.OfCumplimiento = 1 Then
        lsSI = "X"
        lsNO = ""
    Else
        lsSI = ""
        lsNO = "X"
    End If
    
    With oWord.Selection.Find
        .Text = "<<SI>>"
        .Replacement.Text = lsSI
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
        .Text = "<<NO>>"
        .Replacement.Text = lsNO
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    sArchivo = App.Path & "\SPOOLER\DJ_" & oPersona.PersCodigo & "_" & Format(pdFechaDJ, "yyyymmddhhmmss") & ".doc"
    oDoc.SaveAs sArchivo

    oDoc.Close

    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True

    Set oDoc = oWord.Documents.Open(sArchivo)

    oWord.Visible = True
    Set oDoc = Nothing
    Set oWord = Nothing
    Set oPer = Nothing
End Sub
'END EJVG *******
'JACA 20110730**********************************************************************
Private Sub imprimirDJSujetoObligado()
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim oRange As Word.Range
    
    Set oWord = CreateObject("Word.Application")
    'oWord.Visible = True
    
    If oPersona.Personeria = 1 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\DJUIFNATURAL.doc")
         With oWord.Selection.Find
            .Text = "<<PersonaNatural>>"
            .Replacement.Text = oPersona.NombreCompleto
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<Dni>>"
            .Replacement.Text = "DNI:" + oPersona.ObtenerDNI
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
         With oWord.Selection.Find
            .Text = "<<DomicilioNatural>>"
            .Replacement.Text = oPersona.Domicilio
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    ElseIf oPersona.Personeria <> 1 Then
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\DJUIFJURIDICA.doc")
        With oWord.Selection.Find
            .Text = "<<PersonaJuridica>>"
            .Replacement.Text = oPersona.NombreCompleto
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<Ruc>>"
            .Replacement.Text = oPersona.ObtenerRUC
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
         With oWord.Selection.Find
            .Text = "<<DomicilioJuridico>>"
            .Replacement.Text = oPersona.Domicilio
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    
    With oWord.Selection.Find
        .Text = "<<Fecha>>"
        .Replacement.Text = Format(gdFecSis, "dd/mm/yyyy")
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
   
   Dim sArchivo As String
   sArchivo = App.Path & "\SPOOLER\DJ_" & oPersona.PersCodigo & ".doc"
   oDoc.SaveAs sArchivo
   
   
    oDoc.Close
    Set oDoc = Nothing
    
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    
    Set oDoc = oWord.Documents.Open(sArchivo)
    
    oWord.Visible = True
    Set oDoc = Nothing
    Set oWord = Nothing
End Sub
'JACA********************************************************************************
Private Sub CmdPersCancelar_Click()
    'Habilita Todos los Controles
    
    '***CTI3 10092018
    Unload frmPersonaJurDatosAdic
    Set frmPersonaJurDatosAdic = Nothing
    
    If bCIIU = False Then
        CargaCIIU
        bCIIU = True
    End If

    NomMoverSSTabs = -1
    SSTabs.Enabled = True
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = bHabilitarBoton 'True
    cmdPersRelacDel.Enabled = bHabilitarBoton 'True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdSalir.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    CmdPersAceptar.Visible = False
    CmdPersCancelar.Visible = False
    
    If Trim(TxtBCodPers.Text) <> "" Then
    
        Call TxtBCodPers_EmiteDatos
    Else
        TxtBCodPers.Text = ""
        Call LimpiarPantalla
        Call HabilitaControlesPersona(False)
        Call HabilitaControlesPersonaFtesIngreso(False)
        Call HabilitaFichaPersonaJur(False)
        Call HabilitaFichaPersonaNat(False)
    End If
    cmdnuevo.Enabled = BotonNuevo
    cmdEditar.Enabled = BotonEditar
    
    If Me.cmdnuevo.Visible And Me.cmdnuevo.Enabled Then
        cmdnuevo.SetFocus
    ElseIf Me.cmdSalir.Visible And Me.cmdSalir.Enabled Then
        cmdSalir.SetFocus
    Else
        If Me.Visible Then Me.SetFocus
    End If
    bPersonaAct = False
    If bBuscaNuevo Then
        sPersCodNombre = ""
        bBuscaNuevo = False
        Unload Me
    End If
    
    cmdVerFirma.Visible = False
    cmdVerFirma.Enabled = False
    'EJVG20111219
    Me.cmdPersIDAceptar.Visible = False
    Me.cmdPersIDCancelar.Visible = False
    Me.cmdPersIDnew.Visible = True
    Me.cmdPersIDedit.Visible = True
    'EJVG20120323
    If bRealizaMantenimiento Then
        oPersona.TipoActualizacion = PersFilaModificada
        TxtBCodPers.Enabled = False
    End If
    'EJVG20120814 ***
    If bNuevaPersona Then
        Me.TxtBCodPers.Enabled = False
    End If
    'END EJVG *******
    'WIOR 20130827 ***************************
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN ********************************
    'FRHU 20151130 ERS077-2015
    'lblAutorizarUsoDatos.Visible = False 'comentado por pti1 ers070-2018 17/12/2018
    'cboAutoriazaUsoDatos.Visible = False 'comentado por pti1 ers070-2018 17/12/2018
   
   'ADD PTI1 ers070-2018 17/12/2018
    Call ValidaAutorizardatos
    CboAutoriazaUsoDatos.Enabled = False
    If oPersona.AutorizaUsoDatos = -1 Then
        'cboAutoriazaUsoDatos.AddItem "", -1
        CboAutoriazaUsoDatos.ListIndex = -1
    End If
    'end  pti1 ers070-2018 17/12/2018
  
    'FIN FRHU 20151130
       
End Sub

Private Sub CmdPersFteConsultar_Click()
    If Trim(FEFteIng.TextMatrix(0, 1)) = "" Then
        MsgBox "No Existen Fuentes de Ingreso para Consultar", vbInformation, "Aviso"
        Exit Sub
    End If
    If oPersona.NumeroFtesIngreso > 0 Then
        If gsProyectoActual <> "C" Then
            Call frmFteIngresos.ConsultarFuenteIngreso(FEFteIng.row - 1, oPersona)
        Else
            Call frmFteIngresosCS.ConsultarFuenteIngreso(FEFteIng.row - 1, oPersona)
        End If
    Else
        MsgBox "No se puede Consultar la Fuente de Ingreso", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdPersIDAceptar_Click()
Dim mensaje As String
    If Not ValidaRestriccion(mensaje) Then
        If Len(mensaje) > 0 Then
            MsgBox mensaje, vbInformation
            Exit Sub
        End If
    End If

    If Not ValidaDatosDocumentos Then
        Exit Sub
    End If
    
    If cmdPersDocEjecutado = 1 Then
        Call oPersona.AdicionaDocumento(PersFilaNueva, FEDocs.TextMatrix(FEDocs.row, 2), FEDocs.TextMatrix(FEDocs.row, 1))
        Call oPersona.ActualizarDocsTipoAct(PersFilaNueva, FEDocs.row - 1)
    Else
        If cmdPersDocEjecutado = 2 Then
            If oPersona.ObtenerDocTipoAct(FEDocs.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarDocsTipoAct(PersFilaModificada, FEDocs.row - 1)
            End If
        End If
    End If
    
    'Tipo de Docmumento
    Call oPersona.ActualizaDocsTipo(FEDocs.TextMatrix(FEDocs.row, 1), FEDocs.row - 1)
    'Tipo de Numero
    Call oPersona.ActualizaDocsNumero(FEDocs.TextMatrix(FEDocs.row, 2), FEDocs.row - 1)
    
    'Habilitar Controles
    cmdPersDocEjecutado = 0
    FEDocsPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEDocs.lbEditarFlex = False
    cmdPersIDnew.Enabled = True
    cmdPersIDnew.Visible = True
    'cmdPersIDedit.Enabled = True
    If bPermisoEditarTodo = True Then 'EJVG20111219
        cmdPersIDedit.Enabled = True
        cmdPersIDDel.Enabled = True
    End If
    'WIOR 20130827 ******************************
    If Not bPermisoEditarTodo Then
        If fbPermisoCargo Then
            cmdPersIDedit.Enabled = True
            cmdPersIDDel.Enabled = True
        End If
    End If
    'WIOR FIN ***********************************
    cmdPersIDedit.Visible = True
    'cmdPersIDDel.Enabled = True
    cmdPersIDAceptar.Enabled = False
    cmdPersIDAceptar.Visible = False
    cmdPersIDCancelar.Enabled = False
    cmdPersIDCancelar.Visible = False
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    'cmdNuevo.Enabled = True 'WIOR 20130827 COMENTÓ
    'cmdEditar.Enabled = True 'WIOR 20130827COMENTÓ
    SSTDatosGen.Enabled = True
    SSTabs.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Enabled = True
        CmdPersCancelar.Visible = True
    End If
    
    If oPersona.TipoActualizacion = PersFilaNueva Then
        'SSTabs.Tab = 3
        'If CmdFteIngNuevo.Enabled Then
        '    CmdFteIngNuevo.SetFocus
        'End If
        CmdPersAceptar.SetFocus
        
    Else
        FEDocs.SetFocus
    End If
    
End Sub

Private Sub cmdPersIDCancelar_Click()

    Call CargaDocumentos
    
    'Habilitar Controles
    FEDocsPersNoMoverdeFila = -1
    cmdPersDocEjecutado = -1
    NomMoverSSTabs = -1
    FEDocs.lbEditarFlex = False
    cmdPersIDnew.Enabled = True
    'cmdPersIDedit.Enabled = True
    'cmdPersIDDel.Enabled = True
    If bPermisoEditarTodo = True Then 'EJVG20111219
        cmdPersIDedit.Enabled = True
        cmdPersIDDel.Enabled = True
    End If
    'WIOR 20130827 ******************************
    If Not bPermisoEditarTodo Then
        If fbPermisoCargo Then
            cmdPersIDedit.Enabled = True
            cmdPersIDDel.Enabled = True
        End If
    End If
    'WIOR FIN ***********************************
    cmdPersIDnew.Visible = True
    cmdPersIDedit.Visible = True
    cmdPersIDDel.Visible = True
    cmdPersIDAceptar.Enabled = False
    cmdPersIDCancelar.Enabled = False
    cmdPersIDAceptar.Visible = False
    cmdPersIDCancelar.Visible = False
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    'cmdNuevo.Enabled = True 'WIOR 20130827 COMENTÓ
    'cmdEditar.Enabled = True 'WIOR 20130827 COMENTÓ
    SSTDatosGen.Enabled = True
    SSTabs.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    If FEDocs.rows <= 2 Then    'ACA MODIFIQUE
        cmdPersIDDel.Enabled = False
    End If
    
    FEDocs.SetFocus

End Sub

Private Sub cmdPersIDDel_Click()
    Dim bEnc As Boolean
    Dim i As Integer
    bEnc = False
    
    'JUEZ 20131007 **********************************
    If oPersona.MotivoActu = "2" Then
        MsgBox "No se puede eliminar ningún documento ya registrado si selecciona como Motivo de Actualización a la Campaña Datos", vbInformation, "Aviso"
        Exit Sub
    End If
    'END JUEZ ***************************************
    If FEDocs.rows <= 2 And Trim(FEDocs.TextMatrix(1, 0)) = "" Then
        MsgBox " No existe ningun documento para eliminar", vbInformation, "Aviso"
        Exit Sub
    End If
    'MADM 20101221
    For i = 1 To FEDocs.rows - 1
    
    If Trim(Right(cmbNacionalidad.Text, 12)) = "" And oPersona.Personeria = gPersonaNat Then
        MsgBox "No es posible eliminar al Documento, Complete la Nacionalidad de la Persona", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If CInt(Right(cmbPersPersoneria.Text, 2)) = gPersonaNat Then
            If Trim(Right(cmbNacionalidad.Text, 12)) = "04028" And Trim(Right(cmbNacionalidad.Text, 12)) <> "" Then
                If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdDNI Then
                    bEnc = True
                    Exit For
                End If
            Else
                If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdExtranjeria And Trim(Right(cmbNacionalidad.Text, 12)) <> "" Then
                    bEnc = True
                    Exit For
                End If
            End If
        Else
                If CInt(Right(Trim(FEDocs.TextMatrix(i, 1)), 2)) = gPersIdRUC And Trim(Right(cmbNacionalidad.Text, 12)) <> "" Then
                    bEnc = True
                    Exit For
                End If
        End If
    Next i
    
    If bEnc And FEDocs.rows = 2 Then
        MsgBox "No es posible Eliminar al Documento", vbInformation, "Aviso"
        Exit Sub
    End If
    'END MADM
    If MsgBox("Esta Seguro que Desea Eliminar este Documento", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            Call oPersona.ActualizarDocsTipoAct(PersFilaEliminda, FEDocs.row - 1)
            Call CmdPersAceptar_Click
            Call CargaDocumentos
        Else
            Call oPersona.EliminarDocumento(CInt(Trim(Right(FEDocs.TextMatrix(FEDocs.row, 1), 2))), Trim(FEDocs.TextMatrix(FEDocs.row, 2)))
            If FEDocs.rows > 1 Then
                Call FEDocs.EliminaFila(FEDocs.row)
            End If
            
        End If
        
        If FEDocs.rows <= 2 Then 'ACA MODIFIQUE
            cmdPersIDDel.Enabled = False
        End If
    
    End If
End Sub

Private Sub cmdPersIDedit_Click()
'JUEZ 20131007 **********************************
If oPersona.MotivoActu = "2" Then
    MsgBox "No se puede editar ningún documento ya registrado si selecciona como Motivo de Actualización a la Campaña Datos", vbInformation, "Aviso"
    Exit Sub
End If
'END JUEZ ***************************************
'MADM 20110203
If Trim(Right(cmbNacionalidad.Text, 12)) = "" And oPersona.Personeria = gPersonaNat Then
        MsgBox "No es posible Editar al Documento, Complete la Nacionalidad de la Persona", vbInformation, "Aviso"
        Exit Sub
End If
'END MADM
    cmdPersDocEjecutado = 2
    FEDocsPersNoMoverdeFila = FEDocs.row
    FEDocs.lbEditarFlex = True
    FEDocs.SetFocus
    cmdPersIDnew.Enabled = False
    cmdPersIDnew.Visible = False
    cmdPersIDedit.Enabled = False
    cmdPersIDedit.Visible = False
    cmdPersIDDel.Enabled = False
    cmdPersIDAceptar.Enabled = True
    cmdPersIDAceptar.Visible = True
    cmdPersIDCancelar.Enabled = True
    cmdPersIDCancelar.Visible = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = False
    cmbPersPersoneria.Enabled = False
    cmdnuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTDatosGen.Enabled = False
    SSTabs.Enabled = False
        
End Sub

Private Sub cmdPersIDnew_Click()
'MADM 20110203
If Trim(Right(cmbNacionalidad.Text, 12)) = "" And oPersona.Personeria = gPersonaNat Then
        MsgBox "No es posible Agregar el Documento, Complete la Nacionalidad de la Persona", vbInformation, "Aviso"
        Exit Sub
End If
'END MADM

    FEDocs.AdicionaFila
    cmdPersDocEjecutado = 1
    FEDocsPersNoMoverdeFila = FEDocs.rows - 1
    FEDocs.lbEditarFlex = True
    FEDocs.SetFocus
    cmdPersIDnew.Enabled = False
    cmdPersIDnew.Visible = False
    cmdPersIDedit.Enabled = False
    cmdPersIDedit.Visible = False
    cmdPersIDDel.Enabled = False
    cmdPersIDAceptar.Enabled = True
    cmdPersIDAceptar.Visible = True
    cmdPersIDCancelar.Enabled = True
    cmdPersIDCancelar.Visible = True
    CmdPersAceptar.Visible = True
    CmdPersCancelar.Visible = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = False
    cmbPersPersoneria.Enabled = False
    cmdnuevo.Enabled = False
    cmdEditar.Enabled = False
    SSTDatosGen.Enabled = False
    SSTabs.Enabled = False
End Sub

Private Sub cmdPersRelacAceptar_Click()
    
    If Len(Trim(FERelPers.TextMatrix(FERelPers.row, 5))) = 0 Then
        FERelPers.TextMatrix(FERelPers.row, 5) = "0.00"
    End If
    
    If Not ValidaDatosPersRelacion Then
        Exit Sub
    End If
    
   If cmdPersRelaEjecutado = 1 Then
        Call oPersona.AdicionaPersonaRelacion
        Call oPersona.ActualizarRelacPersTipoAct(PersFilaNueva, FERelPers.row - 1)
    Else
        If cmdPersRelaEjecutado = 2 Then
            If oPersona.ObtenerRelaPersTipoAct(FERelPers.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRelacPersTipoAct(PersFilaModificada, FERelPers.row - 1)
            End If
        End If
    End If
    
    'Apellidos y Nombres
    Call oPersona.ActualizaPersRelaPersona(FERelPers.TextMatrix(FERelPers.row, 2), FERelPers.TextMatrix(FERelPers.row, 1), FERelPers.row - 1)
    ' Relacion
    Call oPersona.ActualizaPersRelaRelacion(FERelPers.TextMatrix(FERelPers.row, 3), FERelPers.row - 1)
    'Beneficiario
    Call oPersona.ActualizarRelaPersBenef(FERelPers.TextMatrix(FERelPers.row, 4), FERelPers.row - 1)
    'Beneficiario Porcentaje
    Call oPersona.ActualizarRelaPersBenefPorc(CDbl(FERelPers.TextMatrix(FERelPers.row, 5)), FERelPers.row - 1)
    'AMP
    Call oPersona.ActualizarRelaPersAMP(FERelPers.TextMatrix(FERelPers.row, 6), FERelPers.row - 1)
    
    'Habilitar Controles
    cmdPersRelaEjecutado = 0
    FERelPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FERelPers.lbEditarFlex = False
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = bHabilitarBoton 'True
    cmdPersRelacDel.Enabled = bHabilitarBoton ' True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FERelPers.SetFocus
End Sub

Private Sub cmdPersRelacCancelar_Click()
    CargaRelacionesPersonas
    'Habilitar Controles
    FERelPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FERelPers.lbEditarFlex = False
    cmdPersRelacNew.Enabled = True
    cmdPersRelacEditar.Enabled = bHabilitarBoton 'True
    cmdPersRelacDel.Enabled = bHabilitarBoton 'True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    'cmdNuevo.Enabled = True COMENTADO BY APRI 20170630
    'cmdEditar.Enabled = True COMENTADO BY APRI 20170630
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    cmdPersRelacAceptar.Visible = False
    cmdPersRelacCancelar.Visible = False
    
    FERelPers.SetFocus

End Sub

Private Sub cmdPersRelacDel_Click()
    If MsgBox("Esta Seguro que Desea Eliminar La Relacion con esta Persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRelacPersTipoAct(PersFilaEliminda, FERelPers.row - 1)
        Call CmdPersAceptar_Click
        Call CargaRelacionesPersonas
    End If
End Sub

Private Sub cmdPersRelacEditar_Click()
    If oPersona.NumeroRelacPers > 0 Then
        cmdPersRelaEjecutado = 2
        FERelPersNoMoverdeFila = FERelPers.row
        NomMoverSSTabs = SSTabs.Tab
        FERelPers.lbEditarFlex = True
        FERelPers.SetFocus
        cmdPersRelacNew.Enabled = False
        cmdPersRelacEditar.Enabled = False
        cmdPersRelacDel.Enabled = False
        cmdPersRelacAceptar.Visible = True
        cmdPersRelacCancelar.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        FERelPers.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub cmdPersRelacNew_Click()
    cmdPersRelacAceptar.Visible = True
    cmdPersRelacCancelar.Visible = True
    cmdPersRelacNew.Enabled = False
    cmdPersRelacDel.Enabled = bHabilitarBoton 'False
    cmdPersRelacEditar.Enabled = bHabilitarBoton 'False
    FERelPers.lbEditarFlex = True
    FERelPers.AdicionaFila
    cmdPersRelaEjecutado = 1
    FERelPersNoMoverdeFila = FERelPers.rows - 1
    FERelPers.SetFocus
End Sub

Private Sub cmdRefBanAcepta_Click()

    If Not ValidaDatosRefBancaria Then
        Exit Sub
    End If
        
   If cmdPersRefBancariaEjecutado = 1 Then
        Call oPersona.AdicionaRefBancaria
        Call oPersona.ActualizarRefBanTipoAct(PersFilaNueva, feRefBancaria.row - 1)
    Else
        If cmdPersRefBancariaEjecutado = 2 Then
            If oPersona.ObtenerRefBanTipoAct(feRefBancaria.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRefBanTipoAct(PersFilaModificada, feRefBancaria.row - 1)
            End If
        End If
    End If
    
    'Codigo de la Institucion Financiera
    Call oPersona.ActualizaRefBanCodIF(feRefBancaria.TextMatrix(feRefBancaria.row, 1), feRefBancaria.row - 1)
    'Nombre de la Institución Financiera
    Call oPersona.ActualizaRefBanNombre(feRefBancaria.TextMatrix(feRefBancaria.row, 2), feRefBancaria.row - 1)
    'Número de Cuenta de
    Call oPersona.ActualizaRefBanNumCta(feRefBancaria.TextMatrix(feRefBancaria.row, 3), feRefBancaria.row - 1)
    'Número de Tarjeta
    Call oPersona.ActualizaRefBanNumTar(feRefBancaria.TextMatrix(feRefBancaria.row, 4), feRefBancaria.row - 1)
    'Línea de Crédito
    Call oPersona.ActualizaRefBanLinCred(feRefBancaria.TextMatrix(feRefBancaria.row, 5), feRefBancaria.row - 1)
                
    'Habilitar Controles
    cmdPersRefBancariaEjecutado = 0
    FERefBanPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefBancaria.lbEditarFlex = False
    cmdRefBanNuevo.Enabled = True
    cmdRefBanEdita.Enabled = True
    cmdRefBanElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdRefBanAcepta.Visible = False
    cmdRefBanCancela.Visible = False
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    feRefBancaria.SetFocus

End Sub

Private Sub cmdRefBanCancela_Click()
    CargaRefBancarias
    'Habilitar Controles
    FERefBanPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefBancaria.lbEditarFlex = False
    cmdRefBanNuevo.Enabled = True
    cmdRefBanEdita.Enabled = True
    cmdRefBanElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdRefBanAcepta.Visible = False
    cmdRefBanCancela.Visible = False
    
    feRefBancaria.SetFocus

End Sub

Private Sub cmdRefBanEdita_Click()
   If oPersona.NumeroRefBancaria > 0 Then
        cmdPersRefBancariaEjecutado = 2
        FERefBanPersNoMoverdeFila = feRefBancaria.row
        NomMoverSSTabs = SSTabs.Tab
        feRefBancaria.lbEditarFlex = True
        feRefBancaria.SetFocus
        cmdRefBanNuevo.Enabled = False
        cmdRefBanEdita.Enabled = False
        cmdRefBanElimina.Enabled = False
        cmdRefBanAcepta.Visible = True
        cmdRefBanCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        feRefBancaria.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdRefBanElimina_Click()
    If MsgBox("Esta Seguro que Desea Eliminar la Referencia Bancaria", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRefBanTipoAct(PersFilaEliminda, feRefBancaria.row - 1)
        Call CmdPersAceptar_Click
        Call CargaRefBancarias
    End If
End Sub

Private Sub cmdRefBanNuevo_Click()
    cmdRefBanAcepta.Visible = True
    cmdRefBanCancela.Visible = True
    cmdRefBanNuevo.Enabled = False
    cmdRefBanElimina.Enabled = False
    cmdRefBanEdita.Enabled = False
    feRefBancaria.lbEditarFlex = True
    feRefBancaria.AdicionaFila
    cmdPersRefBancariaEjecutado = 1
    FERefBanPersNoMoverdeFila = feRefBancaria.rows - 1
    feRefBancaria.SetFocus
End Sub

Private Sub cmdRefComAcepta_Click()
    
    If Not ValidaDatosRefComercial Then
        Exit Sub
    End If
    
   If cmdPersRefComercialEjecutado = 1 Then
        Call oPersona.AdicionaRefComercial
        Call oPersona.ActualizarRefComTipoAct(PersFilaNueva, feRefComercial.row - 1)
        lnNumRefCom = lnNumRefCom + 1
        feRefComercial.TextMatrix(feRefComercial.row, 6) = lnNumRefCom
    Else
        If cmdPersRefComercialEjecutado = 2 Then
            If oPersona.ObtenerRefComTipoAct(feRefComercial.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarRefComTipoAct(PersFilaModificada, feRefComercial.row - 1)
            End If
        End If
    End If
    
    'Nombre/Razón Social
    Call oPersona.ActualizaRefComNombre(feRefComercial.TextMatrix(feRefComercial.row, 1), feRefComercial.row - 1)
    'Tipo de Referencia Comercial
    Call oPersona.ActualizaRefComTipoRel(feRefComercial.TextMatrix(feRefComercial.row, 2), feRefComercial.row - 1)
    'Comentario
    Call oPersona.ActualizaRefComComentario(feRefComercial.TextMatrix(feRefComercial.row, 3), feRefComercial.row - 1)
    'Telefono
    Call oPersona.ActualizaRefComFono(feRefComercial.TextMatrix(feRefComercial.row, 4), feRefComercial.row - 1)
    'Direccion
    Call oPersona.ActualizaRefComDireccion(feRefComercial.TextMatrix(feRefComercial.row, 5), feRefComercial.row - 1)
        
    Call oPersona.ActualizaRefComCod(feRefComercial.TextMatrix(feRefComercial.row, 6), feRefComercial.row - 1)
            
    'Habilitar Controles
    cmdPersRefComercialEjecutado = 0
    FERefComPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefComercial.lbEditarFlex = False
    cmdRefComNuevo.Enabled = True
    cmdRefComEdita.Enabled = True
    cmdRefComElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmbPersPersoneria.Enabled = True
    cmdRefComAcepta.Visible = False
    cmdRefComCancela.Visible = False
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    feRefComercial.SetFocus

End Sub

Private Sub cmdRefComCancela_Click()
    
    CargaRefComerciales
    'Habilitar Controles
    FERefComPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    feRefComercial.lbEditarFlex = False
    cmdRefComNuevo.Enabled = True
    cmdRefComEdita.Enabled = True
    cmdRefComElimina.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    
    cmdRefComAcepta.Visible = False
    cmdRefComCancela.Visible = False
    
    feRefComercial.SetFocus
    
End Sub

Private Sub cmdRefComEdita_Click()
    If oPersona.NumeroRefComercial > 0 Then
        cmdPersRefComercialEjecutado = 2
        FERefComPersNoMoverdeFila = feRefComercial.row
        NomMoverSSTabs = SSTabs.Tab
        feRefComercial.lbEditarFlex = True
        feRefComercial.SetFocus
        cmdRefComNuevo.Enabled = False
        cmdRefComEdita.Enabled = False
        cmdRefComElimina.Enabled = False
        cmdRefComAcepta.Visible = True
        cmdRefComCancela.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        feRefComercial.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdRefComElimina_Click()
    If MsgBox("Esta Seguro que Desea Eliminar la Referencia Comercial", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarRefComTipoAct(PersFilaEliminda, feRefComercial.row - 1)
        Call CmdPersAceptar_Click
        Call CargaRefComerciales
    End If
End Sub

Private Sub cmdRefComNuevo_Click()
    cmdRefComAcepta.Visible = True
    cmdRefComCancela.Visible = True
    cmdRefComNuevo.Enabled = False
    cmdRefComElimina.Enabled = False
    cmdRefComEdita.Enabled = False
    feRefComercial.lbEditarFlex = True
    feRefComercial.AdicionaFila
    cmdPersRefComercialEjecutado = 1
    FERefComPersNoMoverdeFila = feRefComercial.rows - 1
    feRefComercial.SetFocus
End Sub

Private Sub cmdsalir_Click()
    If bPersonaAct Then
        MsgBox "Grabe o Cancele los Cambios Antes de Salir", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oPersona = Nothing
    bConsultaVerFirma = False 'WIOR 20130410
    Unload Me
    Set frmFteIngresos = Nothing 'FRHU 20150224 ERS158-2014
    
    '***CTI3 10092018
    Unload frmPersonaJurDatosAdic
    Set frmPersonaJurDatosAdic = Nothing
End Sub

Private Sub cmdVentasAceptar_Click()
    If Not ValidaDatosPersVentas Then
        Exit Sub
    End If
    
    If cmdPersVentasEjecutado = 1 Then
        Call oPersona.AdicionaVentas
        Call oPersona.ActualizarVentasTipoAct(PersFilaNueva, FEVentas.row - 1)
    Else
        If cmdPersVentasEjecutado = 2 Then
            If oPersona.ObtenerVentasTipoAct(FEVentas.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarVentasTipoAct(PersFilaModificada, FEVentas.row - 1)
            End If
        End If
    End If
    
    Call oPersona.ActualizarVentasPersCod(FEVentas.TextMatrix(FEVentas.row, 1), FEVentas.row - 1)
    Call oPersona.ActualizarVentasApellNombres(FEVentas.TextMatrix(FEVentas.row, 2), FEVentas.row - 1)
    
    Call oPersona.ActualizarVentasMonto(FEVentas.TextMatrix(FEVentas.row, 3), FEVentas.row - 1)
    Call oPersona.ActualizarVentasFecha(FEVentas.TextMatrix(FEVisitas.row, 4), FEVentas.row - 1)
    Call oPersona.ActualizarVentasPeriodo(FEVentas.TextMatrix(FEVentas.row, 5), FEVentas.row - 1)
    'Habilitar Controles
    cmdPersVentasEjecutado = 0
    FEVentasPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEVentas.lbEditarFlex = False
    Me.cmdVentasNuevo.Enabled = True
    Me.cmdVentasEditar.Enabled = True
    Me.cmdVentasEliminar.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FEVentas.SetFocus
End Sub

Private Sub cmdVentasCancelar_Click()
    CargaVentasPersonas
    'Habilitar Controles
    FEVentasPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEVentas.lbEditarFlex = False
    Me.cmdVentasNuevo.Enabled = True
    Me.cmdVentasEditar.Enabled = True
    Me.cmdVentasEliminar.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    cmdVentasAceptar.Visible = False
    cmdVentasCancelar.Visible = False
    
    FEVentas.SetFocus
End Sub

Private Sub cmdVentasEditar_Click()
If oPersona.NumeroVentasPers > 0 Then
        cmdPersVentasEjecutado = 2
        FEVentasPersNoMoverdeFila = FEVentas.row
        NomMoverSSTabs = SSTabs.Tab
        FEVentas.lbEditarFlex = True
        FEVentas.SetFocus
        cmdVentasNuevo.Enabled = False
        cmdVentasEditar.Enabled = False
        cmdVentasEliminar.Enabled = False
        cmdVentasAceptar.Visible = True
        cmdVentasCancelar.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        FEVentas.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdVentasEliminar_Click()
    If MsgBox("Esta Seguro que Desea Eliminar las ventas perteneciente a esta Persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarVentasTipoAct(PersFilaEliminda, FEVentas.row - 1)
        Call CmdPersAceptar_Click
        Call CargaVentasPersonas
    End If
End Sub

Private Sub cmdVentasNuevo_Click()
    cmdVentasAceptar.Visible = True
    cmdVentasCancelar.Visible = True
    cmdVentasNuevo.Enabled = False
    cmdVentasEliminar.Enabled = False
    cmdVentasEditar.Enabled = False
    FEVentas.lbEditarFlex = True
    FEVentas.AdicionaFila
    cmdPersVentasEjecutado = 1
    FEVentasPersNoMoverdeFila = FEVentas.rows - 1
    FEVentas.SetFocus
End Sub

Private Sub cmdVerFirma_Click()
    If TxtBCodPers.Text = "" Then
        MsgBox "Debe registrar a la persona primero", vbInformation, "Mensaje"
        Exit Sub
    End If
    Call frmPersonaFirma.Inicio(oPersona.PersCodigo, oPersona.sCodage, False, True)
End Sub

Private Sub cmdVisitasAceptar_Click()
   
    If Not ValidaDatosPersVisita Then
        Exit Sub
    End If
    
   If cmdPersVisitasEjecutado = 1 Then
        Call oPersona.AdicionaVisita
        Call oPersona.ActualizarVisitasTipoAct(PersFilaNueva, FEVisitas.row - 1)
    Else
        If cmdPersVisitasEjecutado = 2 Then
            If oPersona.ObtenerVisitasTipoAct(FEVisitas.row - 1) <> PersFilaNueva Then
                Call oPersona.ActualizarVisitasTipoAct(PersFilaModificada, FEVisitas.row - 1)
            End If
        End If
    End If
    'Call oPersona.ActualizarVisitaPersCod(FEVisitas.TextMatrix(FEVisitas.Row, 2), FEVisitas.Row - 1)
    Call oPersona.ActualizarVisitaApellNombres(FEVisitas.TextMatrix(FEVisitas.row, 2), FEVisitas.row - 1)
    Call oPersona.ActualizarVisitaPersCod(FEVisitas.TextMatrix(FEVisitas.row, 1), FEVisitas.row - 1)
    Call oPersona.ActualizarVisitaDireccion(FEVisitas.TextMatrix(FEVisitas.row, 3), FEVisitas.row - 1)
    Call oPersona.ActualizarVisitaFecha(FEVisitas.TextMatrix(FEVisitas.row, 4), FEVisitas.row - 1)
    Call oPersona.ActualizarVisitaUsual(FEVisitas.TextMatrix(FEVisitas.row, 5), FEVisitas.row - 1)
    Call oPersona.ActualizarVisitaObserva(FEVisitas.TextMatrix(FEVisitas.row, 6), FEVisitas.row - 1)
    
    'Habilitar Controles
    cmdPersVisitasEjecutado = 0
    FEVisitasPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEVisitas.lbEditarFlex = False
    Me.cmdVisitasNuevo.Enabled = True
    Me.cmdVisitasEditar.Enabled = True
    Me.cmdVisitasEliminar.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    FEVisitas.SetFocus
End Sub

Private Sub cmdVisitasCancelar_Click()
 CargaVisitasPersonas
    'Habilitar Controles
    FEVisitasPersNoMoverdeFila = -1
    NomMoverSSTabs = -1
    FEVisitas.lbEditarFlex = False
    Me.cmdVisitasNuevo.Enabled = True
    Me.cmdVisitasEditar.Enabled = True
    Me.cmdVisitasEliminar.Enabled = True
    
    'DesHabilitar Controles
    TxtBCodPers.Enabled = True
    cmdnuevo.Enabled = True
    cmdEditar.Enabled = True
    SSTDatosGen.Enabled = True
    SSTIdent.Enabled = True
    
    If oPersona.ExistenCambios Then
        CmdPersAceptar.Enabled = True
        CmdPersCancelar.Enabled = True
    End If
    cmdVisitasAceptar.Visible = False
    cmdVisitasCancelar.Visible = False
    
    FEVisitas.SetFocus
End Sub

Private Sub cmdVisitasEditar_Click()
If oPersona.NumeroVisitaPers > 0 Then
        cmdPersVisitasEjecutado = 2
        FEVisitasPersNoMoverdeFila = FEVisitas.row
        NomMoverSSTabs = SSTabs.Tab
        FEVisitas.lbEditarFlex = True
        FEVisitas.SetFocus
        cmdVisitasNuevo.Enabled = False
        cmdVisitasEditar.Enabled = False
        cmdVisitasEliminar.Enabled = False
        cmdVisitasAceptar.Visible = True
        cmdVisitasCancelar.Visible = True
        CmdPersAceptar.Visible = True
        CmdPersCancelar.Visible = True
        'DesHabilitar Controles
        TxtBCodPers.Enabled = False
        cmbPersPersoneria.Enabled = False
        cmdnuevo.Enabled = False
        cmdEditar.Enabled = False
        SSTDatosGen.Enabled = False
        SSTIdent.Enabled = False
        
        FEVisitas.SetFocus
    Else
        MsgBox "No Existe Datos para Editar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdVisitasEliminar_Click()
  If MsgBox("Esta Seguro que Desea Eliminar La Visita perteneciente a esta Persona", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Call oPersona.ActualizarVisitasTipoAct(PersFilaEliminda, FEVisitas.row - 1)
        Call CmdPersAceptar_Click
        Call CargaVisitasPersonas
    End If
End Sub

Private Sub cmdVisitasNuevo_Click()
    cmdVisitasAceptar.Visible = True
    cmdVisitasCancelar.Visible = True
    cmdVisitasNuevo.Enabled = False
    cmdVisitasEliminar.Enabled = False
    cmdVisitasEditar.Enabled = False
    FEVisitas.lbEditarFlex = True
    FEVisitas.AdicionaFila
    cmdPersVisitasEjecutado = 1
    FEVisitasPersNoMoverdeFila = FEVisitas.rows - 1
    FEVisitas.SetFocus
End Sub

Private Sub FEDocs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If FEDocs.col = 2 Then
            If cmdPersIDAceptar.Enabled Then
                cmdPersIDAceptar.SetFocus
            End If
        End If
    End If
End Sub

Private Sub FEDocs_RowColChange()
    If cmdPersDocEjecutado = 1 Or cmdPersDocEjecutado = 2 Then
        FEDocs.row = FEDocsPersNoMoverdeFila
    End If
End Sub

Private Sub fePatInmuebles_RowColChange()
      
    If cmdPersPatVehicularEjecutado = 1 Or cmdPersPatVehicularEjecutado = 2 Then
        fePatInmuebles.row = FEPatVehPersNoMoverdeFila
    End If

End Sub

Private Sub fePatOtros_RowColChange()
    If cmdPersPatVehicularEjecutado = 1 Or cmdPersPatVehicularEjecutado = 2 Then
        fePatOtros.row = FEPatVehPersNoMoverdeFila
    End If

End Sub

Private Sub fePatVehicular_RowColChange()
    
    If fePatVehicular.col = 2 Then
        If fePatVehicular.TextMatrix(fePatVehicular.row, 2) <> "" Then
            If CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)) < 1000 Or CInt(fePatVehicular.TextMatrix(fePatVehicular.row, 2)) > Year(gdFecSis) Then
                MsgBox "Año de Fabricación no válido", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
    
    If cmdPersPatVehicularEjecutado = 1 Or cmdPersPatVehicularEjecutado = 2 Then
        fePatVehicular.row = FEPatVehPersNoMoverdeFila
    End If

End Sub

Private Sub feRefBancaria_OnCellChange(pnRow As Long, pnCol As Long)
    If feRefBancaria.col = 5 Then
        If feRefBancaria.TextMatrix(feRefBancaria.row, 5) <> "" Then
            If CCur(feRefBancaria.TextMatrix(feRefBancaria.row, 5)) < 0 Then
                MsgBox "Linea de Crédito no puede ser negativa", vbInformation, "Aviso"
                feRefBancaria.TextMatrix(feRefBancaria.row, 5) = 0
            End If
        End If
    End If
End Sub

Private Sub feRefBancaria_RowColChange()
    
    If cmdPersRefBancariaEjecutado = 1 Or cmdPersRefBancariaEjecutado = 2 Then
        feRefBancaria.row = FERefBanPersNoMoverdeFila
    End If
End Sub

Private Sub feRefComercial_RowColChange()
    
    If cmdPersRefComercialEjecutado = 1 Or cmdPersRefComercialEjecutado = 2 Then
        feRefComercial.row = FERefComPersNoMoverdeFila
    End If
End Sub

Private Sub FERelPers_Click()
    Call FERelPers_RowColChange
End Sub


Private Sub FERelPers_EnterCell()
    FERelPers_RowColChange
End Sub

Private Sub FERelPers_RowColChange()
'Dim oConstante As DConstante
Dim oConstante As COMDConstantes.DCOMConstantes
    If FERelPers.lbEditarFlex Then
        If FERelPersNoMoverdeFila <> -1 Then
            FERelPers.row = FERelPersNoMoverdeFila
        End If
        'Set oConstante = New DConstante
        Set oConstante = New COMDConstantes.DCOMConstantes
        Select Case FERelPers.col
            Case 3 'Relacion de Persona
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
            Case 4 'Beneficiario
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacBenef)
            Case 6 'AMP
                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacAmp)
        End Select
        Set oConstante = Nothing
    End If
End Sub

Private Sub FEVentas_RowColChange()
'Dim oConstante As DConstante
Dim oConstante As COMDConstantes.DCOMConstantes
    If FEVentas.lbEditarFlex Then
        If FEVentasPersNoMoverdeFila <> -1 Then
            FEVentas.row = FEVentasPersNoMoverdeFila
        End If
        'Set oConstante = New DConstante
        Set oConstante = New COMDConstantes.DCOMConstantes
        Select Case FEVentas.col
'            Case 3 'Relacion de Persona
'                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacion)
'            Case 4 'Beneficiario
'                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacBenef)
'            Case 6 'AMP
'                FERelPers.CargaCombo oConstante.RecuperaConstantes(gPersRelacAmp)
        End Select
        Set oConstante = Nothing
    End If
End Sub

Private Sub FEVisitas_Click()
Call FEVisitas_RowColChange
End Sub

Private Sub FEVisitas_EnterCell()
FEVisitas_RowColChange
End Sub

Private Sub FEVisitas_RowColChange()
Dim oPers As COMDPersona.DCOMPersonas
    If FEVisitas.lbEditarFlex Then
        If FEVisitasPersNoMoverdeFila <> -1 Then
            FEVisitas.row = FEVisitasPersNoMoverdeFila
        End If
        Set oPers = New COMDPersona.DCOMPersonas
        Select Case FEVisitas.col
            Case 5
                FEVisitas.CargaCombo oPers.Cargar_AlternativaUsual(gsCodCMAC)
        End Select
        Set oPers = Nothing
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    Set RsCIIUTemp = New ADODB.Recordset
    RsCIIUTemp.CursorLocation = adUseClient
    
    Screen.MousePointer = 11
    If gsProyectoActual = "C" Then
        lblMagnitudEmpresarial.Visible = True
        cmbPersJurMagnitud.Visible = False
    Else
        lblMagnitudEmpresarial.Visible = False
        cmbPersJurMagnitud.Visible = True
    End If
    Call CargaControles
    TxtBCodPers.Enabled = True
    cmdEditar.Enabled = False
    Screen.MousePointer = 0
    bNuevaPersona = False
    
    bPersonaAGrabar = True
    bCIIU = True
    
    ' habilitar para conulta de fuentes de ingresos
    'PTI120170530 según ERS014-2017
 If nTipoInicioFteIngreso = 1 Then
        SSTabs.TabVisible(0) = False
        SSTabs.TabVisible(1) = False
        SSTabs.TabVisible(2) = False
        SSTabs.TabVisible(3) = True
        SSTabs.TabVisible(4) = False
        SSTabs.TabVisible(5) = False
        SSTabs.TabVisible(6) = False
        SSTabs.TabVisible(7) = False
        SSTabs.TabVisible(8) = False
        SSTabs.TabVisible(9) = False
'       SSTDatosGen.TabVisible(0) = False
'       SSTDatosGen.Visible = False
'       SSTDatosGen.Enabled = False
'       SSTDatosGen.TabVisible(1) = False
'       SSTDatosGen.Enabled = False
        SSTIdent.TabVisible(0) = False
        SSTIdent.Enabled = False
        lblPersNac.Visible = False
        lblPersNac.Enabled = False
        txtPersNacCreac.Visible = False
        txtPersNacCreac.Enabled = False
        Label16.Visible = False
        Label16.Enabled = False
        txtPersFallec.Visible = False
        txtPersFallec.Enabled = False
        Label11.Visible = False
        Label11.Enabled = False
        TxtSbs.Visible = False
        TxtSbs.Enabled = False
        FEDocs.Visible = False
        frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        Frame4.Visible = False
        Frame5.Visible = False
        cmdPersIDAceptar.Visible = False
        cmdPersIDedit.Visible = False
        cmdPersIDDel.Visible = False
        CmdPersAceptar.Visible = False
        CmdPersCancelar.Visible = False
        lblPEPS.Visible = False
        lblSujetoObligado.Visible = False
        lblOfCumplimiento.Visible = False
        lblAutorizarUsoDatos.Visible = False
        cboPEPS.Visible = False
        cboSujetoObligado.Visible = False
        cboOfCumplimiento.Visible = False
        CboAutoriazaUsoDatos.Visible = False
        cmdSalir.Visible = False
        cmdnuevo.Visible = False
        cmdEditar.Visible = False
        CdlgImg.Visible = False
        lblFecUltAct.Visible = False
        Label8.Visible = False
        cboMotivoActu.Visible = False
        lblBUsuario.Visible = False
        txtBUsuario.Visible = False
        txtFecUltAct.Visible = False
        frmPersona.Width = 10510
        frmPersona.Height = 3500
        
        If codigoOpcionMenu <> "" Then
        'TxtBCodPers.codigoOpcionMenu = codigoOpcionMenu
    End If

' END PTI120170530 según ERS014-2017
        
    Else
        SSTabs.TabVisible(3) = False 'LUCV20160820, Según ERS004-2016
        SSTabs.TabVisible(5) = False 'LUCV20160820, Según ERS004-2016
        SSTabs.TabVisible(4) = False 'ARCV 25-10-2006
        SSTabs.TabVisible(9) = False 'FRHU 20150311 ERS013-2015
    End If
    
    '*** PEAC 20080715
    SSTabs.TabVisible(7) = False
    'ALPA 20081021**************
    Call LlenarComboTipoDocumento
    fbPermisoEditarSujetoObligadoDJ = TienePermisoEditarSujetoObligadoDJ(gsCodCargo) 'EJVG20120815
    
    'WIOR 20130827 ***************************
    Set rsDocPersActual = Nothing
    Set rsDocPersUlt = Nothing
    fbPermisoCargo = False
    'WIOR FIN ********************************
    'JUEZ 20131007 *************************************************
    Dim oGen As New COMDConstSistema.DCOMGeneral
    txtBUsuario.lbUltimaInstancia = False
    txtBUsuario.psRaiz = "USUARIOS RESPONSABLE DE LA INFORMACIÓN"
    txtBUsuario.rs = oGen.GetUsuariosArbol("", "")
    'END JUEZ ******************************************************
    
    'APRI20170630 TI-ERS025
    Set oHabilitar = New COMDConstSistema.DCOMConstSistema
    bHabilitarBoton = oHabilitar.HabilitarBotonEditarEliminarRelacion(gsCodCargo)
    'END APRI
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'bEstadoCargando = False
    If nCodPersAuto = 99 Then
    
    End If
    If bPersonaAct Then
        MsgBox "Grabe o Cancele los Cambios Antes de Salir", vbInformation, "Aviso"
        Cancel = 1
    End If
    'WIOR 20130410 ************************
    If Cancel = 0 Then
        bConsultaVerFirma = False
    End If
    'WIOR FIN *****************************
    Set frmFteIngresos = Nothing 'FRHU 20150224 ERS158-2014
End Sub

Private Sub lblPersCIIU_Click()
    If bCIIU = False Then
        CargaCIIU
        bCIIU = True
    End If
End Sub

Private Sub SSTabs_Click(PreviousTab As Integer)
    If NomMoverSSTabs > -1 Then
        SSTabs.Tab = NomMoverSSTabs
    End If
    'Nuevos Controles
    
Dim bPrimerTab As Boolean
    bPrimerTab = IIf(SSTabs.Tab = 0, True, False)
    lblPersNombreAP.Visible = bPrimerTab
    txtPersNombreAP.Visible = bPrimerTab
    lblPersNombreAM.Visible = bPrimerTab
    txtPersNombreAM.Visible = bPrimerTab
    lblPersNombreN.Visible = bPrimerTab
    txtPersNombreN.Visible = bPrimerTab
    'JACA 20110426********************************************
'   If Not Nothing Is oPersona Then
'    If oPersona.Sexo = "F" Then
'        If Not lblApCasada.Visible = False Then
'        lblApCasada.Visible = bPrimerTab 'True
'        txtApellidoCasada.Visible = bPrimerTab 'True
'        End If
'    Else
'        lblApCasada.Visible = False
'        txtApellidoCasada.Visible = False
'    End If
'   End If
   'JACA END*************************************************
End Sub

Private Sub SSTIdent_LostFocus()
    If cmdPersIDAceptar.Visible Then
        MsgBox "Acepte o Cancele el Ingreso del Documento", vbInformation, "Aviso"
        cmdPersIDAceptar.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtActiGiro_KeyPress(KeyAscii As Integer)
  KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    If KeyAscii = 13 Then
        'If Me.cboocupa.Enabled Then
            Me.cboCargos.SetFocus '** Juez 20120326
        'Else
        '    Me.cboCargos.SetFocus
        'End If
        
    End If
End Sub

Private Sub TxtApellidoCasada_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ApellidoCasada = Trim(txtApellidoCasada.Text)
    End If
End Sub

Private Sub TxtApellidoCasada_GotFocus()
    fEnfoque txtApellidoCasada
End Sub

Private Sub txtApellidoCasada_KeyPress(KeyAscii As Integer)
    'KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    KeyAscii = SoloLetras2(KeyAscii, True) 'WIOR 20120705
    If KeyAscii = 13 Then
        'txtPersNombreN.SetFocus
        If txtPersNombreN.Enabled Then 'EJVG20120120
            txtPersNombreN.SetFocus
        End If
    End If
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If
    If bCIIU = False Then
        If Not RsCIIUTemp Is Nothing Then
            RsCIIUTemp.MoveFirst
            CargaCIIU
            bCIIU = True
        End If
        
    End If
    Call LimpiarPantalla
    Call HabilitaControlesPersona(False)
    Call HabilitaControlesPersonaFtesIngreso(False)
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona   'DPersona
    
    'oPersona.SCodAge = gsCodAge
    
    'Call oPersona.RecuperaPersona(Trim(TxtBCodPers.Text))
    
    'If oPersona.PersCodigo = "" Then
    '    MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    
    bValidaCampDatos = False 'JUEZ 20131024
    'Call CargaDatos
    If Cargar_Datos_Persona(Trim(TxtBCodPers.Text)) = False Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    End If
    Call LlenarComboTipoDocumento 'ALPA 20080922
    If SSTabs.Enabled And SSTabs.Visible Then
        SSTabs.SetFocus
    End If
    
    If bConsultaVerFirma Then
        cmdVerFirma.Visible = True
        cmdVerFirma.Enabled = True
        
        cmdActualizarFirma.Visible = False
        cmdActualizarFirma.Enabled = False
    End If
    
    'EJVG20120814 ***
    If oPersona.SujetoObligado = -1 Then
        Me.lblOfCumplimiento.Visible = False
        Me.cboOfCumplimiento.Visible = False
    Else
        If oPersona.SujetoObligado = 1 Then
            Me.lblOfCumplimiento.Visible = True
            Me.cboOfCumplimiento.Visible = True
        End If
    End If
    'END EJVG *******
    
    'LUCV20181220, Anexo01 de Acta 199-2018
    If (nTipoAccion = 3 Or nTipoInicioFteIngreso = 1) Then 'Pista para la accion de consulta
        Set objPista = New COMManejador.Pista
        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, Me.Caption, TxtBCodPers.Text, gCodigoPersona
        Set objPista = Nothing
    End If
    'Fin LUCV20181220
    
End Sub

Function Cargar_Datos_Persona(pcPersCod As String) As Boolean
    
    Set oPersona = Nothing
    Set oPersona = New UPersona_Cli ' COMDPersona.DCOMPersona
    oPersona.sCodage = gsCodAge
    
    Cargar_Datos_Persona = True
    
    Call oPersona.RecuperaPersona(pcPersCod, , gsCodUser)
    
    If oPersona.PersCodigo = "" Then
        Cargar_Datos_Persona = False
        Exit Function
    End If
    'add pti1 16/01/2018 acta 199-2108 LUPA
    nIngPromedio = ""
    nIngPromedio = oPersona.IngresoPromedio
    'FIN PTI1
'With oPersona
'    fnTipoPersona = .Personeria
'    fsUbicGeografica = .UbicacionGeografica
'    fsDomicilio = .Domicilio
'    fsCondicionDomic = .CondicionDomicilio
'    fnValComDomicilio = .ValComDomicilio
'    fsTipoSangre = .TipoSangre
'    fsApePat = .ApellidoPaterno
'    fsApeMat = .ApellidoMaterno
'    fsNombres = .Nombres
'    fsNombreCompleto = .NombreCompleto
'    fnTalla = .Talla
'    fnPeso = .Peso
'    fsEmail = .Email
'    fsTelefonos2 = .Telefonos2
'    fsSexo = .Sexo
'    fsApeCas = .ApellidoCasada
'    fsEstadoCivil = .EstadoCivil
'    fnHijos = .Hijos
'    fcNacionalidad = .Nacionalidad
'    fnResidencia = .Residencia
'    fdFechaNac = .FechaNacimiento
'    fsTelefonos = .Telefonos
'    fsCiiu = .CIIU
'    fsEstado = .Estado
'    fsSiglas = .Siglas
'    fsPersCodSbs = .PersCodSbs
'    fsTipoPersJur = .TipoPersonaJur
'    fnPersRelInst = .PersRelInst
'    fsMagnitudEmp = .MagnitudEmpresarial
'    fnNumEmplead = .NumerosEmpleados
'    fnMaxRefCom = .MaxRefComercial
'    fnMaxPatVeh = .MaxPatVehicular
'    Set fRfirma = .RFirma
'    fsPersCod = .PersCodigo
'    fnNumFtes = .NumeroFtesIngreso
'    fnNumPatVeh = .NumeroPatVehicular
'    fnNumRefCom = .NumeroRefComercial
'    fsCodAge = .SCodAge
'    fsActualiza = .CampoActualizacion
'    fdFechaHoy = .dFechaHoy
'    fnTipoAct = .TipoActualizacion
'    fnNumDocs = .NumeroDocumentos
'End With
    
    Call CargaDatos
End Function

Private Sub LimpiarPantalla()
Dim i As Integer
    
    bEstadoCargando = True
    'TxtBCodPers.Text = ""
    cmbPersPersoneria.ListIndex = -1
    txtPersNombreAP.Text = ""
    txtPersNombreAM.Text = ""
    txtPersNombreN.Text = ""
    cmbPersNatSexo.ListIndex = -1
    
    cmbPersNatEstCiv.ListIndex = -1
    txtPersNatHijos.Text = "0"
    '*** PEAC 20080412
    txtPersNatNumEmp.Text = "0"
    
    txtPersNacCreac.Text = "__/__/____"
    '*** PEAC 20080412
    txtPersFecInscRuc.Text = "__/__/____"
    txtPersFecIniActi.Text = "__/__/____"
    
    txtPersTelefono.Text = ""
    CboPersCiiu.ListIndex = -1
    
    '*** PEAC 20080412
    cboTipoComp.ListIndex = -1
    cboTipoSistInfor.ListIndex = -1
    cboCadenaProd.ListIndex = -1
    
    cboMonePatri.ListIndex = -1
    '*** FIN PEAC
    
    cmbPersEstado.ListIndex = -1
    TxtTalla.Text = "0.00"
    TxtPeso.Text = "0.00"
    CboTipoSangre.ListIndex = -1
    '** Juez 20120327 ***********
    'chkResidente.value = 0
    cboResidente.ListIndex = 0
    cboPaisReside.ListIndex = -1
    '** End Juez ****************
    'EJVG20120813 ***
    Me.cboPEPS.ListIndex = -1
    Me.cboSujetoObligado.ListIndex = -1
    Me.cboOfCumplimiento.ListIndex = -1
    'END EJVG *******
    txtPersTelefono2.Text = ""
    TxtEmail.Text = ""
    
    'EJVG20111207*****
    txtCel1.Text = ""
    txtCel2.Text = ""
    txtCel3.Text = ""
    TxtEmail2.Text = ""
    '*****************

    '*** PEAC 20080412
    txtNumDependi.Text = ""
    txtActComple.Text = ""
    txtNumPtosVta.Text = ""
    txtActiGiro.Text = ""
    '*** FIN PEAC
    
    'MADM 20091116
    'txtcargo.Text = ""
    'txtcentro.Text = ""
    txtIngresoProm.Text = "0.00"
    chkaho.value = 0
    chkcred.value = 0
    chkotro.value = 0
    'chkResidente.value = 0
    'MADM 20100322
    cboocupa.ListIndex = -1
    'END MADM
    cboCargos.ListIndex = -1 '** Juez 20120326
    
    Call LimpiaFlex(FEDocs)
    Me.TxtCodCIIU.Text = ""
    cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), "04028")
    cmbPersUbiGeo(1).ListIndex = -1
    cmbPersUbiGeo(2).ListIndex = -1
    cmbPersUbiGeo(3).ListIndex = -1
    cmbPersUbiGeo(4).ListIndex = -1
    txtPersDireccDomicilio.Text = ""
    txtRefDomicilio.Text = "" '*** PEAC 20080801
    cmbPersDireccCondicion.ListIndex = -1
    txtValComercial.Text = ""
    'JUEZ 20131007 ***********************
    cmbNegUbiGeo(0).ListIndex = IndiceListaCombo(cmbNegUbiGeo(0), "04028")
    cmbNegUbiGeo(1).ListIndex = -1
    cmbNegUbiGeo(2).ListIndex = -1
    cmbNegUbiGeo(3).ListIndex = -1
    cmbNegUbiGeo(4).ListIndex = -1
    txtNegDireccion.Text = ""
    txtRefNegocio.Text = ""
    txtNombreCentroLaboral.Text = "" 'marg
    txtBUsuario.Text = ""
    'END JUEZ ****************************
    txtPersNombreRS.Text = ""
    TxtSiglas.Text = ""
    cmbPersJurTpo.ListIndex = -1
    cmbPersJurMagnitud.ListIndex = -1
    txtPersJurObjSocial.Text = "" '** Juez 20120328
    cmbPersNatMagnitud.ListIndex = -1 'JACA 20110428
    lblMagnitudEmpresarial.Caption = ""
    txtPersJurEmpleados.Text = ""
    TxtSbs.Text = ""
    CmbRela.ListIndex = -1
    CboAutoriazaUsoDatos.ListIndex = -1 'ADD PTI1 ers070-2018
    Call LimpiaFlex(FERelPers)
    Call LimpiaFlex(FEFteIng)
    Call LimpiaFlex(feRefComercial)
    IDBFirma.RutaImagen = ""
    bEstadoCargando = False
    
End Sub

Private Sub TxtBCodPers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbPersPersoneria.Enabled Then
            cmbPersPersoneria.SetFocus
        Else
            'If SSTabs.TabVisible(0) Then
            '    SSTabs.Tab = 0
            'Else
            '    SSTabs.Tab = 1
            'End If
            If txtPersNombreAP.Enabled Then
                txtPersNombreAP.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtCel1_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Celular = Trim(txtCel1.Text)
    End If
End Sub
'ORCR20140318 INICIO ***
Private Sub txtCel1_KeyPress(KeyAscii As Integer)
    If txtCel1.Text = "" Then
        If Not DigitoRPM(KeyAscii) Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                If (validarTelefonoControl(txtCel1)) Then
                   txtCel2.SetFocus
                End If
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            If (validarTelefonoControl(txtCel1)) Then
                txtCel2.SetFocus
            End If
        End If
    End If
End Sub
'ORCR20140318 FIN ******

Private Sub txtCel2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Celular2 = Trim(txtCel2.Text)
    End If
End Sub
'ORCR20140318 INICIO ***
Private Sub txtCel2_KeyPress(KeyAscii As Integer)
    If txtCel2.Text = "" Then
        If Not DigitoRPM(KeyAscii) Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                If (validarTelefonoControl(txtCel2)) Then
                   txtCel3.SetFocus
                End If
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            If (validarTelefonoControl(txtCel2)) Then
                txtCel3.SetFocus
            End If
        End If
    End If
End Sub
'ORCR20140318 FIN ***

Private Sub txtCel3_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Celular3 = Trim(txtCel3.Text)
    End If
End Sub
'ORCR20140318 INICIO ***
Private Sub txtCel3_KeyPress(KeyAscii As Integer)
    If txtCel3.Text = "" Then
        If Not DigitoRPM(KeyAscii) Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                If (validarTelefonoControl(txtCel3)) Then
                   cboRemInfoEmail.SetFocus 'JUEZ 20131007
                End If
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            If (validarTelefonoControl(txtCel3)) Then
                cboRemInfoEmail.SetFocus 'JUEZ 20131007
            End If
        End If
    End If
End Sub
'ORCR20140318 FIN ***

Private Sub TxtCodCIIU_KeyPress(KeyAscii As Integer)
    Dim ObjP As COMDPersona.DCOMPersonas
    Dim rs As ADODB.Recordset
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 And Trim(Me.TxtCodCIIU.Text) <> "" Then
         'txtPersTelefono.SetFocus
         Set ObjP = New COMDPersona.DCOMPersonas
         Set rs = New ADODB.Recordset
         Set rs = ObjP.Get_CIIU_Busqueda(CStr(Me.TxtCodCIIU.Text))
         'CboPersCiiu.Clear
        If rs.RecordCount > 0 Then
            'CboPersCiiu.Text = Trim(rs!cCIIUdescripcion) & Space(100) & Trim(rs!cCIIUcod)
            CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, rs!cCIIUcod) '** Juez 20120409 ******
        Else
            MsgBox "Código no existe.", vbInformation, "Aviso"
        End If
'         While Not rs.EOF
'            CboPersCiiu.AddItem Trim(rs!cCIIUdescripcion) & Space(100) & Trim(rs!cCIIUcod)
'            rs.MoveNext
'         Wend
         
         'Call Llenar_Combo_con_Recordset(Rs, CboPersCiiu)
'         If Not (rs.EOF And rs.BOF) Then
'            CboPersCiiu.ListIndex = 0
'            CboPersCiiu.SetFocus
'         End If
         
         Set rs = Nothing
         Set ObjP = Nothing
         bCIIU = False
         
    End If
End Sub

Private Sub TxtEmail_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Email = Trim(TxtEmail.Text)
    End If
End Sub

Private Sub TxtEmail2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Email2 = Trim(TxtEmail2.Text)
    End If
End Sub

Private Sub TxtEmail2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboPersCiiu.SetFocus
    End If
End Sub

'*** PEAC 20080412
Private Sub txtnumdependi_Change()
           
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NumDependi = CInt(IIf(Trim(txtNumDependi.Text) = "", "0", Trim(txtNumDependi.Text)))
        
    End If
End Sub
'*** PEAC 20080412
Private Sub txtactcomple_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ActComple = Trim(txtActComple.Text)
    End If
End Sub

Private Sub txtNumDependi_GotFocus()
    fEnfoque txtNumDependi
End Sub

Private Sub txtNumDependi_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    
End Sub

Private Sub txtNumDependi_LostFocus()
    txtNumDependi.Text = IIf(Trim(txtNumDependi.Text) = "", "0", Trim(txtNumDependi.Text))
End Sub

'*** PEAC 20080412
Private Sub txtnumptosvta_Change()
       
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NumPtosVta = CInt(IIf(Trim(txtNumPtosVta.Text) = "", "0", Trim(txtNumPtosVta.Text)))
   
    End If
End Sub

'*** PEAC 20080412
Private Sub txtactigiro_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ActiGiro = Trim(txtActiGiro.Text)
    End If
End Sub


Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'CboPersCiiu.SetFocus
        TxtEmail2.SetFocus
    End If
End Sub


'CUSCO
Private Sub txtIngresoProm_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        If txtIngresoProm.Text = "" Then txtIngresoProm.Text = "0"
        oPersona.IngresoPromedio = CDbl(txtIngresoProm.Text)
    End If
End Sub
Private Sub txtIngresoProm_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtIngresoProm, KeyAscii)
'If KeyAscii = 13 Then TxtEmail.SetFocus
If KeyAscii = 13 Then EnfocaControl txtCel1 '.SetFocus
End Sub
'''''''''''''''''''''

Private Sub txtNumPtosVta_GotFocus()
    fEnfoque txtNumPtosVta
End Sub

Private Sub txtNumPtosVta_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtNumPtosVta_LostFocus()
    txtNumPtosVta.Text = IIf(Trim(txtNumPtosVta.Text) = "", "0", Trim(txtNumPtosVta.Text))
End Sub

Private Sub txtPersDireccDomicilio_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Domicilio = Trim(txtPersDireccDomicilio.Text)
    End If
End Sub

Private Sub txtPersDireccDomicilio_GotFocus()
    fEnfoque txtPersDireccDomicilio
End Sub

Private Sub txtPersDireccDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersDireccCondicion.SetFocus
    End If
End Sub

Private Sub txtPersJurEmpleados_Change()

    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NumerosEmpleados = CInt(IIf(Trim(txtPersJurEmpleados.Text) = "", "0", Trim(txtPersJurEmpleados.Text)))
    End If
End Sub

Private Sub txtPersJurEmpleados_GotFocus()
    fEnfoque txtPersJurEmpleados
End Sub

Private Sub txtPersJurEmpleados_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 And cmbPersJurMagnitud.Visible Then
        cmbPersJurMagnitud.SetFocus
    End If
End Sub

Private Sub txtPersJurEmpleados_LostFocus()
    txtPersJurEmpleados.Text = IIf(Trim(txtPersJurEmpleados.Text) = "", "0", Trim(txtPersJurEmpleados.Text))
End Sub

Private Sub txtPersNacCreac_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
        If Len(Trim(ValidaFecha(txtPersNacCreac.Text))) = 0 Then
            oPersona.FechaNacimiento = CDate(txtPersNacCreac.Text)
        End If
    End If
    
End Sub

'MAVM 06042009
Private Sub txtPersFallec_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
        If Len(Trim(ValidaFecha(txtPersFallec.Text))) = 0 Then
            oPersona.FechaFallecimiento = CDate(txtPersFallec.Text)
        Else
            oPersona.FechaFallecimiento = CDate("01/01/1900")
        End If
    End If
End Sub

'*** PEAC 20080412
Private Sub txtPersFecInscRuc_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
        If Len(Trim(ValidaFecha(txtPersFecInscRuc.Text))) = 0 Then
            oPersona.FechaInscRuc = CDate(txtPersFecInscRuc.Text)
        End If
    End If
    
End Sub

'*** PEAC 20080412
Private Sub txtPersFecIniActi_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        
        If Len(Trim(ValidaFecha(txtPersFecIniActi.Text))) = 0 Then
            oPersona.FechaIniActi = CDate(txtPersFecIniActi.Text)
        End If
    End If
End Sub

Private Sub txtPersNacCreac_GotFocus()
    fEnfoque txtPersNacCreac
End Sub

'*** PEAC 20080412
Private Sub txtPersFecInscRuc_GotFocus()
    fEnfoque txtPersFecInscRuc
End Sub

'*** PEAC 20080412
Private Sub txtPersFecIniActi_GotFocus()
    fEnfoque txtPersFecIniActi
End Sub


Private Sub txtPersNacCreac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtPersTelefono.SetFocus
    End If
End Sub

Private Sub txtPersNacCreac_LostFocus()
Dim sCad As String

    sCad = ValidaFecha(txtPersNacCreac.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        If txtPersNacCreac.Enabled Then txtPersNacCreac.SetFocus
        Exit Sub
    End If
    If CDate(txtPersNacCreac.Text) >= gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtPersNacCreac.SetFocus
        Exit Sub
    End If

End Sub

'*** PEAC 20080412
Private Sub txtPersFecInscRuc_LostFocus()
Dim sCad As String

    sCad = ValidaFecha(txtPersFecInscRuc.Text)
'    If Not Trim(sCad) = "" Then
'        MsgBox sCad, vbInformation, "Aviso"
'        If txtPersFecInscRuc.Enabled Then txtPersFecInscRuc.SetFocus
'        Exit Sub
'    End If
    
    If Not Trim(sCad) <> "" Then
        If CDate(txtPersFecInscRuc.Text) >= gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtPersFecInscRuc.SetFocus
            Exit Sub
        End If
    End If
End Sub

'*** PEAC 20080412
Private Sub txtPersFecIniActi_LostFocus()
Dim sCad As String

    sCad = ValidaFecha(txtPersFecIniActi.Text)
'    If Not Trim(sCad) = "" Then
'        MsgBox sCad, vbInformation, "Aviso"
'        If txtPersFecIniActi.Enabled Then txtPersFecIniActi.SetFocus
'        Exit Sub
'    End If

If Not Trim(sCad) <> "" Then
    If CDate(txtPersFecIniActi.Text) >= gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtPersFecIniActi.SetFocus
        Exit Sub
    End If
End If

End Sub

Private Sub txtPersNatHijos_Change()
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Hijos = CInt(Trim(IIf(Trim(txtPersNatHijos.Text) = "", "0", txtPersNatHijos.Text)))
      End If
End Sub
'*** PEAC 20080412
Private Sub txtPersNatNumEmp_Change()
      
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NumEmp = CInt(Trim(IIf(Trim(txtPersNatNumEmp.Text) = "", "0", txtPersNatNumEmp.Text)))
      End If
      
End Sub


Private Sub txtPersNatHijos_GotFocus()
    fEnfoque txtPersNatHijos
End Sub

'*** PEAC 20080412
Private Sub txtPersNatNumEmp_GotFocus()
    fEnfoque txtPersNatNumEmp
End Sub

Private Sub txtPersNatHijos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        cmbNacionalidad.SetFocus
    End If
End Sub

'*** PEAC 20080412
Private Sub txtPersNatNumEmp_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtPersNatHijos_LostFocus()
    txtPersNatHijos.Text = IIf(Trim(txtPersNatHijos.Text) = "", "0", txtPersNatHijos.Text)
End Sub

'*** PEAC 20080412
Private Sub txtPersNatNumEmp_LostFocus()
    txtPersNatNumEmp.Text = IIf(Trim(txtPersNatNumEmp.Text) = "", "0", txtPersNatNumEmp.Text)
End Sub

Private Sub txtPersNombreAM_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
        oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.ApellidoMaterno = Trim(txtPersNombreAM.Text)
   End If
End Sub

Private Sub txtPersNombreAM_GotFocus()
    fEnfoque txtPersNombreAM
End Sub

Private Sub txtPersNombreAM_KeyPress(KeyAscii As Integer)
    'KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    KeyAscii = SoloLetras2(KeyAscii, True) 'WIOR 20120705
    If KeyAscii = 13 Then
        If txtApellidoCasada.Visible Then
            txtApellidoCasada.SetFocus
        Else
            txtPersNombreN.SetFocus
        End If
    End If
End Sub

Private Sub txtPersNombreAP_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ApellidoPaterno = Trim(txtPersNombreAP.Text)
    End If
End Sub

Private Sub txtPersNombreAP_GotFocus()
    fEnfoque txtPersNombreAP
End Sub

Private Sub txtPersNombreAP_KeyPress(KeyAscii As Integer)
    'KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    KeyAscii = SoloLetras2(KeyAscii, True) 'WIOR 20120705
    If KeyAscii = 13 Then
        txtPersNombreAM.SetFocus
    End If
End Sub

Private Sub txtPersNombreN_Change()
  If Not bEstadoCargando Then
    If oPersona.TipoActualizacion <> PersFilaNueva Then
      oPersona.TipoActualizacion = PersFilaModificada
    End If
    oPersona.Nombres = Trim(txtPersNombreN.Text)
  End If
End Sub

Private Sub txtPersNombreN_GotFocus()
    fEnfoque txtPersNombreN
End Sub

Private Sub txtPersNombreN_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras(KeyAscii, True) 'Letras(KeyAscii)
    If KeyAscii = 13 Then
        cmbPersNatSexo.SetFocus
    End If
End Sub

Private Sub txtPersNombreRS_Change()

    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NombreCompleto = Trim(txtPersNombreRS.Text)
    End If
End Sub

Private Sub txtPersNombreRS_GotFocus()
    fEnfoque txtPersNombreRS
End Sub

Private Sub txtPersNombreRS_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        TxtSiglas.SetFocus
    End If
End Sub

Private Sub txtPersTelefono_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Telefonos = Trim(txtPersTelefono.Text)
    End If
End Sub
'ORCR20140318 INICIO ***
'Private Sub txtPersTelefono_GotFocus()
'    fEnfoque txtPersTelefono
'End Sub
Private Sub txtPersTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If (validarTelefonoControl(txtPersTelefono)) Then
           txtPersTelefono2.SetFocus
        End If
    End If
End Sub
Private Function validarTelefonoControl(ctrControl As Control) As Boolean
    Dim X As New RegExp
    X.Pattern = "^(\*|\#)?\d+$"
    
    ctrControl.Text = RTrim(LTrim(ctrControl.Text))
    
    If Len(ctrControl.Text) = 0 Then
        validarTelefonoControl = True
        Exit Function
    ElseIf Not (X.Test(ctrControl.Text)) Then
        MsgBox "El Telefono debe contener solo numeros ó comenzar con # ó *", vbInformation, "Aviso"
        ctrControl.SetFocus
        validarTelefonoControl = False
        Exit Function
    Else
        Dim Numero As String
        Dim digito As String
        Dim flat As Boolean
        Dim n As Integer
        n = 6
        
        Numero = ctrControl.Text
        digito = Mid(Numero, 1, 1)
        
        If digito = "#" Or digito = "*" Then
            digito = (Mid(Numero, 2, 1))
            n = n + 1
        End If
        
        If (Len(Numero) < n) Then
            MsgBox "El Telefono deben tener por lo menos 6 digitos", vbInformation, "Aviso"
            ctrControl.SetFocus
            validarTelefonoControl = False
            Exit Function
        End If
        
        X.Pattern = "^(\*|\#)?" + digito + "+$"
        
        flat = Not (X.Test(ctrControl.Text))
        
        If Not flat Then
            MsgBox "No todos los digitos deben ser Iguales", vbInformation, "Aviso"
            ctrControl.SetFocus
        End If
        
        validarTelefonoControl = flat
    End If
End Function
'ORCR20140318 FIN ***

'RECO20140217 ERS160-2013******************************************************
Private Sub txtPersTelefono_LostFocus()
    If Len(Trim(txtPersTelefono.Text)) > 0 Then
        txtCel1.BackColor = 16777215
        txtPersTelefono.BackColor = 12648447
    Else
        txtCel1.BackColor = 12648447
        txtPersTelefono.BackColor = 16777215
    End If
    If Len(Trim(txtPersTelefono.Text)) > 0 And Len(Trim(txtCel1.Text)) > 0 Then
        txtPersTelefono.BackColor = 12648447
        txtCel1.BackColor = 12648447
    End If
End Sub
'RECO FIN************************************************************************

Private Sub txtPersTelefono2_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Telefonos2 = Trim(txtPersTelefono2.Text)
    End If
End Sub
'ORCR20140318 INICIO ***
Private Sub txtPersTelefono2_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If (validarTelefonoControl(txtPersTelefono2)) Then
            txtIngresoProm.SetFocus
        End If
    End If
End Sub
'ORCR20140318 FIN ***

Private Sub TxtPeso_Change()
    'On Error Resume Next
      If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Peso = CDbl(Format(IIf(Trim(TxtPeso.Text) = "", "0", TxtPeso.Text), "#0.00"))
      End If
End Sub

Private Sub TxtPeso_GotFocus()
    fEnfoque TxtPeso
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtPeso, KeyAscii, 5, 2)
    If KeyAscii = 13 Then
        If bPermisoEditarTodo Then  'EJVG20110105
            txtPersNacCreac.SetFocus
        Else
            txtPersTelefono.SetFocus
        End If
    End If
End Sub

Private Sub TxtPeso_LostFocus()
    If Trim(TxtPeso.Text) = "." Then
        TxtPeso.Text = "0.00"
    End If
    TxtPeso.Text = Format(IIf(Trim(TxtPeso.Text) = "", "0.00", TxtPeso.Text), "#0.00")
End Sub

Private Sub txtRefDomicilio_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RefDomicilio = Trim(txtRefDomicilio.Text)
    End If
End Sub

Private Sub txtRefDomicilio_GotFocus()
    fEnfoque txtRefDomicilio
End Sub
Private Sub txtRefDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        'CmdPersAceptar.SetFocus
        cmbNegUbiGeo(0).SetFocus 'JUEZ 20131007
    End If
End Sub

Private Function ValidarCodSBS(ByVal sCodSbs As String) As Boolean
    Dim oCOMDPersona As COMDPersona.DCOMPersonas
    Dim lrsCodSBS As ADODB.Recordset
    Set oCOMDPersona = New COMDPersona.DCOMPersonas

    Set lrsCodSBS = oCOMDPersona.ValidarCodSBS(sCodSbs)

    If lrsCodSBS.RecordCount <> "0" Then
       ValidarCodSBS = True
    Else
        ValidarCodSBS = False
    End If
End Function

'JUEZ 20131007 **********************************************
Private Sub txtNegDireccion_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NegocioDireccion = Trim(txtNegDireccion.Text)
    End If
End Sub

Private Sub txtNegDireccion_GotFocus()
    fEnfoque txtNegDireccion
End Sub

Private Sub txtNegDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtRefNegocio.SetFocus
    End If
End Sub



Private Sub txtRefNegocio_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.RefNegocio = Trim(txtRefNegocio.Text)
    End If
End Sub

' marg 11-05-2016
Private Sub txtNombreCentroLaboral_Change()
If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.NombreCentroLaboral = Trim(txtNombreCentroLaboral.Text)
    End If
End Sub

Private Sub txtRefNegocio_GotFocus()
    fEnfoque txtRefNegocio
End Sub

Private Sub txtNombreCentroLaboral_GotFocus()
    fEnfoque txtNombreCentroLaboral
End Sub

Private Sub txtRefNegocio_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        txtNombreCentroLaboral.SetFocus
    End If
End Sub

Private Sub txtNombreCentroLaboral_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        CmdPersAceptar.SetFocus
    End If
End Sub
'END JUEZ ***************************************************

Private Sub TxtSbs_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.PersCodSbs = Trim(TxtSbs.Text)
    End If
End Sub

Private Sub TxtSbs_GotFocus()
    fEnfoque TxtSbs
End Sub

Private Sub TxtSbs_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        If ValidarCodSBS(TxtSbs.Text) = True Then
            MsgBox "El Codigo SBS ya Existe!!", vbCritical, Me.Caption
            TxtSbs.SetFocus
        Else
            txtPersTelefono.SetFocus
        End If
    End If
End Sub

Private Sub TxtSiglas_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Siglas = Trim(TxtSiglas.Text)
    End If
End Sub

Private Sub txtSiglas_GotFocus()
    fEnfoque TxtSiglas
End Sub

Private Sub txtSiglas_KeyPress(KeyAscii As Integer)
     KeyAscii = Letras(KeyAscii)
    If KeyAscii = 13 Then
        Me.txtPersJurObjSocial.SetFocus
    End If
End Sub

Private Sub TxtTalla_Change()
    'On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.Talla = CDbl(IIf(Trim(TxtTalla.Text) = "", "0.00", Trim(TxtTalla.Text)))
    End If
End Sub

Private Sub TxtTalla_GotFocus()
    fEnfoque TxtTalla
End Sub

Private Sub TxtTalla_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(TxtTalla, KeyAscii, 4, 2)
    If KeyAscii = 13 Then
        CboTipoSangre.SetFocus
    End If
End Sub

Private Sub TxtTalla_LostFocus()
    If TxtTalla.Text = "." Then
        TxtTalla.Text = "0.00"
    End If
    TxtTalla.Text = Format(IIf(Trim(TxtTalla.Text) = "", "0.00", Trim(TxtTalla.Text)), "#0.00")
End Sub

Private Sub txtValComercial_Change()
    'On Error Resume Next
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.ValComDomicilio = txtValComercial.Text
    End If
End Sub

Private Sub txtValComercial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cmdPersIDnew.SetFocus
        txtRefDomicilio.SetFocus 'JUEZ 20131007
    End If
End Sub

Public Function ValidarNombresApellidos(ByVal sNombreApeC As String, ByVal nLargoNombreApeC As Integer, ByRef nSalir As Integer, ByVal nTipo) As Integer
Dim nI As Integer
Dim nJ As Integer
Dim sNombreApe As String
Dim nLargoNombreApe As Integer

sNombreApe = sNombreApeC
nLargoNombreApe = nLargoNombreApeC
If nSalir = 0 Then
For nI = 1 To nLargoNombreApe
    For nJ = 192 To 254
        If Mid(sNombreApe, nI, 1) = Chr(nJ) And ((nJ >= 192 And nJ <= 208) Or (nJ >= 210 And nJ <= 219) Or (nJ >= 221 And nJ <= 240) Or (nJ >= 242 And nJ <= 251) Or (nJ >= 253 And nJ <= 254)) Then
            If nTipo = 1 Then
                nSalir = 1
            ElseIf nTipo = 2 Then
                nSalir = 2
            ElseIf nTipo = 3 Then
                nSalir = 3
            End If
            Exit For
        End If
    Next nJ
    If nSalir = 1 Then
        Exit For
    End If
Next nI
End If
ValidarNombresApellidos = nSalir
End Function

Public Sub LlenarComboTipoDocumento()
Dim i As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
    'FEDocs.Clear
    With rs
        ''Crear RecordSet
        .Fields.Append "cConsDescripcion", adVarChar, 100
        .Fields.Append "nConsValor", adDouble
        .Open
     If nNumeroDoc > 0 Then
        For i = 1 To nNumeroDoc
            If (cmbPersPersoneria.Text) <> "" Then
              If CInt(Right(cmbPersPersoneria.Text, 2)) = 2 Or CInt(Right(cmbPersPersoneria.Text, 2)) = 3 Then
                     If MatrixTipoDoc(2, i) = 2 Then
                            .AddNew
                            .Fields("nConsValor") = MatrixTipoDoc(2, i)
                            .Fields("cConsDescripcion") = MatrixTipoDoc(1, i)
                     End If
              Else
                    If Trim(Right(cmbNacionalidad.Text, 6)) = "04028" Then
                        If MatrixTipoDoc(3, i) = 1 Or MatrixTipoDoc(3, i) = 0 Then
                            .AddNew
                            .Fields("nConsValor") = MatrixTipoDoc(2, i)
                            .Fields("cConsDescripcion") = MatrixTipoDoc(1, i)
                         End If
                     Else
                         If MatrixTipoDoc(3, i) = 2 Or MatrixTipoDoc(3, i) = 0 Then
                            .AddNew
                            .Fields("nConsValor") = MatrixTipoDoc(2, i)
                            .Fields("cConsDescripcion") = MatrixTipoDoc(1, i)
                         End If
                     End If
               End If
            Else
                If Trim(Right(cmbNacionalidad.Text, 6)) = "04028" Then
                    If MatrixTipoDoc(3, i) = 1 Or MatrixTipoDoc(3, i) = 0 Then
                        .AddNew
                        .Fields("nConsValor") = MatrixTipoDoc(2, i)
                        .Fields("cConsDescripcion") = MatrixTipoDoc(1, i)
                     End If
                 Else
                     If MatrixTipoDoc(3, i) = 2 Or MatrixTipoDoc(3, i) = 0 Then
                        .AddNew
                        .Fields("nConsValor") = MatrixTipoDoc(2, i)
                        .Fields("cConsDescripcion") = MatrixTipoDoc(1, i)
                     End If
                 End If
            End If
         Next i
     End If
     End With
     rs.MoveFirst
     FEDocs.CargaCombo rs
End Sub

Private Function BuscaNumDocEnListaNegativa(ByVal pnTipoDoc As Integer, ByVal psNumDoc As String, Optional ByRef pnCondicion As Integer = 0) As Boolean 'WIOR 20121122 AGREGO  pnCondicion
    Dim ObjP As COMDPersona.DCOMPersonas
    Dim rs As ADODB.Recordset
       
    BuscaNumDocEnListaNegativa = False ''no encontrado
         
         Set ObjP = New COMDPersona.DCOMPersonas
         Set rs = New ADODB.Recordset
'         Set rs = ObjP.Get_CIIU_Busqueda(CStr(Me.TxtCodCIIU.Text))
         Set rs = ObjP.BusquedaEnListaNegativa(pnTipoDoc, psNumDoc)
        
        pnCondicion = 0 'WIOR 20121122
        If Not (rs.EOF And rs.BOF) Then
            pnCondicion = CInt(rs!nCondicion) 'WIOR 20121122
            Set rs = Nothing
            Set ObjP = Nothing
            BuscaNumDocEnListaNegativa = True ''encontrado
            Exit Function
        End If
        
'        If rs.RecordCount > 0 Then
'            CboPersCiiu.Text = Trim(rs!cCIIUdescripcion) & Space(100) & Trim(rs!cCIIUcod)
'        Else
'            'MsgBox "Código no existe.", vbInformation, "Aviso"
'            BuscaNumDocEnListaNegativa = False
'            Exit Function
'        End If
         
'         Set rs = Nothing
'         Set ObjP = Nothing

'    End If

End Function

Private Function VerificarAutorizacion(ByVal pConcepto As String) As Boolean
Dim oCapAut As COMDPersona.DCOMPersonas
Dim rsx As New ADODB.Recordset
Dim sNombreCompletox As String
Dim cod As String
Dim clongCadena As Integer
Dim bAutorizaPer As Boolean
Dim bPreAutoriacion As Boolean 'WIOR 20121123

 If Right(Me.cmbPersPersoneria.Text, 1) = "1" Then
        If Right(Me.cmbPersNatSexo.Text, 1) = "F" And Len(Trim(txtApellidoCasada)) > 0 Then
            If Right(Me.cmbPersNatEstCiv.Text, 1) = "3" Then
                sNombreCompletox = txtPersNombreAP & "/" & txtPersNombreAM & "\VDA " & txtApellidoCasada & "," & txtPersNombreN
            Else
                sNombreCompletox = txtPersNombreAP & "/" & txtPersNombreAM & "\" & txtApellidoCasada & "," & txtPersNombreN
            End If
        Else
            sNombreCompletox = txtPersNombreAP & "/" & txtPersNombreAM & "," & txtPersNombreN
        End If
  End If

Set oCapAut = New COMDPersona.DCOMPersonas
bAutorizaPer = False

'WIOR 20121123 VALIDA PRE AUTORIZACION ***************************************************************
bPreAutoriacion = False
If Not oCapAut.ValidaInsercionAprobacion(gdFecSis, sNombreCompletox, gsCodUser, False, True) Then
     If Not oCapAut.ValidaInsercionAprobacion(gdFecSis, sNombreCompletox, gsCodUser, True, True) Then
        
        clongCadena = Len(cboocupa.Text)
        Dim sMovNroPre As String
        sMovNroPre = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

         oCapAut.InsertaPersNegativo_Aprobacion gdFecSis, sNombreCompletox, gsCodUser, gsCodAge, Trim(Mid(Me.cboocupa.Text, 1, clongCadena - 5)), , , pConcepto, sMovNroPre
         bAutorizaPer = True
    Else
        bPreAutoriacion = True
    End If
End If
'WIOR FIN ********************************************************************************************

If bPreAutoriacion Then 'WIOR 20121123
    'Valida Insertar Datos
    If Not oCapAut.ValidaInsercionAprobacion(gdFecSis, sNombreCompletox, gsCodUser) Then
         If Not oCapAut.ValidaInsercionAprobacion(gdFecSis, sNombreCompletox, gsCodUser, True) Then
    
            clongCadena = Len(cboocupa.Text)
             
            Dim oCont As COMNContabilidad.NCOMContFunciones  'NContFunciones
            Dim sMovNro As String, sOperacion As String
                    
            Set oCont = New COMNContabilidad.NCOMContFunciones
            sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set oCont = Nothing
             
             oCapAut.InsertaPersNegativo_Aprobacion gdFecSis, sNombreCompletox, gsCodUser, gsCodAge, Trim(Mid(Me.cboocupa.Text, 1, clongCadena - 5)), , , pConcepto, sMovNro
             bAutorizaPer = True
        Else
            VerificarAutorizacion = True
            Set oCapAut = Nothing
            Exit Function
        End If
    End If
    'MADM 20110223 - MENSAJE VerificarAprobacionNivelGerencialPEPS
    If Not bAutorizaPer Then
            If oCapAut.ValidaInsercionAprobacion(gdFecSis, sNombreCompletox, gsCodUser, True) Then
                VerificarAutorizacion = True
                Set oCapAut = Nothing
                Exit Function
            Else
                VerificarAutorizacion = False
                MsgBox "Comuníquese con un Nivel Gerencial para la Aprobación, No podrá continuar hasta que Autoricen la Operación", vbInformation, "Aviso"
            End If
    Else
        MsgBox "Comuníquese con un Nivel Gerencial para la Aprobación, No podrá continuar hasta que Autoricen la Operación", vbInformation, "Aviso"
        VerificarAutorizacion = False
        Exit Function
    End If
'WIOR 20121123 ************************************************************************
Else
    MsgBox "Comuníquese con un Supervisor de Operaciones y/o Jefe de Agencia para la Pre-Autorización de la Persona," & Chr(10) & "No podrá continuar hasta que Autoricen la Operación", vbInformation, "Aviso"
    VerificarAutorizacion = False
    Exit Function
End If
'WIOR FIN *****************************************************************************
        Set oCapAut = Nothing
End Function
'JACA 20110427**********************************************************************
Private Sub cmbPersNatMagnitud_Change()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.MagnitudPersNat = Trim(Right(cmbPersNatMagnitud.Text, 15))
    End If
End Sub
Private Sub cmbPersNatMagnitud_Click()
    If Not bEstadoCargando Then
        If oPersona.TipoActualizacion <> PersFilaNueva Then
            oPersona.TipoActualizacion = PersFilaModificada
        End If
        oPersona.MagnitudPersNat = Trim(Right(cmbPersNatMagnitud.Text, 15))
    End If
End Sub

Private Sub cmbPersNatMagnitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If bPermisoEditarTodo Then
            cmbPersNatSexo.SetFocus
        Else
            txtPersNatHijos.SetFocus
        End If
    End If
End Sub
'JACA END************************************************************************
'EJVG20111219 ***************************************************
Public Function ObtenerVecesCreditoyAhorroPersona(ByVal psPersCod As String) As Integer
    Dim oCred As DCOMCredito
    Set oCred = New DCOMCredito
    ObtenerVecesCreditoyAhorroPersona = oCred.ObtenerVecesCreditoyAhorroPersona(psPersCod)
    Set oCred = Nothing
End Function
Public Function validaPermisoEditarPersona(ByVal psCargo As String, Optional psPersCod As String = "") As Boolean
    Dim oPersona As New DCOMPersonas
    validaPermisoEditarPersona = oPersona.validaPermisoEditarPersona(psCargo, psPersCod)
    Set oPersona = Nothing
End Function
'EJVG20120815 ***
Public Function TienePermisoEditarSujetoObligadoDJ(ByVal psCodCargo As String) As Boolean
    Dim oPersona As DCOMPersonas
    Set oPersona = New DCOMPersonas
    TienePermisoEditarSujetoObligadoDJ = oPersona.TienePermisoEditarSujetoObligadoDJ(psCodCargo)
    Set oPersona = Nothing
End Function
'END EJVG *******
'JUEZ 20131024 **************************************************
Private Function ValidaDetallesCampDatos() As Boolean
    ValidaDetallesCampDatos = False
    
    Dim oNPers As New COMNPersona.NCOMPersona
    If Not oNPers.VerificaCargoMantenimientoCampDatos(gsCodCargo) Then
        MsgBox "Su cargo no puede actualizar la información con la Campaña Datos", vbInformation, "Aviso"
        Screen.MousePointer = 0
        cboMotivoActu.ListIndex = 0
        ValidaDetallesCampDatos = False
        Exit Function
    End If
    Set oNPers = Nothing
    
    Dim oDPers As New COMDPersona.DCOMPersonas
    If Not oDPers.PermiteParticiparCampañaDatos(oPersona.PersCodigo) Then
        MsgBox "El cliente no puede participar en la Campaña Datos por las siguientes razones:" & vbCrLf & _
        "* El cliente es colaborador activo de Caja Maynas" & vbCrLf & _
        "* El cliente no tiene ni tuvo relación directa o indirecta con algún producto de Ahorro o Crédito" & vbCrLf & _
        "* El cliente no tiene ni tuvo relación directa o indirecta con algún producto de Crédito con calificación A, B o C", vbInformation, "Aviso"
        Screen.MousePointer = 0
        cboMotivoActu.ListIndex = 0
        ValidaDetallesCampDatos = False
        Exit Function
    End If
    Set oDPers = Nothing
    ValidaDetallesCampDatos = True
End Function
'END JUEZ *******************************************************
'FRHU 20140401 ERS027-2014 RQ14132
'Este procedimiento se llama desde el formulario frmPersBusqueda
Public Sub ConsultarPorPersona(ByVal psPersCod As String)
    If Cargar_Datos_Persona(psPersCod) = False Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    End If
    TxtBCodPers.Text = psPersCod
    Me.Caption = "Personas:Consulta"
    cmdEditar.Enabled = False
    cmdnuevo.Enabled = False
    BotonNuevo = cmdnuevo.Enabled
    BotonEditar = cmdEditar.Enabled
    bConsultaVerFirma = True
    lblAutorizarUsoDatos.Visible = True 'ADD por pti1 se cambio de false a true ers070-2018
    CboAutoriazaUsoDatos.Visible = True 'ADD por pti1 se cambio de false a true ers070-2018
    nTipoForm = 5 'add pti1 ers070-2018
    Me.Show 1
End Sub
'FIN FRHU 20140401 ERS027-2014
'FRHU 20151130 ERS077-2015
Private Sub ImprimirPdfCartilla() 'modificado por pti1 ers070-2018
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
   ' If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then 'comentado por pti1 ers070-2018 27/12/2018
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then 'add por pti1 ers070-2018 27/12/2018
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    'oDoc.LoadImageFromFile App.path & "\Logo_CajaMaynas_2015.bmp", "Logo"
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'Observacion
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WImage 65, 480, 50, 100, "Logo"
    'oDoc.WTextBox 80, 50, 15, 500, "CARTILLA AUTORIZACIÓN Y ACTUALIZACIÓN DE DATOS PERSONALES", "F2", 11, hCenter 'comentado por pti1 ers077-2018 14/12/2108
    oDoc.WTextBox 80, 50, 15, 500, "CARTILLA DE ACTUALIZACIÓN DE DATOS PERSONALES", "F2", 11, hCenter 'add por pti1 ers077-2018 14/12/2108
    
    If (Day(gdFecSis)) > 9 Then
        oDoc.WTextBox 120, 32, 20, 500, (Day(gdFecSis)), "F1", 9, hLeft 'ADD POR PTI1
    Else
        oDoc.WTextBox 120, 32, 20, 500, "0" & (Day(gdFecSis)), "F1", 9, hLeft
    End If
    If (Month(gdFecSis)) > 9 Then
        oDoc.WTextBox 120, 62, 20, 500, (Month(gdFecSis)), "F1", 9, hLeft
    Else
        oDoc.WTextBox 120, 62, 20, 500, "0" & (Month(gdFecSis)), "F1", 9, hLeft
    End If 'FIN ADD POR PTI1
    oDoc.WTextBox 119, 95, 20, 500, (Year(gdFecSis)), "F1", 9, hLeft 'ADD POR PTI1
    oDoc.WTextBox 120, 20, 20, 500, "______/______/________", "F1", 9, hLeft
    oDoc.WTextBox 130 + nAltura, 20, 20, 50, "Nombre(s):", "F1", 9, hLeft
    oDoc.WTextBox 130 + nAltura, 70, 20, 500, MatPersona(2).sNombres, "F1", 9, hLeft '100 caracteres
    oDoc.WTextBox 130 + nAltura, 70, 20, 500, "____________________________________________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 20, 20, 50, "Apellidos: ", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 70, 20, 400, MatPersona(2).sApePat & " " & MatPersona(2).sApeMat & IIf(Len(MatPersona(2).sApeCas) = 0, "", " " & IIf(MatPersona(2).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(2).sApeCas), "F1", 9, hLeft '64 caracteres
    oDoc.WTextBox 160 + nAltura, 70, 20, 400, "________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 400, 20, 55, "Estado Civil: ", "F1", 9, hLeft
    oDoc.WTextBox 160 + nAltura, 455, 20, 250, Trim(Left(cmbPersNatEstCiv.Text, 23)), "F1", 9, hLeft '23 caracteres
    oDoc.WTextBox 160 + nAltura, 455, 20, 250, "_______________________", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 20, 20, 140, "Documento de Identificación:", "F1", 9, hLeft
    
    oDoc.WTextBox 190 + nAltura, 150, 20, 30, "D.N.I.", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 180, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    If MatPersona(2).sPersIDTpo = "1" Then oDoc.WTextBox 190 + nAltura, 180, 10, 20, "X", "F1", 8, hCenter
    
    oDoc.WTextBox 190 + nAltura, 210, 20, 80, "Carnet Extranjeria", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 290, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    If MatPersona(2).sPersIDTpo = "4" Then oDoc.WTextBox 190 + nAltura, 290, 10, 20, "X", "F1", 8, hCenter
    
    oDoc.WTextBox 190 + nAltura, 320, 20, 50, "Pasaporte", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 370, 10, 20, "", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    If MatPersona(2).sPersIDTpo = "11" Then oDoc.WTextBox 190 + nAltura, 370, 10, 20, "X", "F1", 8, hCenter
    
    oDoc.WTextBox 190 + nAltura, 420, 20, 10, "Nº ", "F1", 9, hLeft
    oDoc.WTextBox 190 + nAltura, 435, 20, 200, Trim(MatPersona(2).sPersIDnro), "F1", 9, hLeft '27 caracteres
    oDoc.WTextBox 190 + nAltura, 435, 20, 200, "___________________________", "F1", 9, hLeft
    oDoc.WTextBox 220 + nAltura, 20, 20, 100, "Dirección de domicilio:", "F1", 9, hLeft
    oDoc.WTextBox 220 + nAltura, 115, 20, 500, Left(Trim(MatPersona(2).sDomicilio), 90), "F1", 9, hLeft '91 caracteres
    oDoc.WTextBox 220 + nAltura, 115, 20, 500, "___________________________________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 250 + nAltura, 20, 20, 50, "Referencia:", "F1", 9, hLeft
    oDoc.WTextBox 250 + nAltura, 70, 20, 500, Left(Trim(MatPersona(2).sRefDomicilio), 100), "F1", 9, hLeft '100 caracteres
    oDoc.WTextBox 250 + nAltura, 70, 20, 500, "____________________________________________________________________________________________________", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 20, 20, 60, "Departamento:", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 85, 20, 130, Left(cmbPersUbiGeo(1).Text, 25), "F1", 9, hLeft '26
    oDoc.WTextBox 280 + nAltura, 85, 20, 130, "__________________________", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 220, 20, 50, "Provincia:", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 265, 20, 130, Left(cmbPersUbiGeo(2).Text, 25), "F1", 9, hLeft '26
    oDoc.WTextBox 280 + nAltura, 265, 20, 130, "__________________________", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 400, 20, 60, "Distrito:", "F1", 9, hLeft
    oDoc.WTextBox 280 + nAltura, 435, 20, 135, Left(cmbPersUbiGeo(3).Text, 25), "F1", 9, hLeft '27
    oDoc.WTextBox 280 + nAltura, 435, 20, 135, "___________________________", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 20, 20, 50, "Celular:", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 70, 20, 120, MatPersona(2).sCelular, "F1", 9, hLeft '24
    oDoc.WTextBox 310 + nAltura, 70, 20, 120, "________________________", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 195, 20, 50, "Teléfono:", "F1", 9, hLeft
    oDoc.WTextBox 310 + nAltura, 240, 20, 120, MatPersona(2).sTelefonos, "F1", 9, hLeft '24
    oDoc.WTextBox 310 + nAltura, 240, 20, 120, "________________________", "F1", 9, hLeft
    oDoc.WTextBox 340 + nAltura, 20, 20, 80, "Correo electrónico:", "F1", 9, hLeft
    oDoc.WTextBox 340 + nAltura, 100, 20, 260, Trim(MatPersona(2).sEmail), "F1", 9, hLeft '52
    oDoc.WTextBox 340 + nAltura, 100, 20, 260, "____________________________________________________", "F1", 9, hLeft

'    inicio comentado por pti1 ers070-2018 14/12/2018
'    sParrafoUno = "Autorización para Recopilación y Tratamiento de Datos: Ley de protección de datos personales - N° 29733 (en adelante la Ley): por el presente documento el cliente " & _
'                  "entrega a la Caja datos personales que lo identifican y/o lo hacen identificable, y que son considerados datos personales conforme a las disposiciones de la Ley y la " & _
'                  "legislación  vigente. Datos personales que la Caja queda autorizada por el cliente a mantenerlos en su(s) base(s) de datos, así como para que sean almacenados, " & _
'                  "sistematizados y utilizados para los fines que se detallan en el presente documento. El cliente también autoriza a la Caja que: (i) la información podrá ser conservada " & _
'                  "por la Caja de forma indefinida e independientemente de la relación contractual que mantenga o no con la Caja, (ii) su información está protegida por las leyes " & _
'                  "aplicables y procedimientos que la Caja tiene implementados para el ejercicio de sus derechos, con el objeto que se evite la alteración, pérdida o acceso de personas " & _
'                  "personales con terceras personas, dentro o fuera del país, vinculadas o no a la Caja, exclusivamente para la tercerización de tratamientos autorizados y de " & _
'                  "conformidad con las medidas de seguridad exigidas por la Ley, v) la Caja transfiera o comparta sus datos personales con empresas vinculadas a la Caja y/o terceros, " & _
'                  "para fines de publicidad, mercadeo y similares, vi) le envíe, a través de mensajes de texto a su teléfono celular (SMS), llamadas telefónicas a su teléfono fijo o celular, " & _
'                  "mensajes de correo electrónico a su correo personal o comunicaciones enviadas a su domicilio, promociones e información relacionada a los servicios y productos " & _
'                  "que la Caja, sus subsidiarias o afiliadas ofrecen directa o indirectamente a través de las distintas asociaciones comerciales que la Caja pueda tener, e inclusive " & _
'                  "requerimientos de cobranza, directamente o a través de terceros, respecto de las deudas que pueda mantener el cliente con la Caja.  Asimismo, conforme a lo " & _
'                  "estipulado en la ley, el cliente tiene conocimiento que cuenta con el derecho de actualizar, incluir, rectificar y suprimir sus datos personales, así como a oponerse a su " & _
'                  "tratamiento para los fines antes indicados. El cliente también conoce que en cualquier momento, puede revocar la presente autorización para tratar sus datos " & _
'                  "personales, lo cual surtirá efectos en un plazo no mayor de 5 días calendario contados desde el día siguiente de recibida la comunicación. La revocación no surtirá " & _
'                  "efecto frente a hechos cumplidos, ni frente al tratamiento que sea necesario para la ejecución de una relación contractual vigente o sus consecuencias legales, ni " & _
'                  "podrá oponerse a tratamientos permitidos por ley. Para ejercer el derecho de revocatoria o cualquier otro que la Ley establezca con relación a sus datos personales, " & _
'                  "el cliente deberá dirigir una comunicación escrita por cualquiera de los canales de atención proporcionados por la Caja conforme a la Ley."
'    oDoc.WTextBox 400, 20, 360, 555, sParrafoUno, "F1", 7, hjustify
'
'    sParrafoDos = "Nota.- El cliente declara que, antes de suscribir el presente documento, ha sido informado que tiene derecho a no proporcionar a la Caja la autorización para el " & _
'                  "tratamiento de sus datos personales y que si no la proporciona la Caja no podrá tratar sus datos personales en la forma explicada en éste documento, lo que no " & _
'                  "impide su uso para la ejecución y cumplimiento de cualquier relación contractual que mantenga el cliente con la Caja."
'    oDoc.WTextBox 540, 20, 60, 555, sParrafoDos, "F1", 7, hjustify
'   fin comentado por pti1 ers070-2018 14/12/2018

     Dim cfecha  As String 'pti1 add
     cfecha = Choose(Month(gdFecSis), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
                                        
                                        
    If (Day(gdFecSis)) > 9 Then
        oDoc.WTextBox 410, 25, 20, 200, (Day(gdFecSis)), "F1", 9, hLeft 'ADD POR PTI1
    Else
        oDoc.WTextBox 410, 25, 20, 200, "0" & (Day(gdFecSis)), "F1", 9, hLeft
    End If
    oDoc.WTextBox 410, 65, 80, 200, (cfecha), "F1", 9, hLeft  ' ADD PTI1
    oDoc.WTextBox 410, 155, 110, 200, Right(Year(gdFecSis), 2), "F1", 9, hLeft ' ADD PTI1
    
    'oDoc.WTextBox 580, 20, 60, 200, ArmaFecha(gdFecSis), "F1", 9, hLeft 'comentado por pti1 ers070-2018
    oDoc.WTextBox 410, 20, 60, 200, "____ de ______________ del 20____", "F1", 9, hLeft 'descomentado por pti1 ers070-2018
    oDoc.WTextBox 490, 20, 60, 50, "Firma:", "F1", 9, hLeft
    oDoc.WTextBox 490, 50, 60, 150, "___________________________", "F1", 9, hLeft
    
    oDoc.WTextBox 420, 200, 80, 70, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    oDoc.WTextBox 635, 300, 60, 100, "Acepto", "F1", 9, hLeft ' INICIO COMENTADO POR PTI1
'    oDoc.WTextBox 650, 300, 60, 150, "Autorizar uso de mis datos", "F1", 9, hLeft
'
'    oDoc.WTextBox 630, 420, 15, 20, "SI", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'    oDoc.WTextBox 630, 440, 15, 20, "NO", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
'
'    If Trim(Right(cboAutoriazaUsoDatos.Text, 2)) = "1" Then
'        oDoc.WTextBox 645, 420, 15, 20, "X", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        oDoc.WTextBox 645, 440, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    Else
'        oDoc.WTextBox 645, 420, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'        oDoc.WTextBox 645, 440, 15, 20, "X", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
'    End If ' CFIN COMENTADO POR ERS070
            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub 'fin modificado pti1 ers070-2018 14/12/2018
'FIN FRHU 20151130
Private Sub ImprimirPdfCartillaAutorizacion() 'add PTI1 ERS070-2018 11/12/2018
    Dim sParrafoUno As String
    Dim sParrafoDos As String
    Dim sParrafoTres As String
    Dim sParrafoCuatro As String
    Dim sParrafoCinco As String
    Dim sParrafoSeis As String
    Dim sParrafoSiete As String
    Dim sParrafoOcho As String
    Dim oDoc As cPDF
    Dim nAltura As Integer
    
    Set oDoc = New cPDF
    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Cartilla Autorización y Actualización de datos personales"
    oDoc.Title = "Cartilla Autorización y Actualización de datos personales"
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartillaAutorizacionActualizacionDeDatos" & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    
    'oDoc.LoadImageFromFile App.path & "\logo_cmacmaynas.bmp", "Logo"
    oDoc.LoadImageFromFile App.Path & "\Logo_2015.jpg", "Logo" 'O
    
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    '<body>
    nAltura = 20
    oDoc.WTextBox 10, 10, 780, 575, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
    'oDoc.WImage 60, 480, 35, 100, "Logo"
    oDoc.WImage 70, 460, 50, 100, "Logo" 'O
    oDoc.WTextBox 90, 50, 15, 500, "AUTORIZACIÓN PARA EL TRATAMIENTO DE DATOS PERSONALES", "F2", 11, hCenter 'agregado por pti1 ers070-2018 05/12/2018
     

  
    oDoc.WTextBox 125, 56, 360, 520, (MatPersona(2).sNombres & " " & MatPersona(2).sApePat & " " & MatPersona(2).sApeMat & IIf(Len(MatPersona(2).sApeCas) = 0, "", " " & IIf(MatPersona(2).sEstadoCivil = "2", "DE", "VDA") & " " & MatPersona(2).sApeCas)), "F1", 11, hjustify
    oDoc.WTextBox 125, 484, 360, 520, (Trim(MatPersona(2).sPersIDnro)), "F1", 11, hjustify
    oDoc.WTextBox 125, 56, 360, 520, ("___________________________________________________________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 481, 360, 520, ("____________"), "F1", 11, hjustify
    oDoc.WTextBox 125, 35, 10, 520, ("Yo, " & String(120, vbTab) & "  con DOI N° " & String(22, vbTab) & ""), "F1", 11, hjustify
   
      sParrafoUno = "autorizo y otorgo por tiempo indefinido, " & String(0.52, vbTab) & "mi consentimiento libre, previo, expreso, inequívoco e informado a" & Chr$(13) & _
                   "la " & String(0.52, vbTab) & "CAJA MUNICIPAL DE AHORRO Y CRÉDITO DE MAYNAS " & String(0.52, vbTab) & "S.A. " & String(0.52, vbTab) & "(en " & String(0.52, vbTab) & "adelante," & String(0.52, vbTab) & " ""LA CAJA""), " & String(0.51, vbTab) & " para " & String(0.51, vbTab) & " el" & Chr$(13) & _
                   "tratamiento de mis datos personales proporcionados " & String(0.7, vbTab) & " en contexto de la contratación de cualquier producto " & Chr$(13) & _
                   "(activo y/o pasivo)" & String(0.52, vbTab) & " o" & String(0.51, vbTab) & " servicio, " & String(0.52, vbTab) & " así " & String(0.52, vbTab) & "como " & String(0.52, vbTab) & "resultado" & String(0.52, vbTab) & "de " & String(0.52, vbTab) & " la suscripción de contratos, " & String(0.52, vbTab) & " formularios, " & String(0.52, vbTab) & " y a los " & Chr$(13) & _
                   "recopilados anteriormente, actualmente y/o por recopilar por " & String(0.52, vbTab) & "LA CAJA. " & String(0.53, vbTab) & "Asimismo, " & String(0.53, vbTab) & "otorgo " & String(0.53, vbTab) & "mi autorización" & Chr$(13) & _
                   "para el envío de información  promocional y/o publicitaria de los servicios y productos que" & String(0.53, vbTab) & " LA CAJA ofrece, " & Chr$(13) & _
                   "a tráves de cualquier medio de comunicación que se considere apropiado para su difusión, " & String(0.53, vbTab) & "y " & String(0.52, vbTab) & "para" & String(0.53, vbTab) & " su uso " & Chr$(13) & _
                   "en la gestión administrativa " & String(0.53, vbTab) & " y " & String(0.5, vbTab) & " comercial de  " & String(0.53, vbTab) & "LA  " & String(0.53, vbTab) & "CAJA " & String(0.53, vbTab) & " que guarde relación con su objeto social.  " & String(0.53, vbTab) & "En " & String(0.52, vbTab) & "ese " & Chr$(13) & _
                   "sentido, autorizo a LA CAJA al uso de mis datos personales para tratamientos que supongan el " & String(0.52, vbTab) & "desarrollo" & Chr$(13) & _
                   "de acciones y actividades comerciales, incluyendo la realización de estudios  de  mercado, " & String(0.53, vbTab) & " elaboración " & String(0.52, vbTab) & "de" & Chr$(13) & _
                   "perfiles de compra " & String(0.53, vbTab) & " y evaluaciones financieras. " & String(0.54, vbTab) & " El uso y tratamiento de mis datos personales, " & String(0.54, vbTab) & "se sujetan" & Chr$(13) & _
                   "a lo establecido por el artículo 13° de la Ley N° 29733 - Ley de Protección de Datos Personales."
    
  
    sParrafoDos = "Declaro conocer el compromiso de " & String(0.52, vbTab) & "LA CAJA " & String(0.52, vbTab) & " por garantizar el mantenimiento de la confidencialidad" & String(0.52, vbTab) & " y " & String(0.52, vbTab) & "el " & Chr$(13) & _
                  "tratamiento seguro de mis datos personales, incluyendo el resguardo en las transferencias de " & String(0.52, vbTab) & "los mismos, " & Chr$(13) & _
                  "que se realicen " & String(0.53, vbTab) & "en cumplimiento de la " & String(0.55, vbTab) & " Ley N° 29733 - Ley de Protección " & String(0.53, vbTab) & " de Datos Personales. De" & String(0.53, vbTab) & "igual " & Chr$(13) & _
                  "manera, declaro " & String(0.52, vbTab) & "conocer que los datos personales " & String(0.55, vbTab) & "proporcionados por mi persona serán incorporados " & String(0.52, vbTab) & "al " & Chr$(13) & _
                  "Banco de Datos de Clientes de  " & String(0.6, vbTab) & " LA CAJA, el cual  " & String(0.55, vbTab) & "se encuentra debidamente registrado ante la" & String(0.52, vbTab) & " Dirección " & Chr$(13) & _
                  "Nacional  " & String(0.55, vbTab) & " de  " & String(0.55, vbTab) & " Protección de Datos " & String(0.55, vbTab) & "Personales, para lo cual " & String(0.55, vbTab) & " autorizo a LA CAJA " & String(0.52, vbTab) & "que " & String(0.55, vbTab) & " recopile, registre, " & Chr$(13) & _
                  "organice, " & String(0.55, vbTab) & "almacene, " & String(0.55, vbTab) & "conserve, bloquee, suprima, extraiga, consulte, utilice, transfiera, exporte, importe" & String(0.52, vbTab) & " o " & Chr$(13) & _
                  "procese de cualquier otra forma mis datos personales, con las limitaciones que prevé la Ley."
                 
                 
    sParrafoTres = "Del mismo modo, y siempre que así lo estime necesario, declaro conocer que podré ejercitar mis derechos " & Chr$(13) & _
                   "de " & String(0.55, vbTab) & " acceso, " & String(0.56, vbTab) & " rectificación, " & String(0.58, vbTab) & " cancelación " & String(0.55, vbTab) & " y " & String(0.55, vbTab) & " oposición relativos a este tratamiento, de conformidad " & String(0.52, vbTab) & "con lo " & Chr$(13) & _
                   "establecido" & String(0.51, vbTab) & " en " & String(0.5, vbTab) & "el " & String(0.6, vbTab) & " Titulo" & String(0.54, vbTab) & " III " & String(0.54, vbTab) & " de la Ley N° 29733 - Ley de Protección de Datos " & String(0.52, vbTab) & " Personales" & String(0.52, vbTab) & " acercándome " & Chr$(13) & _
                   "a cualquiera de las Agencias de LA CAJA a nivel nacional."

   sParrafoCuatro = "Asimismo, " & String(1.4, vbTab) & " declaro " & String(1.4, vbTab) & " conocer " & String(1.4, vbTab) & " el " & String(1.4, vbTab) & "compromiso " & String(1.4, vbTab) & " de " & String(1.4, vbTab) & " LA " & String(1.4, vbTab) & "CAJA " & String(1.4, vbTab) & " por " & String(1.4, vbTab) & "respetar " & String(1.4, vbTab) & "los " & String(1.4, vbTab) & "principios " & String(1.4, vbTab) & "de " & String(1.4, vbTab) & " legalidad, " & Chr$(13) & _
                    "consentimiento, finalidad, proporcionalidad, calidad, disposición de recurso, y nivel de protección adecuado," & Chr$(13) & _
                    "conforme lo dispone la Ley N° 29733 - Ley de Protección de Datos Personales," & String(1.4, vbTab) & " para " & String(1.4, vbTab) & "el " & String(1.4, vbTab) & "tratamiento de los" & Chr$(13) & _
                    "datos personales otorgados por mi persona."
                  
    sParrafoCinco = "Esta autorización es" & String(1.5, vbTab) & " indefinida y se mantendrá inclusive" & String(0.5, vbTab) & " después de terminada(s) la(s) operación(es)" & String(0.52, vbTab) & " y/o " & Chr$(13) & _
                    "el(los) Contrato(s) que tenga" & String(1.5, vbTab) & " o pueda tener con LA CAJA" & String(1.3, vbTab) & " sin perjuicio de " & String(0.5, vbTab) & "poder ejercer mis derechos " & String(0.52, vbTab) & "de " & Chr$(13) & _
                    "acceso, rectificación, cancelación y oposición mencionados en el presente documento."
                    
     Dim cfecha  As String 'pti1 add
     cfecha = Choose(Month(gdFecSis), "Enero", "Febrero", "Marzo", "Abril", _
                                        "Mayo", "Junio", "Julio", "Agosto", _
                                        "Setiembre", "Octubre", "Noviembre", "Diciembre")
    
    Dim nTamanio As Integer
    Dim Spac As Integer
    Dim Index As Integer
    Dim Princ As Integer
    Dim CantCarac As Integer
    Dim txtcDescrip As String
    Dim contador As Integer
    Dim nCentrar As Integer
    Dim nTamLet As Integer
    Dim spacvar As Integer
    
            nTamanio = Len(sParrafoUno)
            spacvar = 23
            Spac = 138
            Index = 1
            Princ = 1
            CantCarac = 0
            
            nTamLet = 6: contador = 0: nCentrar = 80
            
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoUno, Index, CantCarac)
                        oDoc.WTextBox Spac, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoUno, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoUno, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoUno, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoDos)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoDos, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoDos, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoDos, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoDos, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoTres)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoTres, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoTres, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoTres, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoTres, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            nTamanio = Len(sParrafoCuatro)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCuatro, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCuatro, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCuatro, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCuatro, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
            
            
            nTamanio = Len(sParrafoCinco)
            Spac = Spac + spacvar
            Index = 1
            Princ = 1
            CantCarac = 0
             nTamLet = 6: contador = 0: nCentrar = 80
                  Do While Index <= nTamanio And Spac <> 240
                    If InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) <> 0 Then
                        CantCarac = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare)
                        CantCarac = CantCarac - Index
                        txtcDescrip = Mid(sParrafoCinco, Index, CantCarac)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = InStr(Index, sParrafoCinco, Chr$(13), vbTextCompare) + 1
                        Spac = Spac + 5 + IIf((Len(txtcDescrip) / 50) > 1, ((Round(Len(txtcDescrip) / 50)) * 6) - 4, 0)
                        
                    ElseIf (Index <= nTamanio) And Index <> 1 Then
                        txtcDescrip = Mid(sParrafoCinco, Index, nTamanio)
                        oDoc.WTextBox Spac + contador, 35, 11, 520, txtcDescrip, "F1", 11
                        Index = nTamanio + 1
                    Else
                        oDoc.WTextBox Spac + contador, 35, 11, 520, sParrafoCinco, "F1", 11
                        Index = nTamanio + 1
                    End If
            Loop
    
                  'oDoc.WTextBox 125, 30, 360, 520, sParrafoUno, "F1", 11, hjustify
                  'oDoc.WTextBox 277, 30, 360, 520, sParrafoDos, "F1", 11, hjustify
                  'oDoc.WTextBox 376, 30, 360, 520, sParrafoTres, "F1", 11, hjustify
                  'oDoc.WTextBox 432, 30, 360, 520, sParrafoCuatro, "F1", 11, hjustify
                  'oDoc.WTextBox 484, 30, 360, 520, sParrafoCinco, "F1", 11, hjustify
    
     Dim oNPersona As New COMNPersona.NCOMPersona
     Dim sCiudadAgencia As String
     sCiudadAgencia = oNPersona.ObtenerDistritoAgencia(gsCodAge)

    oDoc.WTextBox 610, 35, 60, 520, ("En " & sCiudadAgencia & " a los " & Day(gdFecSis) & " días del mes de " & cfecha & " de " & Year(gdFecSis)) & ".", "F1", 11, hLeft 'O  agregado  por pti1
    oDoc.WTextBox 670, 35, 90, 200, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 730, 35, 60, 180, "________________________________________", "F1", 8, hCenter
    oDoc.WTextBox 745, 90, 60, 80, "Firma", "F1", 10, hCenter
    
    sParrafoSeis = "¿Autorizas a Caja Maynas para el tratamiento de sus datos personales?"
    
    oDoc.WTextBox 670, 280, 60, 250, sParrafoSeis, "F1", 11, hLeft 'O  agregado  por pti1
   
   
    oDoc.WTextBox 712, 300, 15, 20, "SI", "F1", 8, hCenter
    oDoc.WTextBox 742, 300, 15, 20, "NO", "F1", 8, hCenter
    
    oDoc.WTextBox 690, 420, 70, 80, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    If Trim(Right(CboAutoriazaUsoDatos.Text, 2)) Then
        oDoc.WTextBox 710, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 710, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
    Else
        oDoc.WTextBox 740, 280, 15, 20, "", "F1", 8, hLeft, vMiddle, vbBlack, 1, vbBlack, , 3
        oDoc.WTextBox 740, 280, 15, 20, "X", "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, , 3
    End If
    

            
    oDoc.PDFClose
    oDoc.Show
    '</body>
End Sub
'add pti1 ers070-2018 26/12/2018 *********************
Private Sub ValidaAutorizardatos()
    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    bPemisoAD = oGen.VerificaExistePermisoCargo(gsCodCargo, PermisoCargos.gEdicionPersonas, gsCodPersUser)

    If bPemisoAD Then
     lblAutorizarUsoDatos.Visible = True
     CboAutoriazaUsoDatos.Visible = True
    Else
        If nTipoForm <> 1 Then
          lblAutorizarUsoDatos.Visible = True
          CboAutoriazaUsoDatos.Visible = True
        Else
           lblAutorizarUsoDatos.Visible = False
          CboAutoriazaUsoDatos.Visible = False
        End If
        
   
    End If
End Sub
'fin ADD PTI1 ERS070-2018*******************************
