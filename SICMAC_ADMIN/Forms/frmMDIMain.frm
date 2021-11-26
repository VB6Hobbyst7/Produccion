VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmMdiMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   4485
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7080
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4230
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2999
            MinWidth        =   2999
            TextSave        =   "10:04"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoPrint 
         Caption         =   "Configurar &Impresora"
      End
      Begin VB.Menu mnuArchivoLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivoSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuCaja 
      Caption         =   "&Caja"
      Begin VB.Menu mnuCajaMov 
         Caption         =   "&Movimientos"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuCajaRep 
         Caption         =   "&Reportes"
      End
      Begin VB.Menu mnuCajaLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCajaTransCmac 
         Caption         =   "Transferencia de &CMACs"
      End
      Begin VB.Menu mnuCajaTransCheque 
         Caption         =   "Transferencia de C&heques"
      End
      Begin VB.Menu mnuCajaLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCajaConsolidaEstadistica 
         Caption         =   "Consolidación de &Estadísticas"
      End
      Begin VB.Menu mnuCajaConsolidaBillete 
         Caption         =   "Consolidación de &Billetaje"
      End
      Begin VB.Menu mnuCajaLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCajaCierre 
         Caption         =   "&Cierre Mensual"
      End
   End
   Begin VB.Menu mnuPresup 
      Caption         =   "&Presupuesto"
      Begin VB.Menu mnuPresupMnt 
         Caption         =   "&Mantenimiento de Presupuesto"
      End
      Begin VB.Menu mnuPresupRubros 
         Caption         =   "Ingreso de &Rubros"
      End
      Begin VB.Menu mnuPresupPresup 
         Caption         =   "Ingreso de &Presupuesto"
      End
      Begin VB.Menu mnuPresupEjec 
         Caption         =   "&Ejecución de Presupuesto"
      End
      Begin VB.Menu mnuRaya 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBienes 
         Caption         =   "&Bienes"
         Begin VB.Menu mnuBienesG 
            Caption         =   "Soles"
            Index           =   0
         End
         Begin VB.Menu mnuBienesG 
            Caption         =   "Dolares"
            Index           =   1
         End
      End
      Begin VB.Menu mnuServiosn 
         Caption         =   "Servicios"
         Begin VB.Menu mnuServiosG 
            Caption         =   "Soles"
            Index           =   0
         End
         Begin VB.Menu mnuServiosG 
            Caption         =   "Dolares"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuRRHHG 
      Caption         =   "&Recursos Humanos"
      Begin VB.Menu mnuRRHHSel1 
         Caption         =   "&Selección"
         Begin VB.Menu mnuRRHHSel 
            Caption         =   "&Inicio Proceso Seleccion"
            Index           =   1
            Begin VB.Menu mnuRRHHSelR 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHSelR 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHSelR 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHSelR 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRRHHSel 
            Caption         =   "&Postulantes"
            Index           =   2
            Begin VB.Menu mnuRRHHSelPos 
               Caption         =   "&Resgistro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHSelPos 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHSelPos 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHSelPos 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRRHHSelEva1 
            Caption         =   "&Evaluación"
            Begin VB.Menu mnuRRHHSelEva 
               Caption         =   "&Curricular"
               Index           =   0
               Begin VB.Menu mnuRRHHSelEvaCur 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuRRHHSelEvaCur 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu mnuRRHHSelEvaCur 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu mnuRRHHSelEvaCur 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu mnuRRHHSelEva 
               Caption         =   "&Escrita"
               Index           =   1
               Begin VB.Menu mnuRRHHSelEvaEsc 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuRRHHSelEvaEsc 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu mnuRRHHSelEvaEsc 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu mnuRRHHSelEvaEsc 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu mnuRRHHSelEva 
               Caption         =   "&Psicologica"
               Index           =   2
               Begin VB.Menu mnuRRHHSelEvaPsi 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuRRHHSelEvaPsi 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu mnuRRHHSelEvaPsi 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu mnuRRHHSelEvaPsi 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
            Begin VB.Menu mnuRRHHSelEva 
               Caption         =   "E&ntrevista"
               Index           =   3
               Begin VB.Menu mnuRRHHSelEvaEnt 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuRRHHSelEvaEnt 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu mnuRRHHSelEvaEntCon 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
               Begin VB.Menu mnuRRHHSelEvaEntRep 
                  Caption         =   "&Reporte"
                  Index           =   3
               End
            End
         End
         Begin VB.Menu mnuRRHHSelRes 
            Caption         =   "&Resultados y Cierre"
         End
         Begin VB.Menu mnuRRHHSelCon 
            Caption         =   "&Consulta"
         End
      End
      Begin VB.Menu mnuRRHHCon 
         Caption         =   "&Contratos"
         Begin VB.Menu mnuRRHHConRSel 
            Caption         =   "&Registro basado en el Proceso de Selección"
         End
         Begin VB.Menu mnuRRHHConRMan 
            Caption         =   "R&egistro manual"
         End
         Begin VB.Menu mnuRRHHConMan 
            Caption         =   "&Mantenimiento"
         End
         Begin VB.Menu mnuRRHHFicPer 
            Caption         =   "&Consulta"
         End
         Begin VB.Menu mnuRRHHConRes 
            Caption         =   "Re&scindir "
         End
      End
      Begin VB.Menu mnuRRHHAdenda 
         Caption         =   "&Adenda"
      End
      Begin VB.Menu mnuRRHHCurG 
         Caption         =   "C&urriculum Vitae"
         Begin VB.Menu mnuRRHHCur 
            Caption         =   "&Configuracion"
            Index           =   0
            Begin VB.Menu mnuRRHHCurM 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHCurM 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu mnuRRHHCur 
            Caption         =   "&Registro"
            Index           =   1
         End
         Begin VB.Menu mnuRRHHCur 
            Caption         =   "&Mantenimiento"
            Index           =   2
         End
         Begin VB.Menu mnuRRHHCur 
            Caption         =   "&Consulta"
            Index           =   3
         End
         Begin VB.Menu mnuRRHHCur 
            Caption         =   "&Reporte"
            Index           =   4
         End
      End
      Begin VB.Menu mnuRRHHorLabT 
         Caption         =   "Ho&rario Laboral"
         Begin VB.Menu mnuRRHHorLab 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHorLab 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRRHAsistT 
         Caption         =   "&Asistencia"
         Begin VB.Menu mnuRRHAsist 
            Caption         =   "&Registro Automatico"
            Index           =   0
         End
         Begin VB.Menu mnuRRHAsist 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu mnuRRHAsist 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu mnuRRHHEvaIntG 
         Caption         =   "E&valuacione Interna"
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "&Inicio de proceso de Evaluacion Interna"
            Index           =   0
            Begin VB.Menu mnuRRHHEvaIntIni 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHEvaIntIni 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHEvaIntIni 
               Caption         =   "&Consulta"
               Index           =   2
            End
         End
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "&Curricular"
            Index           =   1
            Begin VB.Menu mnuRRHHEvaIntCur 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHEvaIntCur 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHEvaIntCur 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHEvaIntCur 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "&Escrita"
            Index           =   2
            Begin VB.Menu mnuRRHHEvaIntEscReg 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHEvaIntEsc 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHEvaIntEsc 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHEvaIntEsc 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "&Psicologica"
            Index           =   3
            Begin VB.Menu mnuRRHHEvaIntPsi 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHEvaIntPsi 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHEvaIntPsi 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHEvaIntPsi 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "E&ntrevista"
            Index           =   4
            Begin VB.Menu mnuRRHHEvaIntEnt 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHEvaIntEnt 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHEvaIntEnt 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHEvaIntEnt 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "Resultados y Cie&rre"
            Index           =   5
         End
         Begin VB.Menu mnuRRHHEvaInt 
            Caption         =   "&Consulta"
            Index           =   6
         End
      End
      Begin VB.Menu mnuRRHHPerDG 
         Caption         =   "&Permisos"
         Begin VB.Menu mnuRRHHPerD 
            Caption         =   "&Solicitud"
            Index           =   0
            Begin VB.Menu mnuRRHHPerSolD 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHPerSolD 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu mnuRRHHPerD 
            Caption         =   "&Autorizacion/Rechazo"
            Index           =   1
         End
         Begin VB.Menu mnuRRHHPerD 
            Caption         =   "&Reporte"
            Index           =   2
         End
      End
      Begin VB.Menu mnuRRHHVacD 
         Caption         =   "&Vacaciones"
         Begin VB.Menu mnuRRHHVac 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHVac 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu CD 
         Caption         =   "&Descansos"
         Begin VB.Menu mnuRRHHSan 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHSan 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRRHHMerDemG 
         Caption         =   "&Meritos y Demeritos"
         Begin VB.Menu mnuRRHHMerDem 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu mnuRRHHMerDemT 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHMerDemT 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu mnuRRHHMerDem 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu mnuRRHHMerDem 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "&Suspensiones"
         Index           =   9
         Begin VB.Menu mnuRRHHSanD 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHSanD 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "Car&gos Laborales"
         Index           =   10
         Begin VB.Menu mnuRRHHCar 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu mnuRRHHCarTab 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHCarTab 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu mnuRRHHCar 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "Sueldos"
         Index           =   13
         Begin VB.Menu mnuRRHHSueldo 
            Caption         =   "&Registro"
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "Sistema de Pensiones"
         Index           =   14
         Begin VB.Menu mnuRRHHSistPen 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu mnuRRHHSistPenTabla 
               Caption         =   "&Mantenimiento"
            End
         End
         Begin VB.Menu mnuRRHHSistPen 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "InformeSocial"
         Index           =   15
         Begin VB.Menu mnuRRHHInfSoc 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHInfSoc 
            Caption         =   "&Mantenimiento"
            Index           =   1
         End
         Begin VB.Menu mnuRRHHInfSoc 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu mnuRRHHInfSoc 
            Caption         =   "&Reportes"
            Index           =   3
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "Comentario"
         Index           =   16
         Begin VB.Menu mnuRRHHComentario 
            Caption         =   "&Registro"
         End
      End
      Begin VB.Menu mnuRRHH 
         Caption         =   "Asistencia Medica"
         Index           =   17
         Begin VB.Menu mnuRRHHAsistMed 
            Caption         =   "&Tabla"
            Index           =   0
            Begin VB.Menu mnuRRHHAsistMedTabla 
               Caption         =   "&Mantenimiento"
            End
         End
         Begin VB.Menu mnuRRHHAsistMed 
            Caption         =   "&Registro"
            Index           =   1
         End
      End
      Begin VB.Menu mnuRRHHRemu 
         Caption         =   "Confi&guracion de Planilla de Remuneraciones"
         Begin VB.Menu mnuRRHHRemuConRem 
            Caption         =   "&Configuracion de Conceptos Remunerativos"
            Begin VB.Menu mnuRRHHRemuConManRemu 
               Caption         =   "&Mantenimiento de Conceptos Remunerativos"
            End
            Begin VB.Menu mnuRRHHRemuConManTAlias 
               Caption         =   "&Mantenimiento de &Tabla Alias"
            End
            Begin VB.Menu mnuRRHHRemuConCon 
               Caption         =   "&Consulta"
            End
            Begin VB.Menu mnuRRHHRemuConRep 
               Caption         =   "&Reporte"
            End
         End
         Begin VB.Menu mnuRRHHRemuConPla 
            Caption         =   "C&onfiguracion de Conceptos de Planilla de Remuneraciones"
            Begin VB.Menu mnuRRHHRemuCon 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHRemuCon 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHRemuCon 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHRemuCon 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuRRHHPla 
         Caption         =   "&Planilla de remuneraciones del RR.HH."
         Begin VB.Menu mnuRRHHPlaConFij 
            Caption         =   "&Conceptos Fijos de Planilla de Remuneraciones"
            Begin VB.Menu mnuRRHHPlaConFijD 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHPlaConFijD 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHPlaConFijD 
               Caption         =   "&Reporte"
               Index           =   2
            End
         End
         Begin VB.Menu mnuRRHHPlaVar 
            Caption         =   "Conceptos &Variables de Planilla de Remuneraciones"
            Begin VB.Menu mnuRRHHPlaVarD 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHPlaVarD 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHPlaVarD 
               Caption         =   "&Reporte"
               Index           =   2
            End
         End
         Begin VB.Menu mnuRRHHPlaExt 
            Caption         =   "E&xtra Planilla Variables (Ingresos y Descuentos)"
            Begin VB.Menu mnuRRHHPlaExtD 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu mnuRRHHPlaExtD 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu mnuRRHHPlaExtD 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu mnuRRHHPlaExtD 
               Caption         =   "&Reporte"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuRRHHProT 
         Caption         =   "P&rocesos"
         Begin VB.Menu mnuRRHHPro 
            Caption         =   "&Calculo de Planillas"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHPro 
            Caption         =   "&Abono a Cuenta"
            Index           =   1
         End
         Begin VB.Menu mnuRRHHPro 
            Caption         =   "Cierre &Mensual"
            Index           =   2
         End
         Begin VB.Menu mnuRRHHPro 
            Caption         =   "Cierre &Diario"
            Index           =   3
         End
      End
      Begin VB.Menu mnuRRHHReportesG 
         Caption         =   "&Reportes"
         Begin VB.Menu mnuRRHHReportes 
            Caption         =   "&Reportes"
            Index           =   0
         End
         Begin VB.Menu mnuRRHHReportes 
            Caption         =   "Reportes &Generales"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuLogistica 
      Caption         =   "&Logística"
      Begin VB.Menu mnuLogUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuLogisticaProveeLinea0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogBienServicio 
         Caption         =   "&Bienes"
      End
      Begin VB.Menu mnuLogisticaProveeLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogProveedor 
         Caption         =   "&Proveedores"
      End
      Begin VB.Menu mnuLogisticaProveeLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogPlanAnual 
         Caption         =   "Requerimiento Regular"
         Begin VB.Menu mnuLogRequerimiento 
            Caption         =   "&Solicitud"
         End
         Begin VB.Menu mnuLogTramite 
            Caption         =   "&Trámite"
         End
         Begin VB.Menu mnuLogRegPreRef 
            Caption         =   "Precios Referenciales"
         End
         Begin VB.Menu mnuLogAprobacion 
            Caption         =   "Aprobación o Rechazo"
         End
         Begin VB.Menu mnuLogCtaCnt 
            Caption         =   "Asignación de Ctas.Presupuestrales"
         End
      End
      Begin VB.Menu mnuLogExtem 
         Caption         =   "Requerimiento &Extemporaneo"
         Begin VB.Menu mnuLogExtReq 
            Caption         =   "Solicitud"
         End
         Begin VB.Menu mnuLogExtTra 
            Caption         =   "Trámite"
         End
         Begin VB.Menu mnuLogExtPreRef 
            Caption         =   "Precio Referencial"
         End
         Begin VB.Menu mnuLogExtCtaCnt 
            Caption         =   "Asignación de Ctas.Presupuestrales"
         End
         Begin VB.Menu mnuLogExtApro 
            Caption         =   "Aprobacion o Rechazo"
         End
      End
      Begin VB.Menu mnuLogisticaProveeLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogAdqui 
         Caption         =   "Plan Anual de Adquisiciones y Contrataciones"
         Begin VB.Menu mnuLogAdquiConsol 
            Caption         =   "Consolidación de los Requerimientos al plan Anual de Adquisiciones "
         End
         Begin VB.Menu mnuLogAdquiConsul 
            Caption         =   "Consultas"
         End
      End
      Begin VB.Menu mnuLogisticaProveeLinea4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogSelec 
         Caption         =   "Proceso de Selección"
         Begin VB.Menu mnuLogDefPar 
            Caption         =   "Definición de Parámetros"
         End
         Begin VB.Menu mnuLogSelecIni 
            Caption         =   "Inicio de proceso"
            Begin VB.Menu mnuLogSelecIni1 
               Caption         =   "Resolución o acuerdo de inicio"
               Begin VB.Menu mnuLogSelIni 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelIni 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecIni2 
               Caption         =   "Conformación del Comite especial u organo encargado"
               Begin VB.Menu mnuLogSelCom 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelCom 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecIni3 
               Caption         =   "Detalle del proceso"
               Begin VB.Menu mnuLogSelBas 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelBas 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecIni4 
               Caption         =   "Bases o terminos de referencia"
               Begin VB.Menu mnuLogSelPar 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelPar 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuLogSelecCon 
            Caption         =   "Convocatoria"
            Begin VB.Menu mnuLogSelEcConP 
               Caption         =   "Publicación"
               Begin VB.Menu mnuLogSelPub 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelPub 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelEcConC 
               Caption         =   "Solicitudes de cotización"
               Begin VB.Menu mnuLogSelCot 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelCot 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuLogSelecBas 
            Caption         =   "Bases"
            Begin VB.Menu mnuLogSelecBase 
               Caption         =   "Registro de entrega de Bases"
               Begin VB.Menu mnuLogSelecBases 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecBases 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecConBas 
               Caption         =   "Consulta a Bases"
               Begin VB.Menu mnuLogSelecConsultaBase 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecConsultaBase 
                  Caption         =   "Cosulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecAbsolu 
               Caption         =   "Absolución de Consultas"
               Begin VB.Menu mnuLogSelecAbsolucion 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecAbsolucion 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu mnuLogSelecObserva 
               Caption         =   "Registro de Observaciones"
               Begin VB.Menu mnuLogSelecObsBases 
                  Caption         =   "Registro"
                  Index           =   0
               End
               Begin VB.Menu mnuLogSelecObsBases 
                  Caption         =   "Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu mnuLogSelecPro 
            Caption         =   "Propuestas"
            Begin VB.Menu mnuLogSelTec 
               Caption         =   "Técnica"
            End
            Begin VB.Menu mnuLogSelecO 
               Caption         =   "Económica"
            End
            Begin VB.Menu mnuLogSelGar 
               Caption         =   "Garantía de seriedad de oferta"
            End
         End
         Begin VB.Menu mnuLogSelecEva 
            Caption         =   "Evaluaciones"
         End
      End
      Begin VB.Menu mnuLogDesSelec 
         Caption         =   "Proceso de Selección Desierto"
      End
      Begin VB.Menu mnuLogCanSelec 
         Caption         =   "Cancelacion Proceso Seleccion"
      End
      Begin VB.Menu mnuLogAcepSelec 
         Caption         =   "Adjudicación del Proceso de Selección"
      End
      Begin VB.Menu mnuLogConsenAdju 
         Caption         =   "Consentimiento de Adjudicación"
      End
      Begin VB.Menu mnuLogConSelec 
         Caption         =   "Consultas"
      End
      Begin VB.Menu mnuSelLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogConComMantD 
         Caption         =   "Contratación"
         Index           =   0
         Begin VB.Menu mnuLogConComMantD1 
            Caption         =   "&Solicitud"
            Index           =   0
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu mnuLogConCom 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu mnuLogConComMantD1 
            Caption         =   "&Mantenimiento Contratación"
            Index           =   1
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu mnuLogConComMant 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
         Begin VB.Menu mnuLogConComMantD1 
            Caption         =   "&Impresión de Contratación"
            Index           =   2
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Compra Soles"
               Index           =   0
            End
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Compra Dolares"
               Index           =   1
            End
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Servicio Soles"
               Index           =   2
            End
            Begin VB.Menu mnuLogConComImp 
               Caption         =   "Orden Servicio Dolares"
               Index           =   3
            End
         End
      End
      Begin VB.Menu mnuLogProvision 
         Caption         =   "Provisiones pago de Proveedores"
         Begin VB.Menu mnuLogProvisionG 
            Caption         =   "Ordenes de Compra Soles"
            Index           =   0
         End
         Begin VB.Menu mnuLogProvisionG 
            Caption         =   "Ordenes de Servicios Soles"
            Index           =   1
         End
         Begin VB.Menu mnuLogProvisionG 
            Caption         =   "Ordenes de Compra Dolares"
            Index           =   2
         End
         Begin VB.Menu mnuLogProvisionG 
            Caption         =   "Ordenes de Servicios Dolares"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSelLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogAlmacen 
         Caption         =   "Almacén"
         Begin VB.Menu mnuLogAlmBieSer 
            Caption         =   "Operaciones"
         End
         Begin VB.Menu mnuLogAlmBieSerInven 
            Caption         =   "Inventario"
         End
         Begin VB.Menu mnuKardex 
            Caption         =   "Kardex"
         End
         Begin VB.Menu mnuRecSaldos 
            Caption         =   "&Recalculo de Saldos"
         End
      End
      Begin VB.Menu mnuLogDepreciacion 
         Caption         =   "Depreciación de Bienes"
      End
      Begin VB.Menu mnuLogRemate 
         Caption         =   "Remate de Bienes"
      End
      Begin VB.Menu mnuLogAdjudicacion 
         Caption         =   "Adjudicación de Bienes"
      End
      Begin VB.Menu mnuLogSubasta 
         Caption         =   "Subasta de Bienes"
      End
      Begin VB.Menu mnuSelLinea4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogServicio 
         Caption         =   "Servicios (locación, públicos, privados, móviles, otros)"
         Begin VB.Menu mnuLogSerServicios 
            Caption         =   "Registro"
            Index           =   0
         End
         Begin VB.Menu mnuLogSerServicios 
            Caption         =   "Distribución"
            Index           =   1
         End
         Begin VB.Menu mnuLogSerServicios 
            Caption         =   "Garantías"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmMdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Centralizacion CASL 11.12.2000

Private Sub MDIForm_Load()
    gdFecSis = Date
End Sub

Private Sub mnuBienesG_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501221", True
    Else
        frmLogOCAtencion.Inicio True, "502221", True
    End If
End Sub

Private Sub mnuCajaMov_Click()
    frmLogIngAlmacen.Ini "562403", "HOLA"
End Sub

Private Sub mnuKardex_Click()
    frmLogKardex.Show 1
End Sub

Private Sub mnuLogAdjudicacion_Click()
    frmLogOperacionesIngBS.Inicio "56"
End Sub

Private Sub mnuLogAlmBieSer_Click()
    frmLogOperacionesIngBS.Inicio "59"
End Sub

Private Sub mnuLogConCom_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCompra.Inicio False, "501205"
    ElseIf Index = 1 Then
        frmLogOCompra.Inicio False, "502205"
    ElseIf Index = 2 Then
        frmLogOCompra.Inicio False, "501207", "", False
    ElseIf Index = 3 Then
        frmLogOCompra.Inicio False, "502207", "", False
    End If
End Sub

Private Sub mnuLogConComImp_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501210", False, False, True
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio True, "502210", False, False, True
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio False, "501211", False, False, True
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "501211", False, False, True
    End If
End Sub

Private Sub mnuLogConComMant_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501206", False, True
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio True, "502206", False, True
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio False, "501208", False, True
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "501208", False, True
    End If
End Sub



Private Sub mnuLogProvisionG_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio True, "501221"
    ElseIf Index = 1 Then
        frmLogOCAtencion.Inicio False, "501222"
    ElseIf Index = 2 Then
        frmLogOCAtencion.Inicio True, "502221"
    ElseIf Index = 3 Then
        frmLogOCAtencion.Inicio False, "502222"
    End If
End Sub



Private Sub mnuLogSubasta_Click()
    frmLogSubastaVenta.Show 1
End Sub

Private Sub mnuPresupEjec_Click()
    frmPlaEjecu.Show 1
End Sub

Private Sub mnuPresupMnt_Click()
    frmPreMantenimiento.Show 1
End Sub

Private Sub mnuPresupPresup_Click()
    frmPlaPresu.Show 1
End Sub

Private Sub mnuPresupRubros_Click()
    frmPreRubros.Show 1
End Sub

Private Sub mnuRecSaldos_Click()
    frmLogCalculaSaldos.Show 1
End Sub

Private Sub mnuRRHAsist_Click(Index As Integer)
    If Index = 0 Then
    
    ElseIf Index = 1 Then
        frmRHAsistenciaManual.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:ASISTENCIA:MANUAL"
    ElseIf Index = 2 Then
        frmRHAsistenciaManual.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:ASISTENCIA:MANUAL"
    End If
End Sub

Private Sub mnuRRHHAdenda_Click()
    frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoAdenda, "RECURSOS HUMANOS:ADENDA"
End Sub

Private Sub mnuRRHHAsistMed_Click(Index As Integer)
    If Index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoAMP, "RECURSOS HUMANOS:ASISTENCIA MEDICA;REGISTRO"
    End If
End Sub

Private Sub mnuRRHHAsistMedTabla_Click()
    frmRHAsistMedPrivada.Ini "RECURSOS HUMANOS:ASITENCIA MEDICA:TABLA:MANTENIMIENTO"
End Sub

Private Sub mnuRRHHCar_Click(Index As Integer)
    If Index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoCargo, "RECURSOS HUMANOS:CARGOS LABORALES:REGISTRO"
    End If
End Sub

Private Sub mnuRRHHCarTab_Click(Index As Integer)
    If Index = 0 Then
        frmRHCargos.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CARGOS LABORALES:TABLA:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHCargos.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CARGOS LABORALES:TABLA:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHComentario_Click()
    frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoComentario, "RECURSOS HUMANOS:COMENTARIO:REGISTRO"
End Sub

Private Sub mnuRRHHConMan_Click()
    frmRHEmpleado.Ini gTipoOpeMantenimiento, RHContratoMantTpoFoto, "RECURSOS HUMANOS:CONTRATOS:MANTENIMIENTO"
End Sub

Private Sub mnuRRHHConRes_Click()
    Me.Enabled = False
    frmRHEmpleadoResCont.Ini "RECURSOS HUMANOS:CONTRATO:RESCINDIR CONTRATO", Me
    Me.Enabled = True
End Sub

Private Sub mnuRRHHConRMan_Click()
    frmRHContratoSeleccion.Ini "RECURSOS HUMANOS:CONTRATO:PROCESO SELECCION", ContratoFormaManual
End Sub

Private Sub mnuRRHHConRSel_Click()
    'frmRHEvaluacionNotas.Ini gTipoOpeConsulta, RHTipoOpeEvaConsolidado, True
    frmRHContratoSeleccion.Ini "RECURSOS HUMANOS:CONTRATO:PROCESO SELECCION", ContratoFormaAutomatica
End Sub

Private Sub mnuRRHHCur_Click(Index As Integer)
    If Index = 1 Then
        frmRHCurriculum.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:CURRICULUM VITAE:REGISTRO"
    ElseIf Index = 2 Then
        frmRHCurriculum.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CURRICULUM VITAE:MANTENIMIENTO"
    ElseIf Index = 3 Then
        frmRHCurriculum.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CURRICULUM VITAE:CONSULTA"
    ElseIf Index = 4 Then
        frmRHCurriculum.Ini gTipoOpeReporte, "RECURSOS HUMANOS:CURRICULUM VITAE:REPORTE"
    End If
End Sub

Private Sub mnuRRHHCurM_Click(Index As Integer)
    If Index = 0 Then
        frmRHCurriculumTabla.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:CURRICULUM VITAE:TABLA:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHCurriculumTabla.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:CURRICULUM VITAE:TABLA:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHEvaInt_Click(Index As Integer)
    If Index = 5 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:EVA INT:RESULTADOS Y CIERRE"
    ElseIf Index = 6 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:EVA INT:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHEvaIntCur_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaCur, "7RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:EVA INT:EVALUACION CURRICULAR:REPORTE"
    End If
End Sub

Private Sub mnuRRHHEvaIntEnt_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REPORTE"
    End If
End Sub

Private Sub mnuRRHHEvaIntEscReg_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:EVA INT:EVALUACION Escrita:REPORTE"
    End If
End Sub

Private Sub mnuRRHHEvaIntIni_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:EVA INT:EVALUACION:Proceso EVALUACION Interna:REPORTE"
    End If
End Sub

Private Sub mnuRRHHEvaIntPsi_Click(Index As Integer)
    If Index = 0 Then
        frmRHEvaluacion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REGISTRO"
    ElseIf Index = 1 Then
        frmRHEvaluacion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHEvaluacion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:CONSULTA"
    ElseIf Index = 3 Then
        frmRHEvaluacion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:EVA INT:EVALUACION Psicologica:REPORTE"
    End If
End Sub

Private Sub mnuRRHHFicPer_Click()
    frmRHEmpleado.Ini gTipoOpeConsulta, RHContratoMantTpoFoto, "RECURSOS HUMANOS:FICHA PERSONAL:CONSULTA"
End Sub

Private Sub mnuRRHHInfSoc_Click(Index As Integer)
    If Index = 0 Then
        frmRHInformeSocial.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:INFORME SOCIAL:REGISTRO"
    ElseIf Index = 1 Then
        frmRHInformeSocial.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:INFORME SOCIAL:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHInformeSocial.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:INFORME SOCIAL:CONSULTA"
    ElseIf Index = 3 Then
        frmRHInformeSocial.Ini gTipoOpeReporte, "RECURSOS HUMANOS:INFORME SOCIAL:REPORTE"
    End If
End Sub

Private Sub mnuRRHHMerDem_Click(Index As Integer)
    If Index = 1 Then
        frmRHMerDem.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHMerDem.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHMerDemT_Click(Index As Integer)
    If Index = 0 Then
        frmRHMerDemTabla.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHMerDemTabla.Ini gTipoOpeReporte, "RECURSOS HUMANOS:MERITOS Y DEMERITOS:TABLA:CONSULTA"
    End If
    
End Sub



Private Sub mnuRRHHorLab_Click(Index As Integer)
    If Index = 0 Then
        frmRHAsistenciaAsig.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:HORARIO LABORAL:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHAsistenciaAsig.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:HORARIO LABORAL:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHPerD_Click(Index As Integer)
    If Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoPermisosLicencias, "RECURSOS HUMANOS:PERMISOS:APROBACION/RECHAZO"
    End If
End Sub


Private Sub mnuRRHHPerSolD_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeRegistro, RHEstadosTpoPermisosLicencias, ""
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoPermisosLicencias, ""
    End If
End Sub

Private Sub mnuRRHHPlaConFijD_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeMantenimiento, True
    ElseIf Index = 1 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeConsulta, True
    ElseIf Index = 2 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeReporte, True
    End If
End Sub

Private Sub mnuRRHHPlaExtD_Click(Index As Integer)
    If Index = 0 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeRegistro, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:REGISTRO"
    ElseIf Index = 1 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeMantenimiento, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeConsulta, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:CONSULTA"
    ElseIf Index = 3 Then
        frmRHPlanillaMovExtra.Ini gTipoOpeReporte, "RECURSOS HUMANOS:PLANILLA DE REMUNERACIONES:EXTRA PLANILLA:REPORTE"
    End If
End Sub

Private Sub mnuRRHHPlaVarD_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeMantenimiento, False
    ElseIf Index = 1 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeConsulta, False
    ElseIf Index = 2 Then
        frmRHConceptoAsigRRHH.Ini gTipoOpeReporte, False
    End If
End Sub

Private Sub mnuRRHHPro_Click(Index As Integer)
    Me.Enabled = False
    If Index = 0 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHCalculo, "RECUROS HUMANOS:PROCESOS:CALCULO DE PLANILLAS", Me
    ElseIf Index = 1 Then
        frmRHPlanillas.Ini gTipoProcesoRRHHAbono, "RECUROS HUMANOS:PROCESOS:ABONO DE PLANILLAS", Me
    ElseIf Index = 2 Then
        frmRHCierreMes.Ini "RECUROS HUMANOS:PROCESOS:CIERRE MES"
    ElseIf Index = 3 Then
        frmRHCierreDia.Ini "RECUROS HUMANOS:PROCESOS:CIERRE DIA"
    End If
    Me.Enabled = True
End Sub

Private Sub mnuRRHHRemuCon_Click(Index As Integer)
    If Index = 0 Then
        frmRHConceptoAsigPla.Ini gTipoOpeRegistro, ""
    ElseIf Index = 1 Then
        frmRHConceptoAsigPla.Ini gTipoOpeMantenimiento, ""
    ElseIf Index = 2 Then
        frmRHConceptoAsigPla.Ini gTipoOpeConsulta, ""
    ElseIf Index = 3 Then
        frmRHConceptoAsigPla.Ini gTipoOpeReporte, ""
    End If
End Sub

Private Sub mnuRRHHRemuConCon_Click()
    frmRHConceptoMant.Ini gTipoOpeConsulta
End Sub

Private Sub mnuRRHHRemuConManRemu_Click()
    frmRHConceptoMant.Ini gTipoOpeMantenimiento
End Sub

Private Sub mnuRRHHRemuConManTAlias_Click()
    frmRHTablasAlias.Show 1
End Sub

Private Sub mnuRRHHRemuConRep_Click()
    frmRHConceptoMant.Ini gTipoOpeReporte
End Sub

Private Sub mnuRRHHReportes_Click(Index As Integer)
    Me.Enabled = False
    If Index = 0 Then
        frmRRHHRep.Ini "RECURSOS HUMANOS:REPORTES:REPORTES", Me
    ElseIf Index = 1 Then
        frmRRHHRepGen.Ini "RECURSOS HUMANOS:REPORTES:REPORTES GENERALES", Me
    End If
    Me.Enabled = True
End Sub

Private Sub mnuRRHHSan_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoSubsidiado, "RECURSOS HUMANOS:DESCANSOS:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoSubsidiado, "RECURSOS HUMANOS:DESCANSOS:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHSanD_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoSuspendido, "RECURSOS HUMANOS:SANCIONES:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoSuspendido, "RECURSOS HUMANOS:SANCIONES:CONSULTA"
    End If
End Sub

Private Sub mnuRRHHSelCon_Click()
    frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
End Sub

Private Sub mnuRRHHSelEvaCur_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaCur, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:CURRICULAR:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelEvaEnt_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ENTREVISTA:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelEvaEntRep_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:ENTREVISTA:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:ENTREVISTA:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:ENTREVISTA:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEnt, "RECURSOS HUMANOS:POSTULANTES:ENTREVISTA:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelEvaEsc_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaEsc, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:ESCRITO:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelEvaPsi_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuEvaPsi, "RECURSOS HUMANOS:POSTULANTES:EXAMEN:PSICOLOGICO:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelPos_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuPost, "RECURSOS HUMANOS:POSTULANTES:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelR_Click(Index As Integer)
    If Index = 0 Then
        frmRHSeleccion.Ini gTipoOpeRegistro, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:REGISTRO"
    ElseIf Index = 1 Then
        frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:MANTENIMIENTO"
    ElseIf Index = 2 Then
        frmRHSeleccion.Ini gTipoOpeConsulta, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:CONSULTA"
    ElseIf Index = 3 Then
        frmRHSeleccion.Ini gTipoOpeReporte, RHPSeleccionTpoMnuSel, "RECURSOS HUMANOS:PROCESO SELECCION:REPORTE"
    End If
End Sub

Private Sub mnuRRHHSelRes_Click()
    frmRHSeleccion.Ini gTipoOpeMantenimiento, RHPSeleccionTpoMnuResultado, "RECURSOS HUMANOS:PROCESO SELECCION:RESULTADOS Y CIERRE"
End Sub

Private Sub mnuRRHHSistPen_Click(Index As Integer)
    If Index = 1 Then
        frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoSisPens, "RECURSOS HUMANOS:SISTEMA PENSIONES:REGISTRO"
    End If
End Sub

Private Sub mnuRRHHSistPenTabla_Click()
    frmRHAFP.Ini "RECURSOS HUMANOS:SISTEMA PENSIONES:TABLA:MANTENIMIENTO"
End Sub

Private Sub mnuRRHHSueldo_Click()
    frmRHEmpleado.Ini gTipoOpeRegistro, RHContratoMantTpoSueldo, "RECURSOS HUMANOS:SUELDO:REGISTRO"
End Sub

Private Sub mnuRRHHVac_Click(Index As Integer)
    If Index = 0 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeMantenimiento, RHEstadosTpoVacaciones, "RECURSOS HUMANOS:VACACIONES:MANTENIMIENTO"
    ElseIf Index = 1 Then
        frmRHPeriodosNoLaborales.Ini gTipoOpeConsulta, RHEstadosTpoVacaciones, "RECURSOS HUMANOS:VACACIOBES:CONSULTA"
    End If
End Sub

Private Sub mnuLogAcepSelec_Click()
    frmLogSelConsol.Inicio "4"
End Sub

Private Sub mnuLogAdquiConsol_Click()
    frmLogAdqConsol.Show 1
End Sub

Private Sub mnuLogAdquiConsul_Click()
    frmLogAdqConsul.Show 1
End Sub

Private Sub mnuLogAlmBieSerInven_Click()
    frmLogAlmInven.Show 1
End Sub

Private Sub mnuLogAlmBieSerRecep_Click()
    'frmLogAlmRecep.Show 1
    frmLogOperacionesIngBS.Inicio "56"
    
End Sub

Private Sub mnuLogAlmBieSerTransfe_Click()
    frmLogAlmTransfe.Show 1
End Sub

Private Sub mnuLogAprobacion_Click()
    frmLogReqPrecio.Inicio "1", "3"
    'frmLogReqAproba.Inicio "1"
End Sub
Private Sub mnuLogBienServicio_Click()
    frmLogBieSerMant.Show 1
End Sub

Private Sub mnuLogCanSelec_Click()
    frmLogSelConsol.Inicio "2"
End Sub

Private Sub mnuLogConSelec_Click()
    frmLogSelConsol.Inicio "1"
End Sub

Private Sub mnuLogConsenAdju_Click()
    frmLogSelConsol.Inicio "5"
End Sub

Private Sub mnuLogCtaCnt_Click()
    frmLogReqPrecio.Inicio "1", "2"
    'frmLogReqCtaCont.Inicio "1"
End Sub

Private Sub mnuLogDefPar_Click()
    frmLogSelParaMant.Show 1
End Sub

Private Sub mnuLogDesSelec_Click()
    frmLogSelConsol.Inicio "3"
End Sub

Private Sub mnuLogExtApro_Click()
    frmLogReqPrecio.Inicio "2", "3"
    'frmLogReqAprobaExtem.Show 1
End Sub
Private Sub mnuLogExtCtaCnt_Click()
    frmLogReqPrecio.Inicio "2", "2"
    'frmLogReqCtaCont.Inicio "2"
End Sub

Private Sub mnuLogExtPreRef_Click()
    frmLogReqPrecio.Inicio "2", "1"
End Sub

Private Sub mnuLogExtReq_Click()
    frmLogReqInicio.Inicio "2", "1"
End Sub
Private Sub mnuLogExtTra_Click()
    frmLogReqTramite.Inicio "2"
End Sub
Private Sub mnuLogProveedor_Click()
    frmLogProvMant.Show 1
End Sub

Private Sub mnuLogRegPreRef_Click()
    frmLogReqPrecio.Inicio "1", "1"
End Sub

Private Sub mnuLogRequerimiento_Click()
    frmLogReqInicio.Inicio "1", "1"
End Sub

Private Sub mnuLogSelBas_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "3", "4"), IIf(Index = 0, False, True)
End Sub
Private Sub mnuLogSelCom_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "2", "3"), IIf(Index = 0, False, True)
End Sub
Private Sub mnuLogSelCot_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "6", "7"), IIf(Index = 0, False, True)
    'frmLogSelCotiza.Inicio "1"
End Sub

Private Sub mnuLogSelecAbsolucion_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "9", "10"), IIf(Index = 0, False, True)
End Sub

Private Sub mnuLogSelecBases_Click(Index As Integer)
    'frmLogSelEntBase.Inicio "1"
    frmLogSelInicio.Inicio IIf(Index = 0, "7", "8"), IIf(Index = 0, False, True)
End Sub

Private Sub mnuLogSelecConsultaBase_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "8", "9"), IIf(Index = 0, False, True)
End Sub

Private Sub mnuLogSelecEva_Click()
    frmLogSelCotPro.Inicio "2", "0"
End Sub

Private Sub mnuLogSelEco_Click()
    frmLogSelCotPro.Inicio "1", "2"
End Sub

Private Sub mnuLogSelecObsBases_Click(Index As Integer)
    'frmLogSelEntBase.Inicio "2"
    frmLogSelInicio.Inicio IIf(Index = 0, "10", "11"), IIf(Index = 0, False, True)
End Sub

Private Sub mnuLogSelGar_Click()
    frmLogSelCotPro.Inicio "1", "3"
End Sub

Private Sub mnuLogSelIni_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "1", "2"), IIf(Index = 0, False, True)
End Sub
Private Sub mnuLogSelPar_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "4", "5"), IIf(Index = 0, False, True)
End Sub
Private Sub mnuLogSelPub_Click(Index As Integer)
    frmLogSelInicio.Inicio IIf(Index = 0, "5", "6"), IIf(Index = 0, False, True)
End Sub

Private Sub mnuLogSelTec_Click()
    frmLogSelCotPro.Inicio "1", "1"
End Sub

Private Sub mnuLogSerServicios_Click(Index As Integer)
    '1 - REGISTRO ;     2 - DISTRIBUCION;     3 - GARANTIA
    frmLogSerCon.Inicio (Index + 1)
End Sub

Private Sub mnuLogTramite_Click()
    frmLogReqTramite.Inicio "1"
End Sub
Private Sub mnuLogUsuarios_Click()
    frmLogUsuario.Show 1
End Sub

Private Sub mnuServiosG_Click(Index As Integer)
    If Index = 0 Then
        frmLogOCAtencion.Inicio False, "501222", True
    Else
        frmLogOCAtencion.Inicio False, "502222", True
    End If
End Sub
