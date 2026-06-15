using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCob_ResumenDeComisionesMensuales.
	/// </summary>
	public class frmCob_ResumenDeComisionesMensuales : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblperiodo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtPeriodo;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridComisiones;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCargos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCargos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnCuotaCero;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCob_ResumenDeComisionesMensuales()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCob_ResumenDeComisionesMensuales));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyectado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Cobrar");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Proyectado", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Proyectado", 2, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOperador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Retirado_Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Retirado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Rescate");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_90_120");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_90_120_Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_90_120_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_90_120_Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje_Bono_Trama");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono_Plata");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje_Penalizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Penalizacion_Dinero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Penalizacion_Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono_Total_MEs");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje_Porcentaje_Trama");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retirado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentajes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CobranzaExterna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CarteraEnviada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CarteraATrabajar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_Dias_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado_Dias_Porc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono_Penalizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Penalizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idpersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Retirado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("POrcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bonificacion");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTramo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tramo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTramo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tramo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyectado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cobrar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.lblperiodo = new System.Windows.Forms.Label();
			this.dtPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridComisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridCargos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkCargos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnCuotaCero = new System.Windows.Forms.Button();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.dtPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			this.SuspendLayout();
			// 
			// lblperiodo
			// 
			this.lblperiodo.AutoSize = true;
			this.lblperiodo.BackColor = System.Drawing.Color.Transparent;
			this.lblperiodo.Location = new System.Drawing.Point(8, 10);
			this.lblperiodo.Name = "lblperiodo";
			this.lblperiodo.Size = new System.Drawing.Size(43, 16);
			this.lblperiodo.TabIndex = 731;
			this.lblperiodo.Text = "Periodo";
			this.lblperiodo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtPeriodo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtPeriodo.Appearance = appearance1;
			this.dtPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtPeriodo.DateButtons.Add(dateButton1);
			this.dtPeriodo.Format = "MM/yyyy";
			this.dtPeriodo.Location = new System.Drawing.Point(56, 8);
			this.dtPeriodo.Name = "dtPeriodo";
			this.dtPeriodo.NonAutoSizeHeight = 23;
			this.dtPeriodo.Size = new System.Drawing.Size(104, 21);
			this.dtPeriodo.SpinButtonsVisible = true;
			this.dtPeriodo.TabIndex = 730;
			this.dtPeriodo.Value = ((object)(resources.GetObject("dtPeriodo.Value")));
			this.dtPeriodo.ValueChanged += new System.EventHandler(this.dtPeriodo_ValueChanged);
			// 
			// uGridComisiones
			// 
			this.uGridComisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridComisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridComisiones.DataSource = this.ultraDataSource7;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridComisiones.DisplayLayout.Appearance = appearance2;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Operador/ Cobrador/ Supervisor";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 250;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 3;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance6;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Total Comisión";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn7.Header.Caption = "Tramo";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridComisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridComisiones.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridComisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridComisiones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridComisiones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridComisiones.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 7.75F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridComisiones.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.uGridComisiones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridComisiones.DisplayLayout.Override.RowAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComisiones.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			this.uGridComisiones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridComisiones.Location = new System.Drawing.Point(312, 56);
			this.uGridComisiones.Name = "uGridComisiones";
			this.uGridComisiones.Size = new System.Drawing.Size(920, 376);
			this.uGridComisiones.TabIndex = 737;
			this.uGridComisiones.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridComisiones_InitializeRow);
			this.uGridComisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridComisiones_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn43.DataType = typeof(System.DateTime);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1296, 8);
			this.groupBox1.TabIndex = 739;
			this.groupBox1.TabStop = false;
			// 
			// uGridCargos
			// 
			this.uGridCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCargos.DataSource = this.ultraDataSource4;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridCargos.DisplayLayout.Appearance = appearance13;
			this.uGridCargos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 73;
			ultraGridColumn9.Header.Caption = "Operador/ Cobrador/ Supervisor";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 245;
			ultraGridColumn10.Header.Caption = "...";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Header.Caption = "Personal";
			ultraGridBand2.HeaderVisible = true;
			this.uGridCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCargos.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.uGridCargos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCargos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCargos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.uGridCargos.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8.5F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCargos.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.uGridCargos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCargos.Location = new System.Drawing.Point(8, 232);
			this.uGridCargos.Name = "uGridCargos";
			this.uGridCargos.Size = new System.Drawing.Size(296, 200);
			this.uGridCargos.TabIndex = 744;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn56.DefaultValue = 0;
			ultraDataColumn58.DataType = typeof(bool);
			ultraDataColumn58.DefaultValue = false;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58});
			// 
			// chkCargos
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCargos.Appearance = appearance20;
			this.chkCargos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCargos.Checked = true;
			this.chkCargos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCargos.Location = new System.Drawing.Point(280, 256);
			this.chkCargos.Name = "chkCargos";
			this.chkCargos.Size = new System.Drawing.Size(13, 13);
			this.chkCargos.TabIndex = 746;
			this.chkCargos.CheckedChanged += new System.EventHandler(this.chkCargos_CheckedChanged);
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn59,
																																 ultraDataColumn60});
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(176, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 747;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1160, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 957;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(744, 9);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 958;
			// 
			// btnCuotaCero
			// 
			this.btnCuotaCero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCuotaCero.CausesValidation = false;
			this.btnCuotaCero.Image = ((System.Drawing.Image)(resources.GetObject("btnCuotaCero.Image")));
			this.btnCuotaCero.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCuotaCero.Location = new System.Drawing.Point(1072, 8);
			this.btnCuotaCero.Name = "btnCuotaCero";
			this.btnCuotaCero.Size = new System.Drawing.Size(80, 23);
			this.btnCuotaCero.TabIndex = 959;
			this.btnCuotaCero.Text = "Cuota Cero";
			this.btnCuotaCero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCuotaCero.Visible = false;
			this.btnCuotaCero.Click += new System.EventHandler(this.btnCuotaCero_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource6;
			appearance21.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance21;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 87;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 221;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 54;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand3.Header.Caption = "Cargo";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 8.5F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance24;
			appearance25.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance25.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(296, 168);
			this.ultraGrid1.TabIndex = 960;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn63.DataType = typeof(bool);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63});
			// 
			// ultraCheckEditor1
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			this.ultraCheckEditor1.Appearance = appearance28;
			this.ultraCheckEditor1.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ultraCheckEditor1.Checked = true;
			this.ultraCheckEditor1.CheckState = System.Windows.Forms.CheckState.Checked;
			this.ultraCheckEditor1.Location = new System.Drawing.Point(280, 64);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(13, 13);
			this.ultraCheckEditor1.TabIndex = 961;
			this.ultraCheckEditor1.CheckedChanged += new System.EventHandler(this.ultraCheckEditor1_CheckedChanged);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn69.DataType = typeof(System.Decimal);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70});
			// 
			// frmCob_ResumenDeComisionesMensuales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1240, 438);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnCuotaCero);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.chkCargos);
			this.Controls.Add(this.uGridCargos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridComisiones);
			this.Controls.Add(this.lblperiodo);
			this.Controls.Add(this.dtPeriodo);
			this.KeyPreview = true;
			this.Name = "frmCob_ResumenDeComisionesMensuales";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Resumen de Comisiones Cobranzas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCob_ResumenDeComisionesMensuales_KeyDown_1);
			this.Load += new System.EventHandler(this.frmCob_ResumenDeComisionesMensuales_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
   int iGrupo = 0;
		private void frmCob_ResumenDeComisionesMensuales_Load(object sender, System.EventArgs e)
		{
			
		
			miAcceso = new Acceso(cdsSeteoF, "083606");

			if (!Funcion.Permiso("1212", cdsSeteoF))
			{				
				MessageBox.Show("No puede tiene acceso al reporte selecionado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			//if (miAcceso.Buscar) this.btnCuotaCero.Visible = true;
			if (miAcceso.BExportar) this.btnExcel.Visible = true;

			this.dtPeriodo.Value = DateTime.Today;
			
		  int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
	  	int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
			string ssqlDo = string.Format ("Exec ConsultaCbo_OperadorCobrador {0}, 1", MenuLatinium.IdUsuario ); 
			SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
			drDO.Read();
			if(drDO.HasRows)
			{
				iGrupo = drDO.GetInt32 (0);
			}
			drDO.Close();

     bool bFiltro = true;

//		 bFiltro = Funcion.bEscalarSQL(cdsSeteoF, string.Format(" Exec Cob_ResumenComisionesListaCargos {0},'{1}', {2}",
//			                                   MenuLatinium.IdUsuario, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ,1),true);	
//
//			if (iGrupo == 19 )
//			{
//				this.optFiltros.Value = 1;
//			}
//			if (iGrupo == 31 )
//			{
//				this.optFiltros.Value = 2;
//			}
//			if (iGrupo != 31 || iGrupo != 19 )
//			{
//				this.optFiltros.Value = 0;
//			}
//			if (bFiltro )
//			{
//				this.optFiltros.Visible = true;
//			}
//				 
			Cargos ();
//			if ( this.uGridCargos.Rows.Count > 1)
//			{
//				this.uGridCargos.Visible = true;
//			  this.chkCargos.Visible = true;
//				this.uGridComisiones.Location = new Point(312, 64);				
//				this.uGridComisiones.Size = new Size(1100, 435);
//			}
//			else
//			{
//				this.uGridCargos.Visible = false;
//				this.uGridComisiones.Location = new Point(8, 64);				
//				this.uGridComisiones.Size = new Size(1430, 435);
//				this.chkCargos.Visible = false;
//				this.Size =  new Size(	1632, 472);
//
//			}

//			this.optFiltros.Focus ();

		}
		private void Cargos()
		{
			int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				

			

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaCbo_OperadorCobradorComisiones {0}, {1},'{2}'", 
				MenuLatinium.IdUsuario,iGrupo,Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ));

			this.uGridCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ValidaOperadorcobrador {0}, '{1}'", 
				MenuLatinium.IdUsuario,Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ));
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtPeriodo, "Seleccione el Periodo", false, "", cdsSeteoF)) return;

			int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
			this.lblEstado.Text="";

			string sTramo= "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sTramo.Length == 0) sTramo = dr.Cells["idTramo"].Value.ToString();
					else if (sTramo.Length > 0) sTramo = sTramo + ", " + dr.Cells["idTramo"].Value.ToString();
				}
			}

			if (sTramo.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Cargo ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sCargos = "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCargos.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sCargos.Length == 0) sCargos = dr.Cells["idCargo"].Value.ToString();
					else if (sCargos.Length > 0) sCargos = sCargos + ", " + dr.Cells["idCargo"].Value.ToString();
				}
			}

			if (sCargos.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Personal ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sSQLPa = string.Format("Exec [Cob_ResumenDeComisionesMensuales2] '{0}', '{1}', '{2}'", 
				sCargos, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"), sTramo);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridComisiones);
			this.lblEstado.Text = this.uGridComisiones.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void chkCargos_CheckedChanged(object sender, System.EventArgs e)
		{
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCargos.Rows.All)
				dr.Cells["Sel"].Value = this.chkCargos.Checked;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void uGridComisiones_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
				if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) >= 100.00m) e.Row.Appearance.BackColor = Color.LimeGreen;
				else if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) < 100.00m) e.Row.Appearance.BackColor = Color.Yellow;
				else if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) < 85.00m) e.Row.Appearance.BackColor = Color.Orange;			
				else if (Convert.ToDecimal(e.Row.Cells["Porcentaje"].Value) < 70.00m) e.Row.Appearance.BackColor = Color.Red;		
		
		}

		private void optFiltros_ValueChanged(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtPeriodo, "Seleccione el Periodo", false, "", cdsSeteoF)) return;
			int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
//			this.uGridCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ParametroCOBOPE {0}, '{1}'",(int)this.optFiltros.Value, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));
		   this.btnVer_Click (sender, e);

//			this.Consulta ();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
						FuncionesProcedimientos.ExportaExcel(this.uGridComisiones);
		}

		private void Consulta ()
		{
//			if ( (int) this.optFiltros.Value == 0 || (int) this.optFiltros.Value == 2)
//			{
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Proyectado_Dias"].Hidden = false;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Proyectado_Dias_Cobrado"].Hidden = false;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Proyectado_Dias_Porc"].Hidden = false;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Bono_Penalizacion"].Hidden = false;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Penalizacion"].Hidden = false;
//			}
//
//			if ( (int) this.optFiltros.Value == 1)
//			{
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Proyectado_Dias"].Hidden = true;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Proyectado_Dias_Cobrado"].Hidden = true;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Proyectado_Dias_Porc"].Hidden = true;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Bono_Penalizacion"].Hidden = true;
//				this.uGridComisiones.DisplayLayout.Bands[0].Columns["Penalizacion"].Hidden = true;
//
//			}
		}

		private void uGridComisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtPeriodo, "Seleccione el Periodo", false, "", cdsSeteoF)) return;
			int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtPeriodo.Value).Year));

			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				
			this.uGridCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ValidaOperadorcobrador {0}, '{1}'", 
				MenuLatinium.IdUsuario,Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ));
//
//			this.uGridCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ParametroCOBOPE {0}, '{1}'",(int)this.optFiltros.Value, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));
//			this.btnVer_Click (sender, e);
		}

		private void btnCuotaCero_Click(object sender, System.EventArgs e)
		{
			using (frmReporteCuotaCero CCJ = new frmReporteCuotaCero ())
			{
				if (DialogResult.OK == CCJ.ShowDialog())
				{
						
				}
			}		
		}

		private void frmCob_ResumenDeComisionesMensuales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void ultraCheckEditor1_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				dr.Cells["Sel"].Value = this.ultraCheckEditor1.Checked;
		}

		private void frmCob_ResumenDeComisionesMensuales_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}
	}
}
